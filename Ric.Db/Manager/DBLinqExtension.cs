using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Data;
using System.Data.Linq.Mapping;
using System.Linq.Expressions;
using System.Data.Common;
using System.Diagnostics;

namespace Ric.Db.Manager
{
    /// <summary> 
    /// Compares arrays of objects using the supplied comparer (or default is none supplied) 
    /// </summary> 
    class ArrayComparer<T> : IEqualityComparer<T[]>
    {
        private readonly IEqualityComparer<T> comparer;

        public ArrayComparer() : this(null) { }
        public ArrayComparer(IEqualityComparer<T> comparer)
        {
            this.comparer = comparer ?? EqualityComparer<T>.Default;
        }

        public int GetHashCode(T[] values)
        {
            if (values == null) return 0;
            int hashCode = 1;
            for (int i = 0; i < values.Length; i++)
            {
                hashCode = (hashCode * 13) + comparer.GetHashCode(values[i]);
            }
            return hashCode;
        }
        public bool Equals(T[] lhs, T[] rhs)
        {
            if (ReferenceEquals(lhs, rhs)) return true;
            if (lhs == null || rhs == null || lhs.Length != rhs.Length)
                return false;
            for (int i = 0; i < lhs.Length; i++)
            {
                if (!comparer.Equals(lhs[i], rhs[i])) return false;
            }
            return true;
        }
    }

    /// <summary> 
    /// Represents a single bindable member of a type 
    /// </summary> 
    internal class BindingInfo
    {
        public bool CanBeNull { get; private set; }
        public MemberInfo StorageMember { get; private set; }
        public MemberInfo BindingMember { get; private set; }

        public BindingInfo(bool canBeNull, MemberInfo bindingMember, MemberInfo storageMember)
        {
            CanBeNull = canBeNull;
            BindingMember = bindingMember;
            StorageMember = storageMember;
        }

        public Type StorageType
        {
            get
            {
                switch (StorageMember.MemberType)
                {
                    case MemberTypes.Field:
                        return ((FieldInfo)StorageMember).FieldType;
                    case MemberTypes.Property:
                        return ((PropertyInfo)StorageMember).PropertyType;
                    default:
                        throw new NotSupportedException(string.Format("Unexpected member-type: {0}", StorageMember.Name));
                }
            }
        }
    }

    /// <summary> 
    /// Responsible for creating and caching reader-delegates for compatible 
    /// column sets; thread safe. 
    /// </summary> 
    static class InitializerCache<T>
    {
        private const BindingFlags FLAGS = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
        private const MemberTypes PROP_FIELD = MemberTypes.Property | MemberTypes.Field;

        /// <summary> 
        /// Cache of all readers for this T (by column sets) 
        /// </summary> 
        private static readonly Dictionary<string[], Func<IDataRecord, ExtendDataContext, T>> convertReaders =
            new Dictionary<string[], Func<IDataRecord, ExtendDataContext, T>>(
                new ArrayComparer<string>(StringComparer.InvariantCultureIgnoreCase)),
                vanillaReaders = new Dictionary<string[], Func<IDataRecord, ExtendDataContext, T>>(
                    new ArrayComparer<string>(StringComparer.InvariantCultureIgnoreCase));

        /// <summary> 
        /// Cache of all bindable columns for this T (by source-name) 
        /// </summary> 
        private static readonly SortedList<string, BindingInfo> dataMembers = new SortedList<string, BindingInfo>(StringComparer.InvariantCultureIgnoreCase);

        static InitializerCache()
        {
            Type type = typeof(T);
            foreach (MemberInfo member in type.GetMembers(FLAGS))
            {
                if ((member.MemberType & PROP_FIELD) == 0) continue; // only applies to prop/fields 
                ColumnAttribute col = Attribute.GetCustomAttribute(member, typeof(ColumnAttribute)) as ColumnAttribute;
                if (col == null) continue; // not a column 
                string name = col.Name;
                if (string.IsNullOrEmpty(name))
                { // default to self 
                    name = member.Name;
                }
                string storage = col.Storage;
                MemberInfo storageMember;
                if (string.IsNullOrEmpty(storage) || storage == name)
                { // default to self 
                    storageMember = member;
                }
                else
                {
                    // locate prop/field: case-sensitive first, then insensitive 
                    storageMember = GetBindingMember(storage);
                    if (storageMember == null)
                    {
                        throw new InvalidOperationException("Storage member not found: " + storage);
                    }
                }
                if (storageMember.MemberType == MemberTypes.Property && !((PropertyInfo)storageMember).CanWrite)
                { // write to a r/o prop? 
                    throw new InvalidOperationException("Cannot write to readonly storage property: " + storage);
                }
                // log it... 
                dataMembers.Add(name, new BindingInfo(col.CanBeNull, member, storageMember));
            }
        }

        private static bool TryGetBinding(string columnName, out BindingInfo binding)
        {
            return dataMembers.TryGetValue(columnName, out binding);
        }
        private static MemberInfo GetBindingMember(string name)
        {
            Type type = typeof(T);
            return FirstMember(type.GetMember(name, PROP_FIELD, FLAGS)) ?? FirstMember(type.GetMember(name, PROP_FIELD, FLAGS | BindingFlags.IgnoreCase));
        }
        private static MemberInfo FirstMember(MemberInfo[] members)
        {
            return members != null && members.Length > 0 ? members[0] : null;
        }
        private static Func<IDataRecord, ExtendDataContext, T> CreateInitializer(string[] names, bool useConversion)
        {
            Trace.WriteLine("Creating initializer for: " + typeof(T).Name);
            if (names == null) throw new ArgumentNullException("names");
            ParameterExpression readerParam = Expression.Parameter(typeof(IDataRecord), "record"),
                ctxParam = Expression.Parameter(typeof(ExtendDataContext), "ctx");
            Type entityType = typeof(T),
                underlyingEntityType = Nullable.GetUnderlyingType(entityType) ?? entityType,
                readerType = typeof(IDataRecord);
            List<MemberBinding> bindings = new List<MemberBinding>();
            Type[] byOrdinal = { typeof(int) };
            MethodInfo defaultMethod = readerType.GetMethod("GetValue", byOrdinal),
                isNullMethod = readerType.GetMethod("IsDBNull", byOrdinal),
                convertMethod = typeof(ExtendDataContext).GetMethod("OnConvertValue", BindingFlags.Instance | BindingFlags.NonPublic);
            NewExpression ctor = Expression.New(underlyingEntityType); // try this first... 
            for (int ordinal = 0; ordinal < names.Length; ordinal++)
            {
                string name = names[ordinal];
                BindingInfo bindingInfo;
                if (!TryGetBinding(name, out bindingInfo))
                { // try implicit binding 
                    MemberInfo member = GetBindingMember(name);
                    if (member == null) continue; // not bound 
                    bindingInfo = new BindingInfo(true, member, member);
                }
                //Trace.WriteLine(string.Format("Binding {0} to {1} ({2})", name, bindingInfo.Member.Name, bindingInfo.Member.MemberType)); 
                Type valueType = bindingInfo.StorageType;
                Type underlyingType = Nullable.GetUnderlyingType(valueType) ?? valueType;
                // get the rhs of a binding 
                MethodInfo method = readerType.GetMethod("Get" + underlyingType.Name, byOrdinal);
                Expression rhs;
                ConstantExpression ordinalExp = Expression.Constant(ordinal, typeof(int));
                if (method != null && method.ReturnType == underlyingType)
                {
                    rhs = Expression.Call(readerParam, method, ordinalExp);
                }
                else
                {
                    rhs = Expression.Convert(Expression.Call(readerParam, defaultMethod, ordinalExp), underlyingType);
                }
                if (underlyingType != valueType)
                {   // Nullable<T>; convert underlying T to T? 
                    rhs = Expression.Convert(rhs, valueType);
                }
                if (bindingInfo.CanBeNull && (underlyingType.IsClass || underlyingType != valueType))
                {
                    // reference-type of Nullable<T>; check for null 
                    // (conditional ternary operator) 
                    rhs = Expression.Condition(
                        Expression.Call(readerParam, isNullMethod, ordinalExp),
                        Expression.Constant(null, valueType), rhs);
                }
                if (useConversion)
                {
                    rhs = Expression.Convert(Expression.Call(ctxParam, convertMethod, ordinalExp, readerParam, Expression.Convert(rhs, typeof(object))), valueType);
                }
                bindings.Add(Expression.Bind(bindingInfo.StorageMember, rhs));
            }
            Expression body = Expression.MemberInit(ctor, bindings);
            if (entityType != underlyingEntityType)
            { // entity itself was T? - so convert 
                body = Expression.Convert(body, entityType);
            }
            return Expression.Lambda<Func<IDataRecord, ExtendDataContext, T>>(body, readerParam, ctxParam).Compile();
        }

        public static Func<IDataRecord, ExtendDataContext, T> GetInitializer(string[] names, bool useConversion)
        {
            if (names == null) throw new ArgumentNullException();
            Func<IDataRecord, ExtendDataContext, T> initializer;
            Dictionary<string[], Func<IDataRecord, ExtendDataContext, T>> cache = useConversion ? convertReaders : vanillaReaders;
            lock (cache)
            {
                if (!cache.TryGetValue(names, out initializer))
                {
                    initializer = CreateInitializer(names, useConversion);
                    cache.Add((string[])names.Clone(), initializer);
                }
            }
            return initializer;
        }
    }

    public class ValueConversionEventArgs : EventArgs
    {
        public int Ordinal { get; private set; }
        public object Value { get; set; }
        public IDataRecord Record { get; private set; }

        internal ValueConversionEventArgs() { }

        internal void Init(int ordinal, IDataRecord record, object value)
        {
            Ordinal = ordinal;
            Record = record;
            Value = value;
        }

        public ValueConversionEventArgs(int ordinal, IDataRecord record, object value)
        {
            Init(ordinal, record, value);
        }
    }

    public class ExtendDataContext
    {
        // re-use args to miniimze GEN0 
        private readonly ValueConversionEventArgs conversionArgs = new ValueConversionEventArgs();

        public event EventHandler<ValueConversionEventArgs> ConvertValue;

        public DbConnection Connection { get; set; }

        public ExtendDataContext(DbConnection dbConnection)
        {
            Connection = dbConnection;
        }

        internal object OnConvertValue(int ordinal, IDataRecord record, object value)
        {
            if (ConvertValue == null)
            {
                return value;
            }
            else
            {
                conversionArgs.Init(ordinal, record, value);
                ConvertValue(this, conversionArgs);
                return conversionArgs.Value;
            }
        }

        public IEnumerable<T> ExecuteQuery<T>(string command, params object[] parameters)
        {
            if (parameters == null) throw new ArgumentNullException("parameters");
            using (IDbConnection conn = Connection) // new SqlConnection(Program.CS)) 
            using (IDbCommand cmd = conn.CreateCommand())
            {
                string[] paramNames = new string[parameters.Length];
                for (int i = 0; i < parameters.Length; i++)
                {
                    paramNames[i] = "@p" + i.ToString();
                    IDbDataParameter param = cmd.CreateParameter();
                    param.ParameterName = paramNames[i];
                    param.Value = parameters[i] ?? DBNull.Value;
                    cmd.Parameters.Add(param);
                }
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = string.Format(command, paramNames);
                conn.Open();
                using (IDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult))
                {
                    if (reader.Read())
                    {
                        string[] names = new string[reader.FieldCount];
                        for (int i = 0; i < names.Length; i++)
                        {
                            names[i] = reader.GetName(i);
                        }
                        Func<IDataRecord, ExtendDataContext, T> objInit = InitializerCache<T>.GetInitializer(names, ConvertValue != null);
                        do
                        { // walk the data 
                            yield return objInit(reader, this);
                        } while (reader.Read());
                    }
                    while (reader.NextResult()) { } // ensure any trailing errors caught 
                }
            }
        }

        /// <summary>
        /// Execute query to database and return a collection of data entities.
        /// </summary>
        /// <typeparam name="T">The data entity type.</typeparam>
        /// <param name="command">The SQL command stirng.</param>
        /// <param name="timeout">Command timeout in seconds.</param>
        /// <param name="parameters">Command parameters if there is any.</param>
        /// <returns>A collection of records converted into data entities type.</returns>
        public IEnumerable<T> ExecuteQuery<T>(string command, int timeout, params object[] parameters)
        {
            if (parameters == null) throw new ArgumentNullException("parameters");
            using (IDbConnection conn = Connection) // new SqlConnection(Program.CS)) 
            using (IDbCommand cmd = conn.CreateCommand())
            {
                string[] paramNames = new string[parameters.Length];
                for (int i = 0; i < parameters.Length; i++)
                {
                    paramNames[i] = "@p" + i.ToString();
                    IDbDataParameter param = cmd.CreateParameter();
                    param.ParameterName = paramNames[i];
                    param.Value = parameters[i] ?? DBNull.Value;
                    cmd.Parameters.Add(param);
                }

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = string.Format(command, paramNames);
                cmd.CommandTimeout = timeout;
                conn.Open();

                using (IDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection | CommandBehavior.SingleResult))
                {
                    if (reader.Read())
                    {
                        string[] names = new string[reader.FieldCount];
                        for (int i = 0; i < names.Length; i++)
                        {
                            names[i] = reader.GetName(i);
                        }

                        Func<IDataRecord, ExtendDataContext, T> objInit = InitializerCache<T>.GetInitializer(names, ConvertValue != null);

                        do
                        { // walk the data 
                            yield return objInit(reader, this);
                        } while (reader.Read());
                    }

                    while (reader.NextResult()) { } // ensure any trailing errors caught 
                }
            }
        }
    }

    public static class Extensions
    {
        public static IEnumerable<object[]> ExecuteQuery(
            this DbLinq.Data.Linq.DataContext ctx, string query)
        {

            using (System.Data.Common.DbCommand cmd = ctx.Connection.CreateCommand())
            {
                cmd.CommandText = query;
                ctx.Connection.Open();
                using (System.Data.Common.DbDataReader rdr =
                    cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection))
                {
                    while (rdr.Read())
                    {
                        object[] res = new object[rdr.FieldCount];
                        rdr.GetValues(res);
                        yield return res;
                    }
                }
            }
        }

        public static IEnumerable<object[]> ExecuteQuery(
            this ExtendDataContext ctx, string query)
        {

            using (System.Data.Common.DbCommand cmd = ctx.Connection.CreateCommand())
            {
                cmd.CommandText = query;
                ctx.Connection.Open();
                using (System.Data.Common.DbDataReader rdr =
                    cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection))
                {
                    while (rdr.Read())
                    {
                        object[] res = new object[rdr.FieldCount];
                        rdr.GetValues(res);
                        yield return res;
                    }
                }
            }
        }
    }
}
