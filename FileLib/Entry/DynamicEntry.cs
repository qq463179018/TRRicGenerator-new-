using System.Collections.Generic;
using System.Dynamic;

namespace Ric.FileLib.Entry
{
    public class DynamicEntry : DynamicObject
    {
        private readonly Dictionary<string, object> entry = new Dictionary<string, object>();

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            return entry.TryGetValue(binder.Name, out result);
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            entry[binder.Name] = value;
            return true;
        }

        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            if (binder.Name == "SetProperty")
            {
                if (entry.ContainsKey(args[0] as string))
                {
                    entry[args[0] as string] = args[1];
                }
                else
                {
                    entry.Add(args[0] as string, args[1]);
                }
            }
            result = null;
            return true;
        }
    }
}
