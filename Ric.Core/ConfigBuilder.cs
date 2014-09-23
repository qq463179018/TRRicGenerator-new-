using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Ric.Db.Manager;

namespace Ric.Core
{
    public class ConfigBuilder
    {
        private ConfigBuilder()
        {
            
        }

        public static bool IsConfigStoredInDB(Type type)
        {
            return type.GetCustomAttributes(typeof(ConfigStoredInDBAttribute), false).Length > 0;
        }

        public static bool IsPropertyStoredInDB(Type type, string propName)
        {
            return type.GetProperty(propName).GetCustomAttributes(typeof(StoreInDBAttribute), false).Length > 0;
        }

        public static List<string> GetPropertyNameWithValueInDB(Type type)
        {
            return (from p in type.GetProperties() let attrs = p.GetCustomAttributes(typeof (StoreInDBAttribute), false) where attrs.Length > 0 select p.Name).ToList();
        }

        public static Dictionary<string, string> GetPropertyValueFromDB(int userId, int taskId, List<string> propList)
        {
            return propList.ToDictionary(p => p, p => TaskConfigManager.GetConfigValue(userId, taskId, p, RunTimeContext.Context.DatabaseContext));
        }

        public static object CreateConfigInstance(Type type, int taskId)
        {
            List<string> propList = GetPropertyNameWithValueInDB(type);
            Dictionary<string, string> propDict = GetPropertyValueFromDB(RunTimeContext.Context.CurrentUser.Id, taskId, propList);
            
            object config = Activator.CreateInstance(type);

            foreach (string key in propDict.Keys)
            {
                SetPropertyValue(type, key, propDict[key], config);
            }
            return config;
        }

        public static void UpdateConfigProperty(Type type, object config, int userId, int taskId)
        {
            List<string> propList = GetPropertyNameWithValueInDB(type);
            foreach (string p in propList)
            {
                TaskConfigManager.UpdateConfig(userId, taskId, p, GetPropertyValue(type, p, config), RunTimeContext.Context.DatabaseContext);
            }
        }

        private static void SetPropertyValue(Type type, string property, string value, object obj)
        {
            if (string.IsNullOrEmpty(value))
            {
                DefaultValueAttribute attribute = TypeDescriptor.GetProperties(obj)[property].Attributes[typeof(DefaultValueAttribute)] as DefaultValueAttribute;
                if (attribute != null)
                {
                    value = attribute.Value.ToString();
                }
            }            
            string propertyTypeName = type.GetProperty(property).PropertyType.Name;
            object valueToSet = null;
            if (propertyTypeName.ToLower().Contains("list"))
            {
                valueToSet = value.Split("\r".ToCharArray()).ToList();
            }
            else if (propertyTypeName.ToLower().Contains("string"))
            {
                valueToSet = value;
            }
            else if (propertyTypeName.ToLower().Contains("int"))
            {
                try
                {
                    valueToSet = Int32.Parse(value);
                }
                catch
                {
                    valueToSet = 0;
                }
            }
            else if (propertyTypeName.ToLower().Contains("datetime"))
            {
                try
                {
                    valueToSet = DateTime.Parse(value);
                }
                catch
                {
                    valueToSet = DateTime.Now;
                }
            }
            else if (propertyTypeName.ToLower().Contains("bool"))
            {
                try
                {
                    valueToSet = Boolean.Parse(value);
                }
                catch
                {
                    valueToSet = false;
                }
            }

            type.GetProperty(property).SetValue(obj, valueToSet, null);
        }

        private static string GetPropertyValue(Type type, string property, object obj)
        {
            string propertyTypeName = type.GetProperty(property).PropertyType.Name;
            string result = null;

            object value = type.GetProperty(property).GetValue(obj, null);

            if (propertyTypeName.ToLower().Contains("list"))
            {
                result = string.Join("\r", (value as List<string>).ToArray());
            }
            //else if (propertyTypeName.ToLower().Contains("string"))
            else
            {
                result = value.ToString();
            }
            return result;
        }
    }

}
