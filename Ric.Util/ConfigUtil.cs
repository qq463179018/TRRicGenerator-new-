using System;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace Ric.Util
{
    public class ConfigUtil
    {
        public static object ReadConfig(string configFilePath, Type configObjType)
        {
            using (XmlReader xr = XmlReader.Create(configFilePath))
            {
                XmlSerializer s = new XmlSerializer(configObjType);
                return s.Deserialize(xr);
            }
        }

        public static void WriteConfig(string configFilePath, object configObj)
        {
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "    ",
                Encoding = Encoding.UTF8,
                NewLineChars = "\r\n"
            };

            using (XmlWriter xw = XmlWriter.Create(configFilePath, settings))
            {
                XmlSerializer s = new XmlSerializer(configObj.GetType());
                s.Serialize(xw, configObj);
            }
        }

        public static void WriteXml(string xmlFilePath, object xmlObject)
        {
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "    ",
                Encoding = Encoding.Unicode,
                NewLineChars = "\r\n"
            };

            using (XmlWriter xw = XmlWriter.Create(xmlFilePath, settings))
            {
                XmlSerializer s = new XmlSerializer(xmlObject.GetType());
                s.Serialize(xw, xmlObject);
            }
        }
    }
}