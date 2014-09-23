using System.Collections.Generic;
using System.Xml.Serialization;

namespace Ric.Core
{
    [XmlRoot("root")]
    public class ProductionXmlFileTemplate
    {
        public int channel { get; set; }
        public bool locale { get; set; }
        public Rics rics { get; set; }

        public ProductionXmlFileTemplate()
        {
            this.channel = 10;
            this.locale = false;
            this.rics = new Rics();
        }
    }

    public class Rics
    {
        [XmlElement(ElementName = "ric")]
        public List<Ric> rics { get; set; }

        public Rics()
        {
            rics = new List<Ric>();
        }
    }

    public class Ric
    {
        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }

        [XmlElement(ElementName = "fid")]
        public List<Fid> fids { get; set; }

        public Ric()
        {
            this.fids = new List<Fid>();
        }
    }

    public class Fid
    {
        [XmlAttribute(AttributeName = "id")]
        public int Id { get; set; }

        [XmlAttribute(AttributeName = "offset")]
        public int Offset { get; set; }

        [XmlAttribute(AttributeName = "locale")]
        public bool Locale { get; set; }

        [XmlText]
        public string Value { get; set; }

        public Fid()
        {
            Locale = false;
        }
    }
}