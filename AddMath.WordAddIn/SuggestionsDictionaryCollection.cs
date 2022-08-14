using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace AddMath.WordAddIn
{
    public class SuggestionsDictionary : Dictionary<string, string>, IXmlSerializable
    {
        public SuggestionsDictionary() : base() { }
        public SuggestionsDictionary(IDictionary<string, string> dictionary) : base(dictionary) { }

        public XmlSchema GetSchema() => null;

        public void ReadXml(XmlReader reader)
        {
            while (reader.Read() &&
            !(reader.NodeType == XmlNodeType.EndElement && reader.LocalName == nameof(SuggestionsDictionary)))
            {
                var key = reader["Key"];
                if (key == null)
                    throw new FormatException();

                var value = reader["Value"];
                this[key] = value;
            }
        }

        public void WriteXml(XmlWriter writer)
        {
            foreach (var entry in this)
            {
                writer.WriteStartElement("Pair");
                writer.WriteAttributeString("Key", entry.Key);
                writer.WriteAttributeString("Value", entry.Value);
                writer.WriteEndElement();
            }
        }
    }
    public class SuggestionsDictionaryCollection : Dictionary<string, SuggestionsDictionary>, IXmlSerializable
    {
        public XmlSchema GetSchema() => null;

        public void ReadXml(XmlReader reader)
        {
            while (reader.Read() &&
            !(reader.NodeType == XmlNodeType.EndElement && reader.LocalName == nameof(SuggestionsDictionaryCollection)))
            {
                var key = reader["Key"];
                if (key == null)
                    throw new FormatException();

                var value = new SuggestionsDictionary();
                if (reader.MoveToContent() == XmlNodeType.Element)
                {
                    reader.ReadStartElement();
                    if (reader.LocalName != nameof(SuggestionsDictionary))
                        throw new FormatException();
                    value.ReadXml(reader);
                    reader.ReadEndElement();
                }
                this[key] = value;
            }
        }

        public void WriteXml(XmlWriter writer)
        {
            foreach (var entry in this)
            {
                writer.WriteStartElement("Pair");
                writer.WriteAttributeString("Key", entry.Key);
                writer.WriteStartElement(nameof(SuggestionsDictionary));
                entry.Value.WriteXml(writer);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
        }
    }
}
