using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace SSISWCFTask100
{
    class Serializer
    {
        public static string SerializeToXmlString(object objectToSerialize)
        {
            if (objectToSerialize == null)
                return string.Empty;

            byte[] byteArray = null;

            using (var memoryStream = new MemoryStream())
            {
                var ser = new XmlSerializer(objectToSerialize.GetType());
                ser.Serialize(memoryStream, objectToSerialize);
                byteArray = memoryStream.ToArray();
            }

            return new ASCIIEncoding().GetString(byteArray);
        }

        public static object DeSerializeFromXmlString(System.Type typeToDeserialize, string xmlString)
        {
            if (string.IsNullOrEmpty(xmlString))
                return new object();

            byte[] bytes = Encoding.UTF8.GetBytes(xmlString);
            object objectToDeserialize = null;

            using (MemoryStream memoryStream = new MemoryStream(bytes))
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeToDeserialize);
                objectToDeserialize = xmlSerializer.Deserialize(memoryStream);
            }

            return objectToDeserialize;
        }
    }
}
