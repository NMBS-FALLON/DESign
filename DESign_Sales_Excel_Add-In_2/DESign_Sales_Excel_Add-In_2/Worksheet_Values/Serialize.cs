using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace DESign_Sales_Excel_Add_In_2.Serialize
{
  internal class Serialize
  {
    /// <summary>
    ///   Serializes an object.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="serializableObject"></param>
    /// <param name="fileName"></param>
    public void SerializeObject<T>(T serializableObject, string fileName)
    {
      if (serializableObject == null) return;

      try
      {
        var xmlDocument = new XmlDocument();
        var serializer = new XmlSerializer(serializableObject.GetType());
        using (var stream = new MemoryStream())
        {
          serializer.Serialize(stream, serializableObject);
          stream.Position = 0;
          xmlDocument.Load(stream);
          xmlDocument.Save(fileName);
          stream.Close();
        }
      }
      catch (Exception ex)
      {
        //Log exception here
        ex.ToString();
      }
    }


    /// <summary>
    ///   Deserializes an xml file into an object list
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public T DeSerializeObject<T>(string fileName)
    {
      if (string.IsNullOrEmpty(fileName)) return default(T);

      var objectOut = default(T);

      try
      {
        var xmlDocument = new XmlDocument();
        xmlDocument.Load(fileName);
        var xmlString = xmlDocument.OuterXml;

        using (var read = new StringReader(xmlString))
        {
          var outType = typeof(T);

          var serializer = new XmlSerializer(outType);
          using (XmlReader reader = new XmlTextReader(read))
          {
            objectOut = (T) serializer.Deserialize(reader);
            reader.Close();
          }

          read.Close();
        }
      }
      catch (Exception ex)
      {
        //Log exception here
        ex.ToString();
      }

      return objectOut;
    }
  }
}