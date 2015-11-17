using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using System.Xml;

namespace com.github.yedijas.util
{
    class XMLUtil
    {
        #region static methods
        /// <summary>
        /// Serialize an object to XML string.
        /// Using generic to make sure the type is the same.
        /// </summary>
        /// <typeparam name="T">Type of the object.</typeparam>
        /// <param name="obj">Object to be serialized.</param>
        /// <returns>XML string.</returns>
        public static string Serialize<T>(object obj)
        {
            StringWriter writer = new StringWriter();
            if (obj.GetType() == typeof(T))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(T));
                using (writer)
                {
                    serializer.Serialize(writer, obj);
                }
            }
            else
            {
                throw new ArgumentException("XML is not the same as given type.");
            }
            return writer.ToString();
        }

        /// <summary>
        /// Serialize an object to XML string.
        /// Not using generic.
        /// This method is good when the type is unknown on runtime.
        /// </summary>
        /// <param name="obj">Object to be serialized.</param>
        /// <returns></returns>
        public static string Serialize(object obj)
        {
            StringWriter resultWriter = new StringWriter();
            XmlSerializer serializer = new XmlSerializer(obj.GetType());
            using (resultWriter)
            {
                serializer.Serialize(resultWriter, obj);
            }
            return resultWriter.ToString();
        }

        /// <summary>
        /// Deserialize a string of known type to an object.
        /// </summary>
        /// <typeparam name="T">Type of the object.</typeparam>
        /// <param name="XML"></param>
        /// <returns></returns>
        public static T DeserializeXMLString<T>(string XML)
        {
            try
            {
                T result = (T)DeserializeXMLString(XML, typeof(T));
                return result;
            }
            catch (ArgumentException argEx)
            {
                throw argEx;
            }
        }


        public static object DeserializeXMLString(string XML, Type ResultType)
        {
            MemoryStream XMLstream = new MemoryStream();
            StreamWriter memoryWriter = new StreamWriter(XMLstream);
            XmlSerializer serializer = new XmlSerializer(ResultType);
            memoryWriter.Write(XML);
            memoryWriter.Flush();
            XMLstream.Position = 0;
            XmlReader readerFromStream = new XmlTextReader(XMLstream);
            object result = null;
            if (serializer.CanDeserialize(readerFromStream))
            {
                result = serializer.Deserialize(readerFromStream);
            }
            else
            {
                throw new ArgumentException("XML is not the same as given type.");
            }
            return result;
        }

        public static T DeserializeXMLFile<T>(string FileName)
        {
            try
            {
                T result = (T)DeserializeXMLFile(FileName, typeof(T));
                return result;
            }
            catch (ArgumentException argEx)
            {
                throw argEx;
            }
        }

        public static object DeserializeXMLFile(string FileName, Type ResultType)
        {
            FileStream fileStream = new FileStream(FileName, FileMode.Open);
            XmlSerializer serializer = new XmlSerializer(ResultType);
            fileStream.Position = 0;
            XmlReader readerFromStream = new XmlTextReader(fileStream);
            object result = null;
            if (serializer.CanDeserialize(readerFromStream))
            {
                result = serializer.Deserialize(readerFromStream);
            }
            else
            {
                throw new ArgumentException("XML is not the same as given type.");
            }
            return result;
        }

        public static T DeserializeXMLStream<T>(Stream FileStream)
        {
            try
            {
                T result = (T)DeserializeXMLStream(FileStream, typeof(T));
                return result;
            }
            catch (ArgumentException argEx)
            {
                throw argEx;
            }
        }

        public static object DeserializeXMLStream(Stream FileStream, Type ResultType)
        {
            XmlSerializer serializer = new XmlSerializer(ResultType);
            FileStream.Position = 0;
            XmlReader readerFromStream = new XmlTextReader(FileStream);
            object result = null;
            if (serializer.CanDeserialize(readerFromStream))
            {
                result = serializer.Deserialize(readerFromStream);
            }
            else
            {
                throw new ArgumentException("XML is not the same as given type.");
            }
            return result;
        }
        #endregion
    }
}
