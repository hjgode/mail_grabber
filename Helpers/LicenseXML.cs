using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Xml;
using System.Xml.Serialization;
using System.IO;

namespace Helpers
{
    [XmlRootAttribute("CETerm")]
    public class LicenseXML
    {

        [XmlElement("license")]
        public license[] licenses;

        private XmlSerializer s = null;
        private Type type = null;
        public LicenseXML()
        {
            this.type = typeof(LicenseXML);
            this.s = new XmlSerializer(this.type);
        }
        public static LicenseXML Deserialize(byte[] bytes)
        {
            XmlSerializer xs = new XmlSerializer(typeof(LicenseXML));
            LicenseXML lic = new LicenseXML();
            try
            {
                System.IO.MemoryStream ms = new MemoryStream(bytes);
                ms.Position = 0;    //rewind to start
                lic = (LicenseXML)xs.Deserialize(ms);
            }
            catch (XmlException ex)
            {
                System.Diagnostics.Debug.WriteLine("xml ex:" + ex.Message);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("xml ex:" + ex.Message);
            }
            return lic;
        }

        public static LicenseXML Deserialize(Stream stream)
        {
            XmlSerializer xs = new XmlSerializer(typeof(LicenseXML));
            LicenseXML lic =new LicenseXML();
            try{
                stream.Position = 0;    //rewind to start
                lic=(LicenseXML)xs.Deserialize(stream);
            }catch(XmlException ex){
                System.Diagnostics.Debug.WriteLine("xml ex:" +ex.Message);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("xml ex:" + ex.Message);
            }
            return lic;
        }
        public static LicenseXML Deserialize(string sXMLfile)
        {
            XmlSerializer xs = new XmlSerializer(typeof(LicenseXML));
            StreamReader sr = new StreamReader(sXMLfile);
            LicenseXML lic = (LicenseXML)xs.Deserialize(sr);
            sr.Close();
            return lic;
        }
        public static void serialize(LicenseXML licensexml, string sXMLfile)
        {
            XmlSerializer xs = new XmlSerializer(typeof(LicenseXML));
            //omit xmlns:xsi from xml output
            //Create our own namespaces for the output
            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            //Add an empty namespace and empty value
            ns.Add("", "");            
            StreamWriter sw = new StreamWriter(sXMLfile);
            xs.Serialize(sw, licensexml, ns);
        }
        public static class TestXML
        {
            public static string runTest()
            {
                StringBuilder sb= new StringBuilder();
                //test deserialisation of license xml file
                sb.Append("\r\n### test deserialisation of license xml file ###\r\n");
                string testFile = utils.helpers.getAppPath() + "LicenseXMLFileSample.xml";
                LicenseXML licensexml = LicenseXML.Deserialize(utils.helpers.getAppPath() + "LicenseXMLFileSample.xml");
                sb.Append("\r\n### testing xml file read ###\r\n");
                foreach (license l in licensexml.licenses)
                    sb.Append(l.dumpData());

                //test deserialization of memory stream
                sb.Append("\r\n### testing deserialization of xml stream ###\r\n");
                System.IO.MemoryStream ms = new System.IO.MemoryStream();
                //read file into memorystream
                using (System.IO.FileStream file = new System.IO.FileStream(testFile, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {
                    byte[] bytes = new byte[file.Length];
                    file.Read(bytes, 0, (int)file.Length);
                    ms.Write(bytes, 0, (int)file.Length);
                    ms.Flush();
                    ms.WriteTo(new System.IO.FileStream(testFile + "_cpy", System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite));
                    sb.Append("\r\nmemory stream copy saved to" + testFile + "_out" + "\r\n");
                }

                LicenseXML xmlStreamed = LicenseXML.Deserialize(ms);
                if (xmlStreamed != null)
                {
                    foreach (license l in xmlStreamed.licenses)
                        sb.Append(l.dumpData());
                    //now test serialization of license data
                    LicenseXML.serialize(xmlStreamed, testFile + "_out");
                    sb.Append("\r\ndeserilized stream object serialization test saved to" + testFile + "_out" + "\r\n");
                }
                else
                    sb.Append("\r\nstream read FAILED");
                return sb.ToString();
            }
        }
    }
    public class license
    {
        [XmlElement("id")]
        public string id;
        [XmlElement("user")]
        public string user;
        [XmlElement("key")]
        public string key;

        public string dumpData()
        {
            StringBuilder sb=new StringBuilder();
            sb.Append(String.Format("\r\nid={0}, user={1}, key={2}", id, user, key));
            System.Diagnostics.Debug.WriteLine(sb.ToString());
            return sb.ToString();
        }
    }
}
