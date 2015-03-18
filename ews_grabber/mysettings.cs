using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace ews_grabber
{
    [XmlRoot("license_grabber")]
    public class MySettings
    {
        [XmlElement("ExchangeServiceURL")]
        public string ExchangeServiceURL = @"https://az18-cas-01.global.ds.honeywell.com/EWS/Exchange.asmx";

        [XmlElement("ExchangeWebProxy")]
        public string ExchangeWebProxy = "fr44proxy.honeywell.com";

        [XmlElement("EchangeWebProxyPort")]
        public int EchangeWebProxyPort = 8080;

        [XmlElement("ExchangeUsername")]
        public string ExchangeUsername = "";

        [XmlElement("ExchangeDomainname")]
        public string ExchangeDomainname = "";

        [XmlElement("SQLiteDataBaseFilename")]
        public string SQLiteDataBaseFilename = "licensedata.db";

        [XmlElement("UseWebProxy")]
        public bool UseWebProxy = false;

        [XmlElement("UseLogging")]
        public bool UseLogging = true;

        [XmlIgnore]
        string settingsFile = "mysettings.xml";

        private XmlSerializer xs = null;
        private Type type = null;
        public MySettings()
        {
            string appPath = utils.helpers.getAppPath();
            settingsFile = appPath + settingsFile;
            this.type = typeof(MySettings);
            this.xs = new XmlSerializer(this.type);
        }

        public MySettings load()
        {
            MySettings settings = new MySettings();
            StreamReader sr=null;
            try
            {
                XmlSerializer xs = new XmlSerializer(typeof(MySettings));
                sr = new StreamReader(settingsFile);
                settings = (MySettings)xs.Deserialize(sr);
                sr.Close();
            }
            catch (Exception ex)
            {
                utils.helpers.addExceptionLog("MySettings load Exception: ");
            }
            finally
            {
                if(sr!=null)
                    sr.Close();
            }
            
            return settings;
        }

        public bool Save(MySettings settings)
        {
            StreamWriter sw = null;
            try
            {
                //now global: XmlSerializer xs = new XmlSerializer(typeof(MySettings));
                //omit xmlns:xsi from xml output
                //Create our own namespaces for the output
                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                //Add an empty namespace and empty value
                ns.Add("", "");
                sw = new StreamWriter(settingsFile);
                xs.Serialize(sw, settings, ns);
                return true;
            }
            catch (Exception ex)
            {
                utils.helpers.addExceptionLog(ex.Message);
                return false;
            }
            finally
            {
                if (sw != null)
                    sw.Close();
            }
        }        
    }
}
