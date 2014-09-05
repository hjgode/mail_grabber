using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace outlk_grabber
{
    [XmlRoot("license_grabber")]
    public class MySettings
    {
        [XmlElement("OutlookProfile")]
        public string OutlookProfile = @"";

        [XmlElement("ShowDialog")]
        public bool ShowDialog = false;

        [XmlElement("NewSession")]
        public bool NewSession = false;

        [XmlElement("Username")]
        public string Username = @"global\e841719";

        [XmlElement("Domainname")]
        public string Domainname = "";

        [XmlElement("SQLiteDataBaseFilename")]
        public string SQLiteDataBaseFilename = "licensedata.db";

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
            StreamReader sr = null;
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
                if (sr != null)
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
