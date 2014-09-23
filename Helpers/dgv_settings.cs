using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace utils
{
    [XmlRoot("DatagridviewSettings")]
    public class DataGridViewSettings
    {
        [XmlElement("Columns")]
        public DataGridViewColumn[] columns{get;set;}
        [XmlElement("count")]
        public int count{get;set;}

        [XmlIgnore]
        static string settingsFile = "dgv_settings.xml";
        [XmlIgnore]
        public const string settingsFileConst = "dgv_settings.xml";

        private XmlSerializer xs = null;
        private Type type = null;

        public DataGridViewSettings()
        {
            string appPath = utils.helpers.getAppPath();
            settingsFile = appPath + settingsFile;
            this.type = typeof(DataGridViewSettings);
            this.xs = new XmlSerializer(typeof(DataGridViewSettings));//new XmlSerializer(this.type);
            List<DataGridViewColumn> dgv_cols=new List<DataGridViewColumn>();
            for (int x = 0; x < Helpers.licenseCols.DataGridColHeaders.Length; x++)
                dgv_cols.Add(new DataGridViewColumn(Helpers.licenseCols.DataGridColHeaders[x]));
            columns = dgv_cols.ToArray();
        }
        #region codeproject
        // http://www.codeproject.com/Tips/394133/XML-Serialization-and-Deserialization-in-Csharp
        public static bool Serialize<T>(T value, String filename)
        {
            if (value == null)
            {
                return false;
            }
            try
            {
                XmlSerializer _xmlserializer = new XmlSerializer(typeof(T));
                Stream stream = new FileStream(filename, FileMode.Create);
                _xmlserializer.Serialize(stream, value);
                stream.Close();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public static t Deserialize<t>(String filename)
        {
            if (string.IsNullOrEmpty(filename))
            {
                return default(t);
            }
            try
            {
                XmlSerializer _xmlSerializer = new XmlSerializer(typeof(t));
                Stream stream = new FileStream(filename, FileMode.Open, FileAccess.Read);
                var result = (t)_xmlSerializer.Deserialize(stream);
                stream.Close();
                return result;
            }
            catch (Exception ex)
            {
                return default(t);
            }
        }
        #endregion

        public class DataGridViewColumn
        {
            public bool visible=true;
            public string header = "col";
            public int width = 80;
            public DataGridViewColumn(string s)
            {
                header = s;
            }
            public DataGridViewColumn()
            {
            }
        }

        public DataGridViewSettings load()
        {
            DataGridViewSettings settings = new DataGridViewSettings();
            StreamReader sr=null;
            try
            {
                XmlSerializer xs = new XmlSerializer(typeof(DataGridViewSettings));
                sr = new StreamReader(settingsFile);
                settings = (DataGridViewSettings)xs.Deserialize(sr);
                sr.Close();
            }
            catch (Exception ex)
            {
                utils.helpers.addExceptionLog("DataGridViewSettings load Exception: ");
            }
            finally
            {
                if(sr!=null)
                    sr.Close();
            }
            
            return settings;
        }

        public bool Save(DataGridViewSettings settings)
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
                xs = new XmlSerializer(typeof(DataGridViewSettings));//new XmlSerializer(this.type);
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
