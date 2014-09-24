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

        //form layout
        [XmlElement("form_location")]
        public System.Drawing.Point form_location = new System.Drawing.Point(0, 0);
        [XmlElement("form_size")]
        public System.Drawing.Size form_size = new System.Drawing.Size(800, 600);

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
            count = dgv_cols.Count;
            
        }

        #region codeproject
        // http://www.codeproject.com/Tips/394133/XML-Serialization-and-Deserialization-in-Csharp

        public static bool Serialize(DataGridViewSettings value, String filename)
        {
            if (value == null)
            {
                return false;
            }
            try
            {
                XmlSerializer _xmlserializer = new XmlSerializer(typeof(DataGridViewSettings));
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
        static bool Serialize<T>(T value, String filename)
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
        public static DataGridViewSettings Deserialize(String filename)
        {
            if (string.IsNullOrEmpty(filename))
            {
                return new DataGridViewSettings();
            }
            try
            {
                XmlSerializer _xmlSerializer = new XmlSerializer(typeof(DataGridViewSettings));
                Stream stream = new FileStream(filename, FileMode.Open, FileAccess.Read);
                var result = (DataGridViewSettings)_xmlSerializer.Deserialize(stream);
                stream.Close();
                return result;
            }
            catch (Exception ex)
            {
                return new DataGridViewSettings();
            }
        }
        static t Deserialize<t>(String filename)
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

    }
}
