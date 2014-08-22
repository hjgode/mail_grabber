using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data;
using System.Data.SQLite;

namespace Helpers
{
    /*
     <CETerm>
      <license>
        <id>VM1C141578167</id>
        <user>Honeywell</user>
        <key>ZGNTZ4P5HQ7AAFHP9XZD6QSGU3Y2TQUY</key>
      </license>
     ...
      <license>
        <id>VM1C141578212</id>
        <user>Honeywell</user>
        <key>ZGNVXMF5AQ84X2GB9XZD6QSGU3Y2SD36</key>
      </license>
    </CETerm>
     */

    /// <summary>
    /// the license data consists of: id, user and key
    /// additionally we use ReceivedBy, SentAt,
    /// SQLite DB access
    /// </summary>
    public sealed class LicenseData:IDisposable
    {
        /// <summary>
        /// the deviceID of this license
        /// </summary>
        public string _deviceid;
        /// <summary>
        /// the customer for the license
        /// </summary>
        public string _customer;
        /// <summary>
        /// the key for this device
        /// </summary>
        public string _key;
        /// <summary>
        /// the order number
        /// </summary>
        public string _ordernumber;
        /// <summary>
        /// the order date of this purchase
        /// </summary>
        public DateTime _orderdate;
        /// <summary>
        /// the purchase order number
        /// </summary>
        public string _ponumber;
        /// <summary>
        /// the end customer of this license
        /// </summary>
        public string _endcustomer;
        /// <summary>
        /// the product valid for the license
        /// </summary>
        public string _product;
        /// <summary>
        /// number of licenses send as bundle
        /// </summary>
        public int _quantity;
        /// <summary>
        /// who grabbed this license data off exchange/outlook
        /// </summary>
        public string _receivedby;
        /// <summary>
        /// when was this license email sent
        /// </summary>
        public DateTime _sendat;

        static string dataSource = utils.helpers.getAppPath() + "license_data.db";
        SQLiteDataAdapter sql_data_adapter;
        DataSet data_set;
        System.Windows.Forms.BindingSource bs;

        SQLiteConnection _connection;
        SQLiteConnection connection
        {
            get
            {
                try
                {
                    if (_connection == null)
                    {
                        _connection = new SQLiteConnection();
                        _connection.ConnectionString = "Data Source=" + dataSource;
                        _connection.Open();
                        utils.helpers.addLog("db connection opened");
                    }
                }
                catch (SQLiteException ex)
                {
                    utils.helpers.addExceptionLog("add: " + ex.Message + "\r\n" + ex.StackTrace);
                }
                catch (Exception ex)
                {
                    utils.helpers.addExceptionLog("add: " + ex.Message + "\r\n" + ex.StackTrace);
                }
                return _connection;
            }
            
        }

        #region Singleton
        //singleton pattern: http://msdn.microsoft.com/en-us/library/ff650316.aspx
        private static volatile LicenseData instance;
        private static object syncRoot = new Object();

        private LicenseData()
        {
            createDB();
        }
        public static LicenseData Instance
        {
            get 
            {
                if (instance == null) 
                {
                lock (syncRoot) 
                {
                    if (instance == null) 
                        instance = new LicenseData();
                }
                }
                return instance;
            }
        }
        #endregion

        public System.Windows.Forms.BindingSource getDataset(){
            data_set = new DataSet();
            DataTable data_table = new DataTable();
            bs = new System.Windows.Forms.BindingSource();

            //SQLiteCommand sql_command = connection.CreateCommand();
            string command_string = "select * from licensedata";

            sql_data_adapter = new SQLiteDataAdapter(command_string, connection);
            SQLiteCommandBuilder commandBuilder = new SQLiteCommandBuilder(sql_data_adapter);

            sql_data_adapter.Fill(data_set);
            bs.DataSource = data_set.Tables[0];

            return bs;
        }

        public void filterData(System.Windows.Forms.DataGridView dgv, string s)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            dt.DefaultView.RowFilter = "deviceid like '%" + s + "%'";
        }

        public bool clearData()
        {
            bool bRet = false;
            SQLiteCommand command = connection.CreateCommand();
            string cmdstr = "DELETE FROM licensedata where 1=1;";
            try
            {
                command.CommandText = cmdstr;
                int iRes = command.ExecuteNonQuery();
                utils.helpers.addLog("deleted " + iRes.ToString() + " data");
                command.Dispose();
                sql_data_adapter.Update(data_set);
                data_set.AcceptChanges();
                bRet = true;
            }
            catch (Exception ex)
            {
                utils.helpers.addExceptionLog("clearData Exception: "+ex.Message);
            }
            return bRet;
        }

        bool createDB()
        {
            bool bRes = false;
            string fieldlist = getFieldNamesAndTypeForCreate();
            // Erstellen der Tabelle, sofern diese noch nicht existiert.
            string cmdText = "CREATE TABLE IF NOT EXISTS licensedata " +
                getFieldNamesAndTypeForCreate()+
                //"( id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, " +
                //"  " + licenseCols.dataCols[0].name + "  " + licenseCols.dataCols[0].type.ToString() + ", " +
                //"  " + licenseCols.dataCols[1].name + "  " + licenseCols.dataCols[1].type.ToString() + ", " +
                //"  " + licenseCols.dataCols[2].name + "  " + licenseCols.dataCols[2].type.ToString() + ", " +
                //"  " + licenseCols.dataCols[3].name + "  " + licenseCols.dataCols[3].type.ToString() + ", " +
                //"  " + licenseCols.dataCols[4].name + "  " + licenseCols.dataCols[4].type.ToString() + ", " +
                //"  " + licenseCols.dataCols[5].name + "  " + licenseCols.dataCols[5].type.ToString() + ", " +
                //"  " + licenseCols.dataCols[6].name + "  " + licenseCols.dataCols[6].type.ToString() + ", " +
                //"  " + licenseCols.dataCols[7].name + "  " + licenseCols.dataCols[7].type.ToString() + ", " +
                //"  " + licenseCols.dataCols[8].name + "  " + licenseCols.dataCols[8].type.ToString() + ", " +
                //"  " + licenseCols.dataCols[9].name + "  " + licenseCols.dataCols[9].type.ToString() + ", " +
                //"  " + licenseCols.dataCols[10].name + "  " + licenseCols.dataCols[10].type.ToString() + "  " +
                //" )"+
                ";";
            try
            {
                SQLiteConnection conn = connection;
                SQLiteCommand command = new SQLiteCommand(conn);

                command.CommandText = cmdText;
                command.ExecuteNonQuery();
                command.Dispose();
                bRes = true;
                utils.helpers.addLog("CreateDB OK");
            }
            catch (SQLiteException ex)
            {
                utils.helpers.addExceptionLog("add: " + cmdText + "\r\n" + ex.Message + "\r\n" + ex.StackTrace);
            }
            catch (Exception ex)
            {
                utils.helpers.addExceptionLog("add: " + cmdText + "\r\n" + ex.Message + "\r\n" + ex.StackTrace);
            }
            return bRes;
        }

        string getFieldNamesAndTypeForCreate()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("( id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, ");
            for (int ix = 0; ix < licenseCols.dataCols.Length; ix++)
            {
                sb.Append(string.Format(" {0} {1} ", licenseCols.dataCols[ix].name, licenseCols.dataCols[ix].type.ToString()));
                if (ix < licenseCols.dataCols.Length - 1)
                    sb.Append(",");
            }
            sb.Append(" )");
            return sb.ToString();
        }

        public void Dispose()
        {
            SQLiteConnection conn = connection;
            conn.Close();
            conn.Dispose();
            conn = null;
            utils.helpers.addLog("LicenseDATA Disposed");
        }

        public bool add(            
            string deviceid,
            string customer,
            string key,
            string ordernumber,
            DateTime orderdate,
            string ponumber,
            string endcustomer,
            string product,
            int quantity,
            string receivedby,
            DateTime sendat)
        {
            bool bRes = false;

            _deviceid = deviceid;
            _customer = customer;
            _key = key;
            _ordernumber = ordernumber;
            _orderdate = orderdate;
            _ponumber = ponumber;
            _endcustomer = endcustomer;
            _product = product;
            _quantity = quantity;
            _receivedby = receivedby;
            _sendat = sendat;
            string cmdText = "INSERT INTO licensedata " +
                    getSQLFieldListForInsert() +
                //"(id, "+
                //licenseCols.dataCols[0].name +", "+
                //licenseCols.dataCols[1].name + ", " +
                //licenseCols.dataCols[2].name + ", " +
                //licenseCols.dataCols[3].name + ", " +
                //licenseCols.dataCols[4].name + ", " +
                //licenseCols.dataCols[5].name + ", " +
                //licenseCols.dataCols[6].name + ", " +
                //licenseCols.dataCols[7].name + ", " +
                //licenseCols.dataCols[8].name + 
                //") " +
                    "VALUES(" +
                    "NULL, " +   //auto id
                    "'" + deviceid + "'," +
                    "'" + customer + "'," +
                    "'" + key + "'," +
                    "'" + ordernumber + "'," +
                    "'" + orderdate.ToString("yyyy-MM-dd HH:mm:ss") + "'," +// '2007-01-01 10:00:00', yyyy-mm-dd HH:mm:ss
                    "'" + ponumber + "'," +
                    "'" + endcustomer + "'," +
                    "'" + product + "'," +
                    "" + quantity.ToString() + "," +
                    "'" + receivedby + "'," +
                    "'" + sendat.ToString("yyyy-MM-dd HH:mm:ss") + "'" +
                    ")";
            try
            {
                SQLiteConnection conn = connection;
                SQLiteCommand command = new SQLiteCommand(conn);

                string namelist = getSQLFieldListForInsert();
                // Einfügen eines Test-Datensatzes.
                command.CommandText = cmdText;
                int iRes = command.ExecuteNonQuery();
                command.Dispose();
                sql_data_adapter.Update(data_set);
                data_set.AcceptChanges();
                bRes = true;
                utils.helpers.addLog("added "+iRes.ToString()+" new data");
            }
            catch (SQLiteException ex)
            {
                utils.helpers.addExceptionLog("add: " + cmdText+"\r\n" + ex.Message + "\r\n" + ex.StackTrace);
            }
            catch (Exception ex)
            {
                utils.helpers.addExceptionLog("add: " + cmdText + "\r\n" + ex.Message + "\r\n" + ex.StackTrace);
            }
            finally
            {
            }
            return bRes;
        }

        string getSQLFieldListForInsert()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(" (id, ");
            for (int ix=0; ix<licenseCols.dataCols.Length; ix++)// licenseCols lc in licenseCols.dataCols)
            {
                sb.Append(licenseCols.dataCols[ix].name);
                if (ix < licenseCols.dataCols.Length - 1)
                    sb.Append(",");
            }            
            sb.Append(") ");
            return sb.ToString();
        }
    }
    public class licenseCols
    {
        public string name;
        public SqlDbType type;
        public licenseCols(string s, SqlDbType t)
        {
            name = s;
            type = t;
        }
        public static string[] LicenseDataColumns = {
                "deviceid",//id     0
                "customer",//user   1
                "key",     //       2
                "ordernumber", //   3
                "orderdate",   //   4
                "ponumber", //      5
                "endcustomer",//    6
                "product", //       7
                "quantity", //      8
                "receivedby", //    9
                "sendat" //         10
        };

        public static licenseCols[] dataCols ={
            new licenseCols(LicenseDataColumns[0], SqlDbType.NVarChar),
            new licenseCols(LicenseDataColumns[1], SqlDbType.NVarChar),
            new licenseCols(LicenseDataColumns[2], SqlDbType.NVarChar),
            new licenseCols(LicenseDataColumns[3], SqlDbType.NVarChar),
            new licenseCols(LicenseDataColumns[4], SqlDbType.DateTime),
            new licenseCols(LicenseDataColumns[5], SqlDbType.NVarChar),
            new licenseCols(LicenseDataColumns[6], SqlDbType.NVarChar),
            new licenseCols(LicenseDataColumns[7], SqlDbType.NVarChar),
            new licenseCols(LicenseDataColumns[8], SqlDbType.Int),
            new licenseCols(LicenseDataColumns[9], SqlDbType.NVarChar),
            new licenseCols(LicenseDataColumns[10], SqlDbType.DateTime),
        };
    }
}
