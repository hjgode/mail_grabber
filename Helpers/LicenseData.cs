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
    public class LicenseData:IDisposable
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
        static SQLiteConnection connection = null;

        SQLiteConnection getConnection()
        {
            try
            {
                if (connection == null)
                {
                    connection = new SQLiteConnection();
                    connection.ConnectionString = "Data Source=" + dataSource;
                    connection.Open();
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
            return connection;
        }

        public LicenseData()
        {
            createDB();
        }

        public DataSet getDataset(){
            DataSet data_set = new DataSet();

            SQLiteCommand sql_command = getConnection().CreateCommand();
            string command_string = "select * from licensedata";

            SQLiteDataAdapter sql_data_adapter = new SQLiteDataAdapter(command_string, getConnection());
            SQLiteCommandBuilder sql_command_builder = new SQLiteCommandBuilder(sql_data_adapter);

            sql_data_adapter.Fill(data_set);
            return data_set;
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
                SQLiteConnection conn = getConnection();
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
            SQLiteConnection conn = getConnection();
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
                SQLiteConnection conn = getConnection();
                SQLiteCommand command = new SQLiteCommand(conn);

                string namelist = getSQLFieldListForInsert();
                // Einfügen eines Test-Datensatzes.
                command.CommandText = cmdText;
                int iRes = command.ExecuteNonQuery();
                command.Dispose();
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
