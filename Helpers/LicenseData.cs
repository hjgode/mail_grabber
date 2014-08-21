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
    public class LicenseData:IDisposable
    {
        public string _deviceid;
        public string _customer;
        public string _key;
        public string _ordernumber;
        public DateTime _orderdate;
        public string _ponumber;
        public string _endcustomer;
        public string _product;
        public int _quantity;

        static string dataSource = utils.helpers.getAppPath() + "license_data.db";
        static SQLiteConnection connection = null;
        SQLiteConnection getConnection()
        {
            if (connection == null)
                connection = new SQLiteConnection();
            return connection;
        }
        public LicenseData()
        {
            getConnection().ConnectionString = "Data Source=" + dataSource;
            getConnection().Open();
            createDB();
        }

        void createDB()
        {
            SQLiteConnection conn = getConnection();
            SQLiteCommand command = new SQLiteCommand(conn);

            // Erstellen der Tabelle, sofern diese noch nicht existiert.
            command.CommandText = "CREATE TABLE IF NOT EXISTS licensedata " +
                "( id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, " +
                "  " + licenseCols.dataCols[0].name + "  "+ SqlDbType.NVarChar.ToString() + ", " +
                "  " + licenseCols.dataCols[1].name + "  "+ SqlDbType.NVarChar.ToString() + ", " +
                "  " + licenseCols.dataCols[2].name + "  " + SqlDbType.NVarChar.ToString() + ", " +
                "  " + licenseCols.dataCols[3].name + "  " + SqlDbType.NVarChar.ToString() + ", " +
                "  " + licenseCols.dataCols[4].name + "  " + SqlDbType.DateTime.ToString() + ", " +
                "  " + licenseCols.dataCols[5].name + "  " + SqlDbType.NVarChar.ToString() + ", " +
                "  " + licenseCols.dataCols[6].name + "  " + SqlDbType.NVarChar.ToString() + ", " +
                "  " + licenseCols.dataCols[7].name + "  " + SqlDbType.NVarChar.ToString() + ", " +
                "  " + licenseCols.dataCols[8].name + "  " + SqlDbType.Int.ToString() + "  " +
                " );";

            command.ExecuteNonQuery();
            command.Dispose();
        }

        public void Dispose()
        {
            SQLiteConnection conn = getConnection();
            conn.Close();
            conn.Dispose();
            conn = null;
        }
        public void add(LicenseXML xml)
        {

        }
        public void add(
            string deviceid,
            string customer,
            string key,
            string ordernumber,
            DateTime orderdate,
            string ponumber,
            string endcustomer,
            string product,
            int quantity)
        {

            _deviceid = deviceid;
            _customer = customer;
            _key = key;
            _ordernumber = ordernumber;
            _orderdate = orderdate;
            _ponumber = ponumber;
            _endcustomer = endcustomer;
            _product = product;
            _quantity = quantity;

            SQLiteConnection conn = getConnection();
            SQLiteCommand command = new SQLiteCommand(conn);
            // Einfügen eines Test-Datensatzes.
            string datetime = orderdate.ToString("yyyy-MM-dd HH:mm:ss");
            command.CommandText = "INSERT INTO licensedata "+
                "(id, "+
                licenseCols.dataCols[0].name +", "+
                licenseCols.dataCols[1].name + ", " +
                licenseCols.dataCols[2].name + ", " +
                licenseCols.dataCols[3].name + ", " +
                licenseCols.dataCols[4].name + ", " +
                licenseCols.dataCols[5].name + ", " +
                licenseCols.dataCols[6].name + ", " +
                licenseCols.dataCols[7].name + ", " +
                licenseCols.dataCols[8].name + 
                ") " +
                "VALUES("+
                "NULL, "+   //auto id
                "'" + deviceid +"',"+
                "'" + customer + "'," +
                "'" + key + "'," +
                "'" + ordernumber + "'," +
                "'" + orderdate.ToString("yyyy-MM-dd HH:mm:ss") + "'," +// '2007-01-01 10:00:00', yyyy-mm-dd hh:mm:ss
                "'" + ponumber + "'," +
                "'" + endcustomer + "'," +
                "'" + product + "'," +
                "" + quantity.ToString() + 
                ")";
            command.ExecuteNonQuery();
            command.Dispose();
        }

    }
    public class licenseCols
    {
        public string name;
        public System.Type type;
        public licenseCols(string s, System.Type t)
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
                "quantity" //       8
                                             };
        public static licenseCols[] dataCols ={
            new licenseCols(LicenseDataColumns[0], typeof(string)),
            new licenseCols(LicenseDataColumns[1], typeof(string)),
            new licenseCols(LicenseDataColumns[2], typeof(string)),
            new licenseCols(LicenseDataColumns[3], typeof(string)),
            new licenseCols(LicenseDataColumns[4], typeof(DateTime)),
            new licenseCols(LicenseDataColumns[5], typeof(string)),
            new licenseCols(LicenseDataColumns[6], typeof(string)),
            new licenseCols(LicenseDataColumns[7], typeof(string)),
            new licenseCols(LicenseDataColumns[8], typeof(int)),
                                           };
    }
}
