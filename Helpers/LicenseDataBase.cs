using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data;
using System.Data.SQLite;

using System.Threading;

namespace Helpers
{
    #region XML_SAMPLE_LICENSES
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
    #endregion

    /// <summary>
    /// the license data consists of: id, user and key
    /// additionally we use ReceivedBy, SentAt,
    /// SQLite DB access
    /// </summary>
    public class LicenseDataBase:IDisposable
    {
        bool bHandleEvents=true;

        /// <summary>
        /// object to sync access to the queue
        /// </summary>
        object lockQueue = new object();
        /// <summary>
        /// a queue to hold LicenseData to be processed in the background
        /// </summary>
        Queue<LicenseData> _licenseDataQueue = new Queue<LicenseData>();
        /// <summary>
        /// background thread to dequeue LicenseData
        /// </summary>
        Thread processQueueThread=null;
        static bool bRunProcessQueueThread = true;

        static string dataSource = utils.helpers.getAppPath() + "license_data.db";
        static string connectionString = "Data Source=" + dataSource;
        System.Windows.Forms.DataGridView _dgv;

        /// <summary>
        /// do not refresh datagridview when filtered view
        /// </summary>
        bool bFiltered = false;

        public enum FilterType
        {
            OrderNumber,
            PurchaseOrder,
            DeviceID,
            KeyID,
            none,
        }

        #region Singleton

        /*
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
        */
        #endregion


        public LicenseDataBase(ref System.Windows.Forms.DataGridView dgv){
            _dgv=dgv;
            /*
            processQueueThread = new Thread(processingthread);
            processQueueThread.Name = "process LicenseData queue thread";
            processQueueThread.Start(this);
            */
        }

        public System.Windows.Forms.BindingSource getDataset(){
            System.Windows.Forms.BindingSource bs = new System.Windows.Forms.BindingSource();

            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                DataSet data_set = new DataSet();

                //SQLiteCommand sql_command = connection.CreateCommand();
                string command_string = "select * from licensedata";

                SQLiteDataAdapter sql_data_adapter = new SQLiteDataAdapter(command_string, connection);
                SQLiteCommandBuilder commandBuilder = new SQLiteCommandBuilder(sql_data_adapter);

                sql_data_adapter.Fill(data_set);
                DataTable data_table = data_set.Tables[0];

                bs.DataSource = data_table;
            }

            return bs;
        }

        public void filterData(ref System.Windows.Forms.DataGridView dgv, FilterType ft, string s)
        {
            System.Windows.Forms.BindingSource bs = getDataset();
            DataTable dt = bs.DataSource as DataTable;
            string rowFilter = "";
            if (s == "")
            {
                ft = FilterType.none;
            }

            switch (ft)
            {
                case FilterType.none:
                    rowFilter = "";
                    break;
                case FilterType.DeviceID:
                    rowFilter = "deviceid like '%" + s + "%'";
                    break;
                case FilterType.KeyID:
                    rowFilter = "key like '%" + s + "%'";
                    break;
                case FilterType.OrderNumber:
                    rowFilter = "ordernumber like '%" + s + "%'";
                    break;
                case FilterType.PurchaseOrder:
                    rowFilter = "ponumber like '%" + s + "%'";
                    break;
                default:
                    rowFilter = "";
                    break;
            }
            dt.DefaultView.RowFilter = rowFilter;
            dgv.DataSource = dt;
        }

        public void doRefresh(ref System.Windows.Forms.DataGridView dgv)
        {
            if (bFiltered)
                return;
            int iRow = 0, iCol = 0;
            if (dgv.Rows.Count > 0)
            {
                iRow = dgv.CurrentCell.RowIndex;
                iCol = dgv.CurrentCell.ColumnIndex;
            }
            dgv.ResetBindings();
            //dgv.DataSource=null;
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                System.Data.DataSet ds = new DataSet();
                SQLiteDataAdapter da = new SQLiteDataAdapter("SELECT * FROM licensedata", connection);
                da.Fill(ds);
                dgv.DataSource = ds.Tables[0];
            }
            if (dgv.Rows.Count > 0)
            {
                dgv.CurrentCell = dgv.Rows[iRow].Cells[iCol];
            }
        }

        public bool clearData()
        {
            bool bRet = false;
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                SQLiteCommand command = connection.CreateCommand();
                string cmdstr = "DELETE FROM licensedata where 1=1;";
                try
                {
                    command.CommandText = cmdstr;
                    int iRes = command.ExecuteNonQuery();
                    utils.helpers.addLog("deleted " + iRes.ToString() + " data");
                    command.Dispose();

                    doRefresh(ref _dgv);
                    bRet = true;
                }
                catch (Exception ex)
                {
                    utils.helpers.addExceptionLog("clearData Exception: " + ex.Message);
                }
            } return bRet;
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
                using (SQLiteConnection connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    SQLiteCommand command = new SQLiteCommand(connection);
                    command.CommandText = cmdText;
                    command.ExecuteNonQuery();
                    command.Dispose();
                    bRes = true;
                    utils.helpers.addLog("CreateDB OK");
                }
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
            bHandleEvents = false;
            if (processQueueThread != null)
            {
                bRunProcessQueueThread = false;
                Thread.Sleep(1000);
                if(!processQueueThread.Join(1000))
                    processQueueThread.Abort();
                processQueueThread = null;
            }
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
            if (existsData(deviceid, key))
            {
                utils.helpers.addLog("add abandoned for existing datarow");
                return true;
            }
            LicenseData licenseData = new LicenseData();
            licenseData._deviceid = deviceid;
            licenseData._customer = customer;
            licenseData._key = key;
            licenseData._ordernumber = ordernumber;
            licenseData._orderdate = orderdate;
            licenseData._ponumber = ponumber;
            licenseData._endcustomer = endcustomer;
            licenseData._product = product;
            licenseData._quantity = quantity;
            licenseData._receivedby = receivedby;
            licenseData._sendat = sendat;

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
                    "'" + licenseData._deviceid + "'," +
                    "'" + licenseData._customer + "'," +
                    "'" + licenseData._key + "'," +
                    "'" + licenseData._ordernumber + "'," +
                    "'" + licenseData._orderdate.ToString("yyyy-MM-dd HH:mm:ss") + "'," +// '2007-01-01 10:00:00', yyyy-mm-dd HH:mm:ss
                    "'" + licenseData._ponumber + "'," +
                    "'" + licenseData._endcustomer + "'," +
                    "'" + licenseData._product + "'," +
                    "" + licenseData._quantity.ToString() + "," +
                    "'" + licenseData._receivedby + "'," +
                    "'" + licenseData._sendat.ToString("yyyy-MM-dd HH:mm:ss") + "'" +
                    ")";
            try
            {
                using (SQLiteConnection connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    SQLiteCommand command = new SQLiteCommand(connection);

                    string namelist = getSQLFieldListForInsert();
                    // Einfügen eines Test-Datensatzes.
                    command.CommandText = cmdText;
                    int iRes = command.ExecuteNonQuery();
                    command.Dispose();
                    doRefresh(ref _dgv);
                    bRes = true;
                    utils.helpers.addLog("added " + iRes.ToString() + " new data");
                }
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

        public static void processingthread(object param)
        {
            LicenseDataBase ldb=(LicenseDataBase)param;
            bool bNewDataAvailable = false;
            do
            {
                try
                {
                    lock (ldb.lockQueue)
                    {
                        while (ldb._licenseDataQueue.Count > 0)
                        {
                            LicenseData licData = ldb._licenseDataQueue.Dequeue();
                            ldb.add(licData);
                            bNewDataAvailable = true;
                        }
                        if (bNewDataAvailable)
                        {
                            //fire update event
                            ldb.OnStateChanged(new StatusEventArgs(StatusType.bulk_mail, "new license data received"));
                            bNewDataAvailable = false;
                        }
                    }
                    Thread.Sleep(5000);
                }
                catch (ThreadAbortException ex)
                {
                    utils.helpers.addExceptionLog("Exception in processingthread(): " + ex.Message);
                    bRunProcessQueueThread = false;
                }
                catch (Exception ex)
                {
                    utils.helpers.addExceptionLog("Exception in processingthread(): " + ex.Message);
                }
            } while (bRunProcessQueueThread);
        }

        public event Helpers.StateChangedEventHandler StateChanged;
        protected virtual void OnStateChanged(StatusEventArgs args)
        {
            if (!bHandleEvents)
                return;
            System.Diagnostics.Debug.WriteLine("onStateChanged: " + args.eStatus.ToString() + ":" + args.strMessage);
            StateChangedEventHandler handler = StateChanged;
            if (handler != null)
            {
                handler(this, args);
            }
        }

        public bool addQueued(LicenseData licenseData)
        {
            bool bRes = false;
            lock (lockQueue)
            {
                _licenseDataQueue.Enqueue(licenseData);
                bRes = true;
            }
            return bRes;
        }

        bool existsData(string sDeviceID, string sKey)
        {
            bool bRet = false;
            string cmdText=string.Format("SELECT * FROM licensedata WHERE deviceid='{0}' AND key='{1}';", sDeviceID, sKey);
            try
            {
                using (SQLiteConnection connection = new SQLiteConnection(connectionString))
                {
                openConn:
                    try
                    {
                        connection.Open();
                    }
                    catch (SQLiteException ex)
                    {
                        if (ex.ErrorCode == (int)SQLiteErrorCode.Locked)
                        {
                            Thread.Sleep(1000);
                            goto openConn;
                        }
                    }

                    SQLiteCommand command = new SQLiteCommand(connection);
                    command.CommandText = cmdText;

                    SQLiteDataReader rdr = command.ExecuteReader(CommandBehavior.SingleResult);
                    if (rdr.HasRows)
                    {
                        utils.helpers.addLog("found at least one existing datarow for " + sDeviceID + " and " + sKey);
                        bRet = true;
                    }
                    else
                        utils.helpers.addLog("found no existing datarow for " + sDeviceID + " and " + sKey);

                    rdr.Close();
                    command.Dispose();
                }
            }
            catch (SQLiteException ex)
            {
                utils.helpers.addExceptionLog("add: " + cmdText + "\r\n" + ex.Message + "\r\n" + ex.StackTrace);
            }
            catch (Exception ex)
            {
                utils.helpers.addExceptionLog("add: " + cmdText + "\r\n" + ex.Message + "\r\n" + ex.StackTrace);
            }
            finally
            {
            }

            return bRet;
        }
        /// <summary>
        /// add LicenseData to DataBase, blocks for some time
        /// use manual doRefresh
        /// </summary>
        /// <param name="licenseData"></param>
        /// <returns></returns>
        bool add(LicenseData licenseData)
        {
            bool bRes = false;
            if (existsData(licenseData._deviceid, licenseData._key))
            {
                utils.helpers.addLog("add abandoned for existing datarow");
                return true;
            }
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
                    "'" + licenseData._deviceid + "'," +
                    "'" + licenseData._customer + "'," +
                    "'" + licenseData._key + "'," +
                    "'" + licenseData._ordernumber + "'," +
                    "'" + licenseData._orderdate.ToString("yyyy-MM-dd HH:mm:ss") + "'," +// '2007-01-01 10:00:00', yyyy-mm-dd HH:mm:ss
                    "'" + licenseData._ponumber + "'," +
                    "'" + licenseData._endcustomer + "'," +
                    "'" + licenseData._product + "'," +
                    "" + licenseData._quantity.ToString() + "," +
                    "'" + licenseData._receivedby + "'," +
                    "'" + licenseData._sendat.ToString("yyyy-MM-dd HH:mm:ss") + "'" +
                    ")";
            try
            {
                using (SQLiteConnection connection = new SQLiteConnection(connectionString))
                {
                openConn:
                    try
                    {
                        connection.Open();
                    }
                    catch (SQLiteException ex)
                    {
                        if (ex.ErrorCode == (int)SQLiteErrorCode.Locked)
                        {
                            Thread.Sleep(1000);
                            goto openConn;
                        }
                    }

                    SQLiteCommand command = new SQLiteCommand(connection);

                    string namelist = getSQLFieldListForInsert();
                    // Einfügen eines Test-Datensatzes.
                    command.CommandText = cmdText;
                    int iRes = command.ExecuteNonQuery();
                    //get id of last insert
                    command.CommandText = "select last_insert_rowid()";
                    // The row ID is a 64-bit value - cast the Command result to an Int64.
                    //
                    long LastRowID64 = (long)command.ExecuteScalar();
                    // Then grab the bottom 32-bits as the unique ID of the row.
                    int LastRowID = (int)LastRowID64;
                    command.Dispose();

                    /*
                    //now add same data to drid, can not be done from separate thread!
                    DataTable dt = (DataTable)_dgv.DataSource;
                    DataRow dr = dt.NewRow();
                    dr["id"] = LastRowID;
                    dr[licenseCols.LicenseDataColumns[0]] = licenseData._deviceid;
                    dr[licenseCols.LicenseDataColumns[1]] = licenseData._customer;
                    dr[licenseCols.LicenseDataColumns[2]] = licenseData._key;
                    dr[licenseCols.LicenseDataColumns[3]] = licenseData._ordernumber;
                    dr[licenseCols.LicenseDataColumns[4]] = licenseData._orderdate;
                    dr[licenseCols.LicenseDataColumns[5]] = licenseData._ponumber;
                    dr[licenseCols.LicenseDataColumns[6]] = licenseData._endcustomer;
                    dr[licenseCols.LicenseDataColumns[7]] = licenseData._product;
                    dr[licenseCols.LicenseDataColumns[8]] = licenseData._quantity;
                    dr[licenseCols.LicenseDataColumns[9]] = licenseData._receivedby;
                    dr[licenseCols.LicenseDataColumns[10]] = licenseData._sendat;
                    _dgv.Rows.Add(dr);
                    */

                    bRes = true;
                    utils.helpers.addLog("added " + iRes.ToString() + " new data");
                }
            }
            catch (SQLiteException ex)
            {
                utils.helpers.addExceptionLog("add: " + cmdText + "\r\n" + ex.Message + "\r\n" + ex.StackTrace);
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
