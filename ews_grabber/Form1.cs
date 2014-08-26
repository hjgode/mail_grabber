using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using ExchangeMail;

using Helpers;

namespace ews_grabber
{
    public partial class Form1 : Form
    {
        ews _ews;
        utils.UserData _userData;
        //data and database
        LicenseDataBase _licenseDataBase;
        LicenseMail _licenseMail = new LicenseMail();
        bool _bFilterActive = false;

        public Form1()
        {
            InitializeComponent();
            string appPath = utils.helpers.getAppPath();

            _licenseMail = new LicenseMail();
            //subscribe to new license mails
            _licenseMail.StateChanged += new StateChangedEventHandler(on_new_licensemail);

            addLog("Please select Exchange-Connect to start test");
            string dbFile = ews_grabber.Properties.Settings.Default.SQLiteDataBaseFilename;
            if (!dbFile.Contains('\\'))
            {  //build full file name
                dbFile = utils.helpers.getAppPath() + dbFile;
            }
            _licenseDataBase = new LicenseDataBase(ref this.dataGridView1, dbFile);
            if (_licenseDataBase.bIsValidDB == false)
            {
                MessageBox.Show("Error accessing or creating database file: " + dbFile, "Fatal ERROR");
                Application.Exit();
            }
            else
                loadData();
        }

        public bool addDataToGrid(
            int id,
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
            bool bRet = true;
            if (_bFilterActive)
                lblStatus.Text = "use manual refresh";
            else
                doRefresh();
            //dataGridView1.ResetBindings();
            //dataGridView1.DataSource = _licenseDataBase.getDataset();

            //dataGridView1.DataSource = _licenseDataBase.getDataset();
            //return bRet;
            //if (dataGridView1.DataSource != null)
            //{
            //    BindingSource bs = dataGridView1.DataSource as BindingSource;
            //    if (bs != null)
            //    {
            //        DataTable dt = bs.DataSource as DataTable;
            //        DataRow dr = dt.NewRow();
            //        dr["id"] = id;
            //        dr["deviceid"] = deviceid;//id     0
            //        dr["customer"] = customer;//user   1
            //        dr["key"] = key;     //       2
            //        dr["ordernumber"] = ordernumber; //   3
            //        dr["orderdate"] = orderdate;   //   4
            //        dr["ponumber"] = ponumber; //      5
            //        dr["endcustomer"] = endcustomer;//    6
            //        dr["product"] = product; //       7
            //        dr["quantity"] = quantity; //      8
            //        dr["receivedby"] = receivedby; //    9
            //        dr["sendat"] = sendat; //         10
            //        dr.AcceptChanges();
            //        dt.AcceptChanges();
            //        dataGridView1.Update();
            //    }
            //    else
                    //dataGridView1.Rows.Add(id, deviceid, customer, key, ordernumber, orderdate, ponumber, endcustomer, product, quantity, receivedby, sendat);
            //}

            //DataGridViewRow dgv = (DataGridViewRow)dataGridView1.Rows[0].Clone();
            return bRet;
        }
        /// <summary>
        /// called when a new license has been received
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        void on_new_licensemail(object sender, StatusEventArgs args)
        {
            LicenseData data=args.licenseData;
            //add data and refresh datagrid
            int InsertedID = _licenseDataBase.add(data._deviceid, data._customer, data._key, data._ordernumber, data._orderdate, data._ponumber, data._endcustomer, data._product, data._quantity, data._receivedby, data._sendat);
            //add data to datagrid
            //if (InsertedID!=-1)
            //    addDataToGrid(InsertedID, data._deviceid, data._customer, data._key, data._ordernumber, data._orderdate, data._ponumber, data._endcustomer, data._product, data._quantity, data._receivedby, data._sendat);
        }

        /// <summary>
        /// called for informational purposes by the Exchange Web Service class
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        void _ews_stateChanged1(object sender, StatusEventArgs args)
        {
            Cursor.Current = Cursors.Default;
            switch (args.eStatus)
            {
                case StatusType.success:
                    addLog("success " + args.strMessage);
                    break;
                case StatusType.validating:
                    addLog("validating " + args.strMessage);
                    break;
                case StatusType.error:
                    addLog("got invalid results");
                    break;
                case StatusType.busy:
                    addLog("exchange is busy..." + args.strMessage);
                    break;
                case StatusType.idle:
                    addLog("exchange idle...");
                    break;
                case StatusType.url_changed:
                    addLog("url changed: " + args.strMessage);
                    break;
                case StatusType.none:
                    addLog("wait..." + args.strMessage);
                    break;
                case StatusType.license_mail:
                    if(args.strMessage!=null)
                        addLog("license_mail: " + args.strMessage);
                    else
                        addLog("license_mail: " + args.licenseData._deviceid);
                    if (args.licenseData != null)
                    {
                        //add data to datagrid
                    }
                    break;
                case StatusType.bulk_mail:
                    addLog("bulk_mail: " + args.strMessage);
                    //a number of license data has been processed
                    //do a refresh
                    _licenseDataBase.doRefresh(ref this.dataGridView1);
                    break;
                case StatusType.other_mail:
                    addLog("other_mail: " + args.strMessage);
                    break;
            }
        }


        delegate void SetTextCallback(string text);
        public void addLog(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.textBox1.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(addLog);
                try
                {
                    this.Invoke(d, new object[] { text });
                }
                catch (Exception) { }
            }
            else
            {
                textBox1.Text += text + "\r\n";
                textBox1.SelectionLength = 0;
                textBox1.SelectionStart = textBox1.Text.Length - 1;
                textBox1.ScrollToCaret();
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(_ews!=null)
                _ews.Dispose();
        }

        private void mnuConnect_Click(object sender, EventArgs e)
        {
            _userData = new utils.UserData("Global", "E841719", "", false);
            Helpers.GetLogonData dlg = new Helpers.GetLogonData(ref _userData);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                //store settings for later use
                ews_grabber.Properties.Settings.Default.UseWebProxy = _userData.bUseProxy;
                ews_grabber.Properties.Settings.Default.ExchangeDomainname = _userData.sDomain;
                ews_grabber.Properties.Settings.Default.ExchangeUsername = _userData.sUser;
                ews_grabber.Properties.Settings.Default.Save();

                _ews = new ews();
                _ews.StateChanged += new StateChangedEventHandler(_ews_stateChanged1);
                _ews.start();
                if (_ews.logon(_userData.sDomain, _userData.sUser, _userData.sPassword,
                    ews_grabber.Properties.Settings.Default.UseWebProxy, 
                    ews_grabber.Properties.Settings.Default.ExchangeWebProxy, 
                    ews_grabber.Properties.Settings.Default.EchangeWebProxyPort))
                {// "Global", "E841719", ""))
                    Cursor.Current = Cursors.WaitCursor;
                    Application.DoEvents();
                    _ews.getMailsAsync();
                }
            }

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void mnuTest_xml_Click(object sender, EventArgs e)
        {
            addLog(LicenseXML.TestXML.runTest());
        }

        private void mnuTest_DB_Click(object sender, EventArgs e)
        {
            _licenseDataBase.add("cn70123", "customer", "key", "ordern#", DateTime.Now, "1234", "intermec", "CN70E", 1, "heinz-josef.gode@honeywel.com", DateTime.Now);
        }

        private void mnuProcess_Mail_Click(object sender, EventArgs e)
        {
            string sBody =
                "Hello Yolanda  Xie,\r\n" +
                "\r\n" +
                "Thank you for purchasing Naurtech software. Here are the registration keys for your software licenses. Please note that both the License ID and Registration keys are case sensitive. You can find instructions to manually register your license at:\r\n" +
                "http://www.naurtech.com/wiki/wiki.php?n=Main.TrainingVideoManualRegistration\r\n" +
                "\r\n" +
                "     Order Number:     15109\r\n" +
                "     Order Date:       4/10/2014\r\n" +
                "     Your PO Number:   PO93504\r\n" +
                "     End Customer:     Honeywell\r\n" +
                "     Product:          [NAU-1504] CETerm for Windows CE 6.0 / 5.0 / CE .NET\r\n" +
                "     Quantity:         46\r\n" +
                "\r\n" +
                "     Qty Ordered...............: 46\r\n" +
                "     Qty Shipped To Date.......: 46\r\n" +
                "\r\n" +
                "     Qty Shipped in this email.: 46\r\n" +
                "\r\n" +
                "\r\n" +
                "**** Registration Keys for Version 5.7 ***** \r\n" +
                "\r\n" +
                "\r\n" +
                "Version 5.1 and above support AUTOMATED LICENSE REGISTRATION. Please use the attached license file. This prevents you from having to type each key to register your copy of the software. Please refer to support wiki article http://www.naurtech.com/wiki/wiki.php?n=Main.AutoRegistration\r\n" +
                "";

            string testFile = utils.helpers.getAppPath() + "LicenseXMLFileSample.xml";
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            //read file into memorystream
            using (System.IO.FileStream file = new System.IO.FileStream(testFile, System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                byte[] bytes = new byte[file.Length];
                file.Read(bytes, 0, (int)file.Length);
                ms.Write(bytes, 0, (int)file.Length);
                ms.Flush();
            }

            Attachement[] atts = new Attachement[] { new Attachement(ms, "test.xml") };
            MailMessage msg = new MailMessage("E841719", sBody, "License Keys - Order: 15476: [NAU-1504] CETerm for Windows CE 6.0 / 5.0 / CE .NET", atts, DateTime.Now);
            
            this.Cursor = Cursors.WaitCursor;
            this.Enabled = false;
            Application.DoEvents();
            int i = _licenseMail.processMail(msg);
            this.Enabled = true; 
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            
            //lblStatus.Text = "Updated data. Use refresh";
            doRefresh();
        }

        void doRefresh()
        {
            DataGridView dgv = dataGridView1;
            int iRow = 0, iCol = 0;
            if (dgv.Rows.Count > 0)
            {
                iRow = dgv.CurrentCell.RowIndex;
                iCol = dgv.CurrentCell.ColumnIndex;
            } 
            dataGridView1.DataSource = _licenseDataBase.getDataset();
            if (dgv.Rows.Count > 0)
            {
                dgv.CurrentCell = dgv.Rows[iRow].Cells[iCol];
            }
        }

        void loadData()
        {
            dataGridView1.DataSource = _licenseDataBase.getDataset();//.Tables[0].DefaultView;
        }

        private void mnuClearData_Click(object sender, EventArgs e)
        {
            _licenseDataBase.clearData();
            doRefresh();
        }

        private void mnuRefresh_Click(object sender, EventArgs e)
        {
            doRefresh();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            _licenseDataBase.filterData(ref this.dataGridView1, _filterType, txtFilter.Text);
            if (_filterType != LicenseDataBase.FilterType.none || txtFilter.Text.Length == 0)
                _bFilterActive = true;
            else
                _bFilterActive = false;
        }

        LicenseDataBase.FilterType _filterType
        {
            get { 
                    LicenseDataBase.FilterType filterType;
                    if (radioDeviceID.Checked)
                        filterType = LicenseDataBase.FilterType.DeviceID;
                    else if (radioKeyNumber.Checked)
                        filterType = LicenseDataBase.FilterType.KeyID;
                    else if (radioOrderNumber.Checked)
                        filterType = LicenseDataBase.FilterType.OrderNumber;
                    else if (radioPOnumber.Checked)
                        filterType = LicenseDataBase.FilterType.PurchaseOrder;
                    else
                        filterType = LicenseDataBase.FilterType.none;
                    return filterType;
                }
        }

        private void btnAll_Click(object sender, EventArgs e)
        {
            _licenseDataBase.filterData(ref this.dataGridView1, LicenseDataBase.FilterType.none, "");
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection rows = dataGridView1.SelectedRows;
            if (rows.Count > 0)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.DefaultExt = "xml";
                sfd.CheckPathExists = true;
                sfd.Filter = "*.xml|xml files|*.*|all files";
                sfd.FileName = "license.xml";
                sfd.FilterIndex = 0;
                sfd.OverwritePrompt = true;
                sfd.RestoreDirectory = true;
                sfd.Title="Enter name of file to export these "+rows.Count.ToString()+" license(s)";
                sfd.ValidateNames=true;
                string sFileName="";
                if(sfd.ShowDialog()==System.Windows.Forms.DialogResult.OK){
                    sFileName=sfd.FileName;
                }
                else{
                    return;
                }
                LicenseXML licenseXML = new LicenseXML();
                List<Helpers.license> licenseList = new List<license>();
                foreach (DataGridViewRow drv in rows)
                {
                    DataRow dr = (drv.DataBoundItem as DataRowView).Row;
                    Helpers.license lic = new license();
                    lic.id = dr["deviceid"].ToString();
                    lic.key = dr["key"].ToString();
                    lic.user = dr["endcustomer"].ToString();
                    licenseList.Add(lic);
                }
                licenseXML.licenses = licenseList.ToArray();
                if (LicenseXML.serialize(licenseXML, sFileName))
                {
                    MessageBox.Show("license(s) saved to " + sFileName);
                }
            }
        }

        private void mnuSettings_Click(object sender, EventArgs e)
        {
            SettingsForm dlg = new SettingsForm();
            dlg.ShowDialog();
        }
    }
}
