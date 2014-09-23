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
    public partial class MainForm : Form
    {
        ews _ews;
        utils.UserData _userData;
        //data and database
        LicenseDataBase _licenseDataBase;
        LicenseMail _licenseMail = new LicenseMail();
        bool _bFilterActive = false;
        MySettings _mysettings = new MySettings();
        utils.DataGridViewSettings dgvSettings = new utils.DataGridViewSettings();

        public MainForm()
        {
            InitializeComponent();
            #if !DEBUG
            mnuAdmin.Visible = false;
            #endif
            OpenDB:
            //load settings
            _mysettings = new MySettings();
            _mysettings = _mysettings.load();

            string appPath = utils.helpers.getAppPath();

            _licenseMail = new LicenseMail();
            //subscribe to new license mails
            _licenseMail.StateChanged += on_new_licensemail;

            addLog("Please select Exchange-Connect to start test");
            string dbFile = _mysettings.SQLiteDataBaseFilename;

            if (!dbFile.Contains('\\'))
            {  //build full file name
                dbFile = utils.helpers.getAppPath() + dbFile;
            }
            _licenseDataBase = new LicenseDataBase(ref this.dataGridView1, dbFile);

            if (_licenseDataBase.bIsValidDB == false)
            {
                MessageBox.Show("Error accessing or creating database file: " + dbFile, "Fatal ERROR");
                SettingsForm dlg = new SettingsForm();
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    goto OpenDB;    //try again
                }
                else
                {
                    MessageBox.Show("Cannot run without valid database path and file name", "Exchange License Mail Grabber");
                    exitApp();
                }
            }
            else
            {
                _licenseDataBase.StateChanged += new StateChangedEventHandler(_licenseDataBase_StateChanged);
                loadData();
            }

            readDGVlayout();
        }

        void readDGVlayout()
        {
            dgvSettings = utils.DataGridViewSettings.Deserialize<utils.DataGridViewSettings>(utils.helpers.getAppPath() + utils.DataGridViewSettings.settingsFileConst);
            //dgvSettings=dgvSettings.load();
            if (dgvSettings == null)
                return;
            for (int x = 0; x < dataGridView1.Columns.Count; x++ )
            {
                dataGridView1.Columns[x].Visible = dgvSettings.columns[x].visible;
                dataGridView1.Columns[x].Width = dgvSettings.columns[x].width;
                dataGridView1.Columns[x].HeaderText = dgvSettings.columns[x].header;
            }
        }

        void saveDGVlayout()
        {
            for (int x = 0; x < dataGridView1.Columns.Count; x++ )
            {
                 dgvSettings.columns[x].visible=dataGridView1.Columns[x].Visible   ;
                 dgvSettings.columns[x].width = dataGridView1.Columns[x].Width;
                 dgvSettings.columns[x].header = dataGridView1.Columns[x].HeaderText;
            }
            utils.DataGridViewSettings.Serialize<utils.DataGridViewSettings>(dgvSettings, utils.helpers.getAppPath() + utils.DataGridViewSettings.settingsFileConst);
//            dgvSettings.Save(dgvSettings);
        }

        void exitApp()
        {
            if (System.Windows.Forms.Application.MessageLoop)
            {
                // WinForms app
                System.Windows.Forms.Application.Exit();
            }
            else
            {
                // Console app
                System.Environment.Exit(1);
            }
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
                setLblStatus("use manual refresh");
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
            if (args.eStatus == StatusType.license_mail)
            {
                utils.helpers.addLog("received new License Data\r\n");
                LicenseData data = args.licenseData;
                //add data and refresh datagrid
                int InsertedID = _licenseDataBase.add(data._deviceid, data._customer, data._key, data._ordernumber, data._orderdate, data._ponumber, data._endcustomer, data._product, data._quantity, data._receivedby, data._sendat);
                //add data to datagrid
                //if (InsertedID!=-1)
                //    addDataToGrid(InsertedID, data._deviceid, data._customer, data._key, data._ordernumber, data._orderdate, data._ponumber, data._endcustomer, data._product, data._quantity, data._receivedby, data._sendat);
                setLblStatus("new data");
            }
            else
            {
                _ews_stateChanged1(sender, args);
            }
        }

        void _licenseDataBase_StateChanged(object s, StatusEventArgs args)
        {
            _ews_stateChanged1(s, args);
        }

        static bool _bLEDToggle = false;
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
                case StatusType.ews_pulse:
                    if (_bLEDToggle)
                        setLEDColor(Color.Green);
                    else
                        setLEDColor(Color.LightGreen);
                    _bLEDToggle = !_bLEDToggle;
                    break;
                case StatusType.ews_started:
                    lblLED.BackColor = Color.LightGreen;
                    mnuDisconnect.Enabled = true;
                    mnuConnect.Enabled = false;
                    break;
                case StatusType.ews_stopped:
                    lblLED.BackColor = Color.Red;
                    break;
                case StatusType.success:
                    addLog("success " + args.strMessage);
                    setLblStatus("success");
                    break;
                case StatusType.validating:
                    addLog("validating " + args.strMessage);
                    setLblStatus("validating");
                    break;
                case StatusType.error:
                    addLog("got invalid results");
                    if (args.strMessage != null)
                        setLblStatus(args.strMessage);
                    else
                        setLblStatus("error");
                    break;
                case StatusType.busy:
                    addLog("exchange is busy..." + args.strMessage);
                    setLblStatus("busy");
                    break;
                case StatusType.idle:
                    addLog("exchange idle...");
                    setLblStatus("idle");
                    break;
                case StatusType.url_changed:
                    addLog("url changed: " + args.strMessage);
                    setLblStatus(args.strMessage);
                    break;
                case StatusType.none:
                    addLog("wait..." + args.strMessage);
                    setLblStatus(args.strMessage);
                    break;
                case StatusType.license_mail:
                    if (args.strMessage != null)
                    {
                        addLog("license_mail: " + args.strMessage);
                        setLblStatus(args.strMessage);
                    }
                    else
                        addLog("license_mail: " + args.licenseData._deviceid);
                    if (args.licenseData != null)
                    {
                        //add data to datagrid
                        setLblStatus(args.licenseData._deviceid);
                    }
                    break;
                case StatusType.other_mail:
                    addLog("other_mail: " + args.strMessage);
                    setLblStatus(args.strMessage);
                    break;
            }
        }

        delegate void setLEDColorD(Color color);
        void setLEDColor(Color color)
        {
            if (lblLED.InvokeRequired)
            {
                setLEDColorD d = new setLEDColorD(setLEDColor);
                try
                {
                    this.Invoke(d, new object[] { color });
                }
                catch (Exception)
                {}
            }
            else {
                lblLED.ForeColor = color;
            }
        }

        delegate void setLblStatusD(string s);
        void setLblStatus(string s)
        {
            if (this.lblStatus.InvokeRequired)
            {
                setLblStatusD d = new setLblStatusD(setLblStatus);
                try
                {
                    this.Invoke(d, new object[] { s });
                }
                catch (Exception)
                { }
            }
            else
            {
                lblStatus.Text = s;
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
                if (textBox1.Text.Length > 30000)
                    textBox1.Text = "";
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
            saveDGVlayout();
        }

        private void mnuConnect_Click(object sender, EventArgs e)
        {
            _mysettings = _mysettings.load();
            _userData = new utils.UserData(_mysettings.ExchangeDomainname, _mysettings.ExchangeUsername, "", _mysettings.UseWebProxy);
            Helpers.GetLogonData dlg = new Helpers.GetLogonData(ref _userData);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                _ews = new ews(ref _licenseMail);
                _ews.StateChanged += new StateChangedEventHandler(_ews_stateChanged1);
                _ews.start();

                if (_ews.logon(_userData.sDomain, _userData.sUser, _userData.sPassword,
                    _mysettings.UseWebProxy,
                    _mysettings.ExchangeWebProxy,
                    _mysettings.EchangeWebProxyPort))
                {// "Global", "E841719", ""))
                    Cursor.Current = Cursors.WaitCursor;
                    Application.DoEvents();
                    _ews.getMailsAsync();
                }
            }
            mnuConnect.Enabled = false;
            mnuDisconnect.Enabled = true;
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
            
            //setLblStatus("Updated data. Use refresh";
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
            //dataGridView1.Columns[0].HeaderText = "ID";
            for (int x=0; x<dataGridView1.Columns.Count; x++)
                dataGridView1.Columns[x].HeaderText = Helpers.licenseCols.DataGridColHeaders[x];
            dataGridView1.Columns[0].Visible=false;
            if (dgv.Rows.Count > 0 && iCol>0)
            {
                dgv.CurrentCell = dgv.Rows[iRow].Cells[iCol];
            }
        }

        void loadData()
        {
            doRefresh();
            //dataGridView1.DataSource = _licenseDataBase.getDataset();//.Tables[0].DefaultView;
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
            if (dataGridView1.RowCount > 0)
            {
                _licenseDataBase.filterData(ref this.dataGridView1, _filterType, txtFilter.Text);
                if (_filterType != LicenseDataBase.FilterType.none || txtFilter.Text.Length == 0)
                    _bFilterActive = true;
                else
                    _bFilterActive = false;
            }
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
            try
            {
                _licenseDataBase.filterData(ref this.dataGridView1, LicenseDataBase.FilterType.none, "");
            }
            catch (Exception)
            {
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 0)
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
                    sfd.Title = "Enter name of file to export these " + rows.Count.ToString() + " license(s)";
                    sfd.ValidateNames = true;
                    string sFileName = "";
                    if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        sFileName = sfd.FileName;
                    }
                    else
                    {
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
        }

        private void mnuSettings_Click(object sender, EventArgs e)
        {
            SettingsForm dlg = new SettingsForm();
            dlg.ShowDialog();
        }

        private void mnuDisconnect_Click(object sender, EventArgs e)
        {
            if (_ews != null)
            {
                _ews.Dispose();
                _ews = null;
            }
            mnuConnect.Enabled = true;
            mnuDisconnect.Enabled = false;
        }

        private void mnuExport_Click(object sender, EventArgs e)
        {
            btnExport_Click(this, e);
        }

        private void mnuAbout_Click(object sender, EventArgs e)
        {
            AboutBox1 dlg = new AboutBox1();
            dlg.ShowDialog();
            dlg.Dispose();
        }
    }
}
