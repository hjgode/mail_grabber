using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Helpers;
using OutlookMail;

namespace outlk_grabber
{
    public partial class outlookMain : Form
    {
        utils.UserData _userData;
        MySettings _mysettings = new MySettings();
        LicenseMail _licenseMail = new LicenseMail();
        bool _bFilterActive = false;
        OutlookMail.OutlookMail _olk;
        static bool _bLEDToggle = false;
        //data and database
        LicenseDataBase _licenseDataBase;

        public outlookMain()
        {
            InitializeComponent();
            //subscribe to new license mails
            _licenseMail.StateChanged += on_new_licensemail;

             _olk = new OutlookMail.OutlookMail(ref _licenseMail);

             string dbFile = _mysettings.SQLiteDataBaseFilename;

             if (!dbFile.Contains('\\'))
             {  //build full file name
                 dbFile = utils.helpers.getAppPath() + dbFile;
             }
             _licenseDataBase = new LicenseDataBase(ref this.dataGridView1, dbFile);
             _licenseDataBase.StateChanged += new StateChangedEventHandler(_licenseDataBase_StateChanged);
             if (_licenseDataBase.bIsValidDB == false)
             {
                 MessageBox.Show("Error accessing or creating database file: " + dbFile, "Fatal ERROR");
                 Application.Exit();
             }
             else
                 loadData();

        }

        void _licenseDataBase_StateChanged(object s, StatusEventArgs args)
        {
            _olk_stateChanged(s, args);
        }

        void loadData()
        {
            dataGridView1.DataSource = _licenseDataBase.getDataset();//.Tables[0].DefaultView;
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
                _olk_stateChanged(sender, args);
            }
        }

        void _olk_stateChanged(object sender, StatusEventArgs args)
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
        private void mnuConnect_Click(object sender, EventArgs e)
        {
            _mysettings = _mysettings.load();
            _userData = new utils.UserData(_mysettings.ShowDialog,_mysettings.OutlookProfile, _mysettings.Username, "");
            Helpers.GetLogonDataOutlook dlg = new Helpers.GetLogonDataOutlook(ref _userData);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                _olk = new OutlookMail.OutlookMail(ref _licenseMail);
                _olk.StateChanged += new StateChangedEventHandler(_olk_stateChanged);

                //_olk.StateChanged += new StateChangedEventHandler(_olk_StateChanged);
                _olk.start();

                if (_olk.logon(_userData.sOutlookProfile, _userData.sUser, _userData.sPassword, _userData.bShowOutlookDialog))
                {// "Global", "E841719", ""))
                    Cursor.Current = Cursors.WaitCursor;
                    Application.DoEvents();
                    _olk.getMailsAsync();
                }
            }
            mnuConnect.Enabled = false;
            mnuDisconnect.Enabled = true;

        }

        private void mnuDisconnect_Click(object sender, EventArgs e)
        {
            if (_olk != null)
            {
                _olk.Dispose();
                _olk = null;
            }
            mnuConnect.Enabled = true;
            mnuDisconnect.Enabled = false;
        }

        private void mnuSettings_Click(object sender, EventArgs e)
        {
            SettingsForm dlg = new SettingsForm();
            dlg.ShowDialog();
            dlg.Dispose();
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
                { }
            }
            else
            {
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

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_olk != null)
                _olk.Dispose();
            Application.Exit();
        }

        private void mnuRefresh_Click(object sender, EventArgs e)
        {
            doRefresh();
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
            get
            {
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
    }
}
