using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ews_grabber
{
    public partial class SettingsForm : Form
    {
        MySettings _mysettings = new MySettings();
        public SettingsForm()
        {
            InitializeComponent();
            loadSettings();
        }

        void loadSettings()
        {
            _mysettings = _mysettings.load();            
            txtDomain.Text = _mysettings.ExchangeDomainname;
            txtUser.Text = _mysettings.ExchangeUsername;
            txtExchangeServiceURL.Text = _mysettings.ExchangeServiceURL;
            chkUseWebProxy.Checked = _mysettings.UseWebProxy;
            txtWebProxy.Text = _mysettings.ExchangeWebProxy;
            numProxyPort.Value = _mysettings.EchangeWebProxyPort;

            txtDatabaseFile.Text = _mysettings.SQLiteDataBaseFilename;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            _mysettings.ExchangeDomainname = txtDomain.Text;
            _mysettings.ExchangeUsername = txtUser.Text;
            _mysettings.ExchangeServiceURL = txtExchangeServiceURL.Text;
            _mysettings.ExchangeWebProxy = txtWebProxy.Text;
            _mysettings.EchangeWebProxyPort = (int)numProxyPort.Value;
            _mysettings.UseWebProxy = chkUseWebProxy.Checked;

            string db_file = txtDatabaseFile.Text;
            if (db_file != _mysettings.SQLiteDataBaseFilename)
            {
                MessageBox.Show("Changing the db file needs a program restart!", "database file changed");
            }
            _mysettings.SQLiteDataBaseFilename = db_file;

            _mysettings.Save(_mysettings);

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void SettingsForm_Activated(object sender, EventArgs e)
        {
            loadSettings();
        }

        private void btnGetFile_Click(object sender, EventArgs e)
        {
            string db_file = _mysettings.SQLiteDataBaseFilename;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.CheckPathExists = true;
            ofd.DereferenceLinks = true;
            ofd.Filter = "*.db|db files|*.*|all files";
            ofd.FilterIndex = 0;
            ofd.Multiselect = false;
            ofd.RestoreDirectory = true;
            ofd.CheckFileExists = false;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtDatabaseFile.Text = ofd.FileName;
            }
            ofd.Dispose();
        }
    }
}
