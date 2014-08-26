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
        
        public SettingsForm()
        {
            InitializeComponent();
            loadSettings();
        }

        void loadSettings()
        {
            ews_grabber.Properties.Settings.Default.Upgrade();
            ews_grabber.Properties.Settings.Default.Reload();
            txtDomain.Text = ews_grabber.Properties.Settings.Default.ExchangeDomainname;
            txtUser.Text = ews_grabber.Properties.Settings.Default.ExchangeUsername;
            txtExchangeServiceURL.Text = ews_grabber.Properties.Settings.Default.ExchangeServiceURL;
            chkUseWebProxy.Checked = ews_grabber.Properties.Settings.Default.UseWebProxy;
            txtWebProxy.Text = ews_grabber.Properties.Settings.Default.ExchangeWebProxy;
            numProxyPort.Value = ews_grabber.Properties.Settings.Default.EchangeWebProxyPort;

            txtDatabaseFile.Text = ews_grabber.Properties.Settings.Default.SQLiteDataBaseFilename;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            ews_grabber.Properties.Settings.Default.ExchangeDomainname = txtDomain.Text;
            ews_grabber.Properties.Settings.Default.ExchangeUsername = txtUser.Text;
            ews_grabber.Properties.Settings.Default.ExchangeServiceURL = txtExchangeServiceURL.Text;
            ews_grabber.Properties.Settings.Default.ExchangeWebProxy = txtWebProxy.Text;
            ews_grabber.Properties.Settings.Default.EchangeWebProxyPort = (int)numProxyPort.Value;
            ews_grabber.Properties.Settings.Default.UseWebProxy = chkUseWebProxy.Checked;

            string db_file = ews_grabber.Properties.Settings.Default.SQLiteDataBaseFilename;
            if (db_file != txtDatabaseFile.Text)
            {
                MessageBox.Show("Changing the db file needs a program restart!", "database file changed");
            }
            ews_grabber.Properties.Settings.Default.SQLiteDataBaseFilename = txtDatabaseFile.Text;

            ews_grabber.Properties.Settings.Default.Save();

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
            string db_file = ews_grabber.Properties.Settings.Default.SQLiteDataBaseFilename;
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
