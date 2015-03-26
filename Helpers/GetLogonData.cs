using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Helpers
{
    public partial class GetLogonData : Form
    {
        utils.UserData _userData;
        public GetLogonData(ref utils.UserData data)
        {
            _userData = data;
            InitializeComponent();
            txtDomain.Text = data.sDomain;
            txtUser.Text = data.sUser;
            txtPassword.Text = data.sPassword;
            chkUseProxy.Checked = data.bUseProxy;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            _userData.sDomain = txtDomain.Text;
            _userData.sUser = txtUser.Text;
            _userData.sPassword = txtPassword.Text;
            _userData.bUseProxy = chkUseProxy.Checked;
            this.DialogResult= DialogResult.OK;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void GetLogonData_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\r')
                this.btnOK_Click(this, new EventArgs());
        }
    }
}
