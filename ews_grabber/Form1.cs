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
        utils.userData _userData;
        public Form1()
        {
            InitializeComponent();
            string appPath = utils.helpers.getAppPath();
            addLog("Please select Exchange-Connect to start test");
        }

        List<Microsoft.Exchange.WebServices.Data.EmailMessage> mailList = new List<Microsoft.Exchange.WebServices.Data.EmailMessage>();

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
                    addLog("license_mail: " + args.strMessage);
                    Helpers.LicenseMail.processMail(args.mailmsg);
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
            _userData = new utils.userData("Global", "E841719", "");
            Helpers.GetLogonData dlg = new Helpers.GetLogonData(ref _userData);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                _ews = new ews();
                _ews.StateChanged += new StateChangedEventHandler(_ews_stateChanged1);
                _ews.start();
                if (_ews.logon(_userData.sDomain, _userData.sUser, _userData.sPassword))
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
    }
}
