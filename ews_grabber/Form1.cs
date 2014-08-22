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
        //data and database
        LicenseData _licenseDataBase;

        public Form1()
        {
            InitializeComponent();
            string appPath = utils.helpers.getAppPath();
            addLog("Please select Exchange-Connect to start test");
            _licenseDataBase = new LicenseData();
            loadData();
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
                    if(args.strMessage!=null)
                        addLog("license_mail: " + args.strMessage);
                    else
                        addLog("license_mail: " + args.mailmsg.Subject);
                    if (args.mailmsg != null)
                    {
                        if (Helpers.LicenseMail.processMail(args.mailmsg, ref _licenseDataBase) > 0)
                            dataGridView1.Refresh();
                    }
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
            _userData = new utils.userData("Global", "E841719", "", false);
            Helpers.GetLogonData dlg = new Helpers.GetLogonData(ref _userData);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                _ews = new ews();
                _ews.StateChanged += new StateChangedEventHandler(_ews_stateChanged1);
                _ews.start();
                if (_ews.logon(_userData.sDomain, _userData.sUser, _userData.sPassword, _userData.bUseProxy))
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
            LicenseData ld = new LicenseData();
            ld.add("cn70123", "customer", "key", "ordern#", DateTime.Now, "1234", "intermec", "CN70E", 1, "heinz-josef.gode@honeywel.com", DateTime.Now);

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
            int i = LicenseMail.processMail(msg, ref _licenseDataBase);
            dataGridView1.Refresh();

        }

        void loadData()
        {
            LicenseData licenseData = new LicenseData();

            dataGridView1.DataSource = licenseData.getDataset().Tables[0].DefaultView;
            dataGridView1.Refresh();
        }
    }
}
