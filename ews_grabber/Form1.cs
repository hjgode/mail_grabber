using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using ExchangeMail;

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
                    addLog("got valid results");
                    List<Microsoft.Exchange.WebServices.Data.EmailMessage> list = _ews._mailList;
                    foreach (Microsoft.Exchange.WebServices.Data.EmailMessage m in list)
                    {
                        processMail(m);
                    }
                    mailList.AddRange(list);
                    break;
                case StatusType.validating:
                    addLog("validating "+args.message);
                    break;
                case StatusType.error:
                    addLog("got invalid results");
                    break;
                case StatusType.busy:
                    addLog("exchange is busy..." + args.message);
                    break;
                case StatusType.idle:
                    addLog("exchange idle...");
                    break;
                case StatusType.url_changed:
                    addLog("url changed: " + args.message);
                    break;
                case StatusType.none:
                    addLog("wait..."+args.message);
                    break;
                case StatusType.license_mail:
                    addLog(args.message);
                    processMail(args.excMail);
                    break;
            }
        }

        int processMail(Microsoft.Exchange.WebServices.Data.EmailMessage m)
        {
            int iRet = 0;
            if (m == null)
            {
                addLog("processMail: null msg");
                return iRet;
            }
            try
            {
                addLog(m.Sender.Name + m.Subject + m.Attachments.Count.ToString() + "\r\n");
                if (m.HasAttachments)
                {
                    // Request all the attachments on the email message. This results in a GetItem operation call to EWS.
                    m.Load(new Microsoft.Exchange.WebServices.Data.PropertySet(Microsoft.Exchange.WebServices.Data.EmailMessageSchema.Attachments));
                    foreach (Microsoft.Exchange.WebServices.Data.Attachment att in m.Attachments)
                    {
                        if (att is Microsoft.Exchange.WebServices.Data.FileAttachment)
                        {
                            Microsoft.Exchange.WebServices.Data.FileAttachment fileAttachment = att as Microsoft.Exchange.WebServices.Data.FileAttachment;
                            
                            //get a temp file name
                            string fname = System.IO.Path.GetTempFileName(); //utils.helpers.getAppPath() + fileAttachment.Id.ToString() + "_" + fileAttachment.Name

                            /*
                            // Load the file attachment into memory. This gives you access to the attachment content, which 
                            // is a byte array that you can use to attach this file to another item. This results in a GetAttachment operation
                            // call to EWS.
                            fileAttachment.Load();
                            */

                            // Load attachment contents into a file. This results in a GetAttachment operation call to EWS.
                            fileAttachment.Load(fname);
                            addLog("Attachement file saved to: " + fname);

                            /*
                            // Put attachment contents into a stream.
                            using (System.IO.FileStream theStream =
                                new System.IO.FileStream(utils.helpers.getAppPath() + fileAttachment.Id.ToString() + "_" + fileAttachment.Name, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite))
                            {
                                //This results in a GetAttachment operation call to EWS.
                                fileAttachment.Load(theStream);
                            }
                            */

                            /*
                            //load into memory stream, seems the only stream supported
                            using (System.IO.MemoryStream ms = new System.IO.MemoryStream(att.Size))
                            {
                                fileAttachment.Load(ms);
                                using (System.IO.FileStream fs = new System.IO.FileStream(fname, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite))
                                {
                                    ms.CopyTo(fs);
                                    fs.Flush();
                                }                                
                            }
                            */
                            addLog("saved attachement: " + fname);
                            iRet++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                addLog("Exception: " + ex.Message);
            }
            addLog("processMail did process " + iRet.ToString() + " files");
            return iRet;
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
            utils.userData _userData = new utils.userData("Global", "E841719", "");
            Helpers.GetLogonData dlg = new Helpers.GetLogonData(ref _userData);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                _ews = new ews();
                _ews.stateChangedEvent += new StateChangedEventHandler(_ews_stateChanged1);
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
    }
}
