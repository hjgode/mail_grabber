using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.Exchange.WebServices.Data;

namespace ews_grabber
{
    public partial class Form1 : Form
    {
        ews _ews;
        public Form1()
        {
            InitializeComponent();
            _ews = new ews();
            _ews.MailRead += _ews_MailRead;
            if (_ews.logon("Global", "E841719", "Chopper+8"))
            {// if(_ews.logon())// if(_ews.logon("heinz-josef.gode@honeywell.com"))
                //_ews.getMails();
                Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();
                _ews.getMailsAsync();
            }

        }

        void _ews_MailRead(ews.ReadMailEventArgs args)
        {
            Cursor.Current = Cursors.Default;
            if (args.bValid)
                addLog("got valid results");
            else
                addLog("got invalid results");
            List<EmailMessage> list = _ews._mailList;
            foreach (EmailMessage m in list)
            {
                addLog( m.Sender.Name + m.Subject + m.Attachments.Count.ToString() + "\r\n");
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
                this.Invoke(d, new object[] { text });
            }
            else
            {
                textBox1.Text += text + "\r\n";
                textBox1.SelectionLength = 0;
                textBox1.SelectionStart = textBox1.Text.Length - 1;
                textBox1.ScrollToCaret();
            }
        }
    }
}
