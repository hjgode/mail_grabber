using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office;
using Microsoft.Office.Interop.Outlook;

namespace outlookGrabber
{
    class OutlookClass:IDisposable
    {
        Microsoft.Office.Interop.Outlook.Application MyApp = null;
        Microsoft.Office.Interop.Outlook.NameSpace MailNS = null;
        Microsoft.Office.Interop.Outlook.MAPIFolder MyInbox = null;

        public string _sPassword { get; set; }
        public string _sProfile { get; set; }

        public OutlookClass()
        {
            try
            {
                //start outlook
                MyApp = new Application();
                MailNS = MyApp.GetNamespace("MAPI");

                string sProfile = "", sPassword="";
                
                bool bShowDialog = true, bNewSession=false;

                MailNS.Logon(sProfile, sPassword, bShowDialog, bNewSession);

                MyInbox = MailNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                

            }
            catch (System.Exception ex)
            {
                helpers.addLog("Exception: " + ex.Message);
            }
        }

        public string getMails()
        {
            StringBuilder sb = new StringBuilder();
            Microsoft.Office.Interop.Outlook.Items oItems = MyInbox.Items;

            if (MyInbox.Items.Count > 0)
            {
                int i = 1;
                for (i = 1; i < MyInbox.Items.Count; i++)
                {
                    try
                    {
                        sb.Append("\r\n===============================================\r\n");
                        MailItem mi = (MailItem)oItems[i];
                        sb.Append(mi.SenderEmailAddress + "\r\n");
                        sb.Append(mi.Subject + "\r\n");
                        //sb.Append(mi.Body);
                        if (mi.Attachments.Count > 0)
                        {
                            sb.Append("----------------------------------------------\r\n");
                            foreach (Attachment att in mi.Attachments)
                            {
                                sb.Append(att.FileName+"\r\n");
                            }
                            sb.Append("----------------------------------------------\r\n");
                        }
                        sb.Append("===============================================\r\n");

                    }
                    catch (System.Exception ex)
                    {
                        sb.Append("Exception: " + ex.Message + " for " + i.ToString() + "'" + oItems[i].ToString());
                    }
                }
            }
            helpers.addLog(sb.ToString());
            return sb.ToString();
        }
        
        public void Dispose()
        {
            helpers.addLog("Disposing...");
            MailNS.Logoff();
            MailNS = null;
            MyInbox = null;
            //InboxMailItem = null;
            ((Microsoft.Office.Interop.Outlook._Application)MyApp).Quit();
            ////MyApp.Quit();
            MyApp = null;
            helpers.addLog("Dispose done.");
        }
    }
}
