using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Outlook.Extensions;

using System.IO;

namespace outlookGrabber
{
    class OutlookClass:IDisposable
    {
        Microsoft.Office.Interop.Outlook.Application MyApp = null;
        Microsoft.Office.Interop.Outlook.NameSpace MailNS = null;
        Microsoft.Office.Interop.Outlook.MAPIFolder MyInbox = null;

        public string _sPassword { get; set; }
        public string _sProfile { get; set; }

        const string sLicenseSubject = "Your license order";

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

        public string findLicenseMails()
        {
            StringBuilder sb = new StringBuilder();

            Microsoft.Office.Interop.Outlook._MailItem InboxMailItem = null;
            Microsoft.Office.Interop.Outlook.Items oItems = MyInbox.Items; 
            string Query = "[Subject] contains 'Your license order'";
            //string filter = "urn:schemas:mailheader:subject LIKE '%" + wordInSubject + "%'";

            InboxMailItem = (Microsoft.Office.Interop.Outlook._MailItem)oItems.Find(Query);
            while (InboxMailItem != null)
            {
                //ListViewItem myItem = lvwMails.Items.Add(InboxMailItem.SenderName);
                //myItem.SubItems.Add(InboxMailItem.Subject);
                sb.Append(InboxMailItem.SenderName + ": ");
                sb.Append(InboxMailItem.Subject + "\r\n");
                string AttachmentNames = string.Empty;
                foreach (Microsoft.Office.Interop.Outlook.Attachment item in InboxMailItem.Attachments)
                {
                    AttachmentNames += item.DisplayName;
                    sb.Append("\t" + item.DisplayName + "\r\n");
                    //item.SaveAsFile(this.GetAttachmentPath(item.FileName, "c:\\test\\attachments\\"));
                }
                //myItem.SubItems.Add(AttachmentNames);
                //InboxMailItem.Delete();
                InboxMailItem = (Microsoft.Office.Interop.Outlook._MailItem)oItems.FindNext();
            }
            return sb.ToString();
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

        #region TESTING_AREA
        //####################################### TESTING ###########################################
        /*
        public void getLicenseMails() //see http://blogs.msdn.com/b/philliphoff/archive/2008/03/03/filter-outlook-items-by-date-with-linq-to-dasl.aspx
        {
            Microsoft.Office.Interop.Outlook.Folder folder = (Microsoft.Office.Interop.Outlook.Folder)MailNS.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);

            string subject = sLicenseSubject;// "Your license order";

            var results =

                    from item in folder.Items.AsQueryable<Microsoft.Office.Interop.Outlook.Extensions.Linq.Mail>()

                    where item.Subject.Contains(subject) && item.Attachments.Count>0 //item.CreationTime <= DateTime.Now - new TimeSpan(7, 0, 0, 0)

                    select item;

            foreach (var result in results)
            {
                System.Diagnostics.Debug.WriteLine(String.Format("Body: {0}", result));
            }
        }
        */

        private void SearchRecurringAppointments()
        {
            Microsoft.Office.Interop.Outlook.AppointmentItem appt = null;
            Microsoft.Office.Interop.Outlook.Folder folder = MailNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
                //Microsoft.Office.Interop.Outlook.Application.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar)
                as Microsoft.Office.Interop.Outlook.Folder;
            // Set start value
            DateTime start =
                new DateTime(2006, 8, 9, 0, 0, 0);
            // Set end value
            DateTime end =
                new DateTime(2006, 12, 14, 0, 0, 0);
            // Initial restriction is Jet query for date range
            string filter1 = "[Start] >= '" +
                start.ToString("g")
                + "' AND [End] <= '" +
                end.ToString("g") + "'";
            Microsoft.Office.Interop.Outlook.Items calendarItems = folder.Items.Restrict(filter1);
            calendarItems.IncludeRecurrences = true;
            calendarItems.Sort("[Start]", Type.Missing);
            // Must use 'like' comparison for Find/FindNext
            string filter2;
            filter2 = "@SQL="
                + "\"" + "urn:schemas:httpmail:subject" + "\""
                + " like '%Office%'";
            // Create DASL query for additional Restrict method
            string filter3;
            if (MailNS.DefaultStore.IsInstantSearchEnabled)// Microsoft.Office.Interop.Outlook.Application.Session.DefaultStore.IsInstantSearchEnabled)
            {
                filter3 = "@SQL="
                    + "\"" + "urn:schemas:httpmail:subject" + "\""
                    + " ci_startswith 'Office'";
            }
            else
            {
                filter3 = "@SQL="
                    + "\"" + "urn:schemas:httpmail:subject" + "\""
                    + " like '%Office%'";
            }
            // Use Find and FindNext methods
            appt = calendarItems.Find(filter2)
                as Microsoft.Office.Interop.Outlook.AppointmentItem;
            while (appt != null)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(appt.Subject);
                sb.AppendLine("Start: " + appt.Start);
                sb.AppendLine("End: " + appt.End);
                System.Diagnostics.Debug.WriteLine(sb.ToString());
                // Find the next appointment
                appt = calendarItems.FindNext()
                    as Microsoft.Office.Interop.Outlook.AppointmentItem;
            }
            // Restrict calendarItems with DASL query
            Microsoft.Office.Interop.Outlook.Items restrictedItems =
                calendarItems.Restrict(filter3);
            foreach (Microsoft.Office.Interop.Outlook.AppointmentItem apptItem in restrictedItems)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(apptItem.Subject);
                sb.AppendLine("Start: " + apptItem.Start);
                sb.AppendLine("End: " + apptItem.End);
                sb.AppendLine();
                System.Diagnostics.Debug.WriteLine(sb.ToString());
            }
        }

        private void DemoTableColumns()
        {
            const string PR_HAS_ATTACH =
                "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B";
            // Obtain Inbox
            Microsoft.Office.Interop.Outlook.Folder folder = MailNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox) as Folder;//Microsoft.Office.Interop.Outlook.Application.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox)as Microsoft.Office.Interop.Outlook.Folder;
            // Create filter
            string filter = "@SQL=" + "\""
                + PR_HAS_ATTACH + "\"" + " = 1";
            // Must use 'like' comparison for Find/FindNext
            string filter1 = "@SQL="
                + "\"" + "urn:schemas:httpmail:subject" + "\""
                + " like '%License Keys - Order%'";

            Microsoft.Office.Interop.Outlook.Table table =
                folder.GetTable(filter,
                Microsoft.Office.Interop.Outlook.OlTableContents.olUserItems);
            // Remove default columns
            table.Columns.RemoveAll();
            // Add using built-in name
            table.Columns.Add("EntryID");
            table.Columns.Add("Subject");
            table.Columns.Add("ReceivedTime");
            table.Sort("ReceivedTime", Microsoft.Office.Interop.Outlook.OlSortOrder.olDescending);
            // Add using namespace
            // Date received
            table.Columns.Add(
                "urn:schemas:httpmail:datereceived");
            while (!table.EndOfTable)
            {
                Microsoft.Office.Interop.Outlook.Row nextRow = table.GetNextRow();
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(nextRow["Subject"].ToString());
                // Reference column by name 
                sb.AppendLine("Received (Local): "
                    + nextRow["ReceivedTime"]);
                // Reference column by index
                sb.AppendLine("Received (UTC): " + nextRow[4]);
                sb.AppendLine();
                System.Diagnostics.Debug.WriteLine(sb.ToString());
            }
        }

        private string GetAttachmentPath(string AttachfileName, string AttachPath)
        {
            string filename = Path.GetFileNameWithoutExtension(AttachfileName);
            string ext = Path.GetExtension(AttachfileName);
            int app = 0;
            if (File.Exists(AttachPath + filename + ext))
            {
                while (File.Exists(AttachPath + filename + app + ext))
                {
                    app++;
                }
                filename = filename + app;
            }
            return AttachPath + filename + ext;
        }
        #endregion
    }
}
