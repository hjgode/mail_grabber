using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Redemption;

using utils;

namespace outlookGrabber_red
{
    class OutlookRedemptionClass:IDisposable
    {
        Redemption.RDOSession rdo;
        public OutlookRedemptionClass(){
            rdo = RedemptionLoader.new_RDOSession();                        
        }

        /// <summary>
        /// logon using default outlook mail profile
        /// </summary>
        /// <returns>true if no error</returns>
        public bool logon()
        {
            bool bRet = false;
            try
            {
                rdo.Logon();
                bRet = true;
            }
            catch (Exception ex)
            {
                helpers.addExceptionLog(ex.Message);
            }
            return bRet;
        }


        /// <summary>
        /// logon to exchange server
        /// </summary>
        /// <param name="sEMail">email name, ie "heinz-josef.gode@"</param>
        /// <param name="sHost">host, ie "honeywell.com"</param>
        /// <returns>true if no error</returns>
        public bool logon(string sEMail, string sHost)
        {
            bool bRet = false;
            try
            {
                rdo.LogonExchangeMailbox(sEMail, sHost);
                bRet = true;
            }
            catch (Exception ex)
            {
                helpers.addExceptionLog(ex.Message);
            }
            return bRet;
        }

        public string getMails()
        {
            StringBuilder sb = new StringBuilder();
            RDOFolder folder = rdo.GetDefaultFolder(rdoDefaultFolders.olFolderInbox);

            string sSQL = "Select Subject from Folder WHERE Subject LIKE '%Your license order%'";
            //sSQL = "Select Subject from Folder where Subject<>''";
            
            RDOMail item;
            RDOItems items = folder.Items;
            item = items.Find(sSQL);
            if (item == null)
                sb.Append("no data found");
            while (item != null)
            {
                sb.Append("===========================\r\n");
                sb.Append(item.ReceivedByName + ":");
                sb.Append(item.Subject + ":");
                if (item.Attachments.Count > 0)
                {
                    foreach (RDOAttachment a in item.Attachments)
                    {
                        sb.Append("\t" + a.FileName + "\r\n");
                    }
                }
                sb.Append("\r\n");

                item = items.FindNext();
            }

            return sb.ToString();
            
            foreach (RDOMail m in folder.Items)
            {
                try
                {
                    sb.Append("===========================\r\n");
                    sb.Append(m.ReceivedByName + ":");
                    sb.Append(m.Subject + ":");
                    if(m.Attachments.Count>0){
                        foreach(RDOAttachment a in m.Attachments){
                            sb.Append("\t" + a.FileName + "\r\n");
                        }
                    }
                    sb.Append("\r\n");
                }
                catch (Exception ex)
                {
                    helpers.addExceptionLog(m.ToString() + " is not an email");
                }
            }
            return sb.ToString();
        }
        public void Dispose()
        {
            rdo.Logoff();            
        }
    }
}
