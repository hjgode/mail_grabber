using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Helpers
{

    public class LicenseMail
    {
        public string OrderNumber;
        string OrderDate;
        string PONumber;
        string EndCustomer;
        string Product;
        int Quantity;

        public LicenseMail(string on, string od, string po, string ec, string pr, int qu)
        {
            OrderNumber = on;
            OrderDate = od;
            PONumber = po;
            EndCustomer = ec;
            Product = pr;
            Quantity = qu;
        }

        //we need to parse the body for "Order number:", "Order Date" ...
        // Subject: "License Keys - Order: 15476: [NAU-1504] CETerm for Windows CE 6.0 / 5.0 / CE .NET"
        // body=
        //Order Number:     15476
        //Order Date:       6/20/2014
        //Your PO Number:   PO96655
        //End Customer:     Honeywell
        //Product:          [NAU-1504] CETerm for Windows CE 6.0 / 5.0 / CE .NET
        //Quantity:         28

        //Qty Ordered...............: 28
        //Qty Shipped To Date.......: 28

        //Qty Shipped in this email.: 28

        public static int processMail(IMailMessage m)// (Microsoft.Exchange.WebServices.Data.EmailMessage m)
        {
            int iRet = 0;
            if (m == null)
            {
                utils.helpers.addLog("processMail: null msg");
                return iRet;
            }
            try
            {
                utils.helpers.addLog(m.User + m.Subject + "# attachements: " + m.Attachements.Length.ToString() + "\r\n");
                //get order ID out of subject text
                string sOrderNumber = "";//TODO

                if (m.Attachements.Length > 0)
                {
                    //process each attachement
                    foreach (Attachement a in m.Attachements)
                    {
                        processAttachement(a, m.User, sOrderNumber, m.timestamp);
                    }

                    #region alternative_code
                    /*
                    // Request all the attachments on the email message. This results in a GetItem operation call to EWS.
                    m.Load(new Microsoft.Exchange.WebServices.Data.PropertySet(Microsoft.Exchange.WebServices.Data.EmailMessageSchema.Attachments));
                    foreach (Microsoft.Exchange.WebServices.Data.Attachment att in m.Attachments)
                    {
                        if (att is Microsoft.Exchange.WebServices.Data.FileAttachment)
                        {
                            Microsoft.Exchange.WebServices.Data.FileAttachment fileAttachment = att as Microsoft.Exchange.WebServices.Data.FileAttachment;
                            
                            //get a temp file name
                            string fname = System.IO.Path.GetTempFileName(); //utils.helpers.getAppPath() + fileAttachment.Id.ToString() + "_" + fileAttachment.Name
                            
                            // Load the file attachment into memory. This gives you access to the attachment content, which 
                            // is a byte array that you can use to attach this file to another item. This results in a GetAttachment operation
                            // call to EWS.
                            fileAttachment.Load();

                            // Load attachment contents into a file. This results in a GetAttachment operation call to EWS.
                            fileAttachment.Load(fname);
                            addLog("Attachement file saved to: " + fname);
                            
                            // Put attachment contents into a stream.
                            using (System.IO.FileStream theStream =
                                new System.IO.FileStream(utils.helpers.getAppPath() + fileAttachment.Id.ToString() + "_" + fileAttachment.Name, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite))
                            {
                                //This results in a GetAttachment operation call to EWS.
                                fileAttachment.Load(theStream);
                            }

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
                            addLog("saved attachement: " + fname);
                            iRet++;
                        }
                    }
                    */
                    #endregion
                }
            }
            catch (Exception ex)
            {
                utils.helpers.addLog("Exception: " + ex.Message);
            }
            utils.helpers.addLog("processMail did process " + iRet.ToString() + " files");
            return iRet;
        }


        public static void processAttachement(Attachement att, string user, string order, DateTime dt)
        {
        }
    }
}
