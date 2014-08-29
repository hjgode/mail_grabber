using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Text.RegularExpressions;

namespace Helpers
{
    /// <summary>
    /// this class processes LicenseMails
    /// the license data must be processed i the StateChanged event handler in a subscriber
    /// 
    /// </summary>
    public class LicenseMail:IDisposable
    {
//        static LicenseDataBase _licenseDataBase;

        public string ReceivedBy;
        public DateTime SendAt;
        public string OrderNumber;
        string OrderDate;
        string PONumber;
        string EndCustomer;
        string Product;
        int Quantity;

        public LicenseMail()
        {
        }

        public LicenseMail(string rb, DateTime sa, string on, string od, string po, string ec, string pr, int qu)
        {
            //email general
            ReceivedBy = rb;
            SendAt = sa;
            //email body
            OrderNumber = on;
            OrderDate = od;
            PONumber = po;
            EndCustomer = ec;
            Product = pr;
            Quantity = qu;

        }
        
        public LicenseMail(IMailMessage msg)
        {
            ReceivedBy = msg.User;
            SendAt = msg.timestamp;
        }

        public void Dispose()
        {
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
        public class LicenseMailBodyData
        {
            public string OrderNumber;
            public DateTime OrderDate;
            public string yourPOnumber;
            public string EndCustomer;
            public string Product;
            public int Quantity;

            public LicenseMailBodyData()
            {
            }
            public string dump()
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("OrderNumber: " + this.OrderNumber +"\r\n");
                sb.Append("OrderDate: " + this.OrderDate.ToShortDateString() + "\r\n");
                sb.Append("Your PO Number: " + this.yourPOnumber + "\r\n");
                sb.Append("EndCustomer: " + this.EndCustomer + "\r\n");
                sb.Append("Product: " + this.Product + "\r\n");
                sb.Append("Quantity: " + this.Quantity.ToString() + "\r\n");
                return sb.ToString();
            }
            public static LicenseMailBodyData get(IMailMessage msg)
            {
                LicenseMailBodyData data = new LicenseMailBodyData();
                string sBody = msg.Body;
                sBody = sBody.Replace("\r\n", ";");
                var expression = new Regex(
                    @"Order Number:[ ]+(?<order_number>[\S]+);" + 
                    @".+Order Date:[ ]+(?<order_date>[\S]+);" +
                    @".+Your PO Number:[ ]+(?<po_number>[\S]+);"+
                    @".+End Customer:[ ]+(?<end_customer>[\S]+);"+
                    ""
                );

                var match = expression.Match(sBody);
                //utils.helpers.addLog(string.Format("order_number......{0}", match.Groups["order_number"]));
                data.OrderNumber = match.Groups["order_number"].Value;
                
                //utils.helpers.addLog(string.Format("order_date........{0}", match.Groups["order_date"]));
                data.OrderDate = getDateTimeFromUSdateString(match.Groups["order_date"].Value);
                
                //utils.helpers.addLog(string.Format("po_number........ {0}", match.Groups["po_number"]));
                data.yourPOnumber = match.Groups["po_number"].Value;

                //utils.helpers.addLog(string.Format("end_customer..... {0}", match.Groups["end_customer"]));
                data.EndCustomer = match.Groups["end_customer"].Value;

                expression = new Regex(@".+Product:[ ]+(?<product>.+);.+Quantity");
                match = expression.Match(sBody); 
                //utils.helpers.addLog(string.Format("product...........{0}", match.Groups["product"]));
                data.Product = match.Groups["product"].Value;

                expression = new Regex(@".+Quantity:[ ]+(?<quantity>[0-9]+);");
                match = expression.Match(sBody); 
                //utils.helpers.addLog(string.Format("quantity..........{0}", match.Groups["quantity"]));
                data.Quantity = Convert.ToInt16(match.Groups["quantity"].Value);

                return data;
            }
            static DateTime getDateTimeFromUSdateString(string s)
            {
                DateTime dt=new DateTime();
                string[] ds = s.Split('/');
                if(ds.Length==3)
                    dt=new DateTime(Convert.ToInt16(ds[2]), Convert.ToInt16(ds[0]), Convert.ToInt16(ds[1]));
                else
                    dt=new DateTime(1999, 1, 1);
                return dt;
            }
        }
        public static LicenseMailBodyData processMailBody(IMailMessage msg)
        {
            LicenseMailBodyData bodyData = LicenseMailBodyData.get(msg);
            return bodyData;
        }

        /// <summary>
        /// process an email and call processAttachement for every attachement
        /// which fires an event for every new attachement
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        public int processMail(IMailMessage m)// (Microsoft.Exchange.WebServices.Data.EmailMessage m)
        {

            int iRet = 0;
            if (m == null)
            {
                utils.helpers.addLog("processMail: null msg");
                OnStateChanged(new StatusEventArgs(StatusType.error, "NULL msg"));
                return iRet;
            }

            try
            {
                OnStateChanged(new StatusEventArgs(StatusType.none, "processing "+m.Attachements.Length.ToString()+" attachements" ));
                utils.helpers.addLog(m.User +","+ m.Subject + ", # attachements: " + m.Attachements.Length.ToString() + "\r\n");
                //get data from email
                string sReceivededBy = m.User;
                DateTime dtSendAt = m.timestamp;
                
                //get data from body
                LicenseMailBodyData bodyData = new LicenseMailBodyData();
                OnStateChanged(new StatusEventArgs(StatusType.none, "processing mail body"));
                utils.helpers.addLog("processing mail body");
                bodyData = processMailBody(m);
                utils.helpers.addLog( bodyData.dump() );

                if (m.Attachements.Length > 0)
                {
                    //process each attachement
                    foreach (Attachement a in m.Attachements)
                    {
                        try
                        {
                            utils.helpers.addLog("start processAttachement...\r\n");
                            iRet += processAttachement(a, bodyData, m);
                            OnStateChanged(new StatusEventArgs(StatusType.none, "processed " + iRet.ToString() + " licenses"));
                            utils.helpers.addLog("processAttachement done\r\n");
                        }
                        catch (Exception ex)
                        {
                            utils.helpers.addExceptionLog(ex.Message);
                            OnStateChanged(new StatusEventArgs(StatusType.error, "Exception1 in processAttachement: " + ex.Message));
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                utils.helpers.addLog("Exception: " + ex.Message);
                OnStateChanged(new StatusEventArgs(StatusType.error, "Exception2 in processAttachement: " + ex.Message));
            }
            utils.helpers.addLog("processMail did process " + iRet.ToString() + " files");
            return iRet;
        }


        public int processAttachement(Attachement att, LicenseMailBodyData data, IMailMessage mail)
        {
            int iCount=0;
            LicenseXML xmlData = LicenseXML.Deserialize(att.data);

            foreach(license ldata in xmlData.licenses){
                utils.helpers.addLog("processAttachement: new LicenseData...\r\n");
                LicenseData licenseData = new LicenseData(ldata.id, ldata.user, ldata.key, data.OrderNumber, data.OrderDate, data.yourPOnumber, data.EndCustomer, data.Product, data.Quantity, mail.User, mail.timestamp);
                //if (_licenseDataBase.addQueued(licenseData))
                utils.helpers.addLog("firing license_mail event\r\n");
                OnStateChanged(new StatusEventArgs(StatusType.license_mail, licenseData));
                iCount++;
                //if (_licenseDataBase.add(ldata.id, ldata.user, ldata.key, data.OrderNumber, data.OrderDate, data.yourPOnumber, data.EndCustomer, data.Product, data.Quantity, mail.User, mail.timestamp))
                //    iCount++;
                //utils.helpers.addLog("start _licenseDataBase.add() done\r\n");
               
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
            return iCount;
        }

        public event Helpers.StateChangedEventHandler StateChanged;
        protected virtual void OnStateChanged(StatusEventArgs args)
        {

            System.Diagnostics.Debug.WriteLine("LicensMail onStateChanged: " + args.eStatus.ToString() + ":" + args.strMessage);
            StateChangedEventHandler handler = StateChanged;
            if (handler != null)
            {
                handler(this, args);
            }
        }
    }
}
