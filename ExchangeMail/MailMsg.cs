using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Helpers;

using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeMail
{
    /// <summary>
    /// this class implements the IMailMessage interface for Exchange Microsoft.Exchange.WebServices.Data.EmailMessage
    /// </summary>
    public class MailMsg:IMailMessage
    {
        public string User {
            get { return _User; }
            set { _User = value; }
        }
        public string Body
        {
            get { return _Body; }
            set { _Body = value; }
        }
        public string Subject
        {
            get { return _Subject; }
            set { _Subject = value; }
        }
        public Attachement[] Attachements
        {
            get { return _Attachements; }
            set { _Attachements = value; }
        }

        public string id
        {
            get { return _id; }
            set { _id = value; }
        }

        string _User;
        string _Body;
        string _Subject;
        Attachement[] _Attachements;

        List<Attachement> attList = new List<Attachement>();
        string _id;

        public DateTime timestamp { 
            get { return _timestamp; } 
            set {_timestamp=value; }
        }
        DateTime _timestamp;        

        public MailMsg(EmailMessage msg, string user)
        {
            this._User = user;
            this._id = msg.Id.ToString();
            timestamp = msg.DateTimeSent;
            this._Subject = msg.Subject;
            this._Body = msg.Body;

            if (msg.HasAttachments)
            {
                //we will add all attachements to a list
                List<Attachement> oList = new List<Attachement>();
                //load all attachements
                msg.Load(new Microsoft.Exchange.WebServices.Data.PropertySet(Microsoft.Exchange.WebServices.Data.EmailMessageSchema.Attachments));
                foreach (Microsoft.Exchange.WebServices.Data.Attachment att in msg.Attachments)
                {
                    if (att is Microsoft.Exchange.WebServices.Data.FileAttachment)
                    {
                        //load the attachement
                        Microsoft.Exchange.WebServices.Data.FileAttachment fileAttachment = att as Microsoft.Exchange.WebServices.Data.FileAttachment;
                        //load into memory stream, seems the only stream supported
                        System.IO.MemoryStream ms = new System.IO.MemoryStream(att.Size);
                        fileAttachment.Load(ms);//blocks some time
                        ms.Position = 0;
                        oList.Add(new Attachement(ms, fileAttachment.Name));
                        ms.Close();
                    }
                }
                /*
                foreach (Attachment a in msg.Attachments)
                {
                    if (a is Microsoft.Exchange.WebServices.Data.FileAttachment)
                    {
                        FileAttachment fa = a as FileAttachment;
                        System.IO.MemoryStream ms=new System.IO.MemoryStream(fa.Size);
                        fa.Load(ms);
                        oList.Add(new Attachement(ms, fa.Name));
                    }
                }
                */
                this._Attachements = oList.ToArray();
                this.attList.AddRange(oList);
            }
        }
        public void addAttachement(Attachement att)
        {
            this.attList.Add(att);
        }
    }
}
