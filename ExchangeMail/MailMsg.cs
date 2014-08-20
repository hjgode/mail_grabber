using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Helpers;

using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeMail
{
    public class MailMsg:IMailMessage
    {
        public string Sender {
            get { return _Sender; }
            set { _Sender = value; }
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
        public excAttachement[] Attachements
        {
            get { return _Attachements; }
            set { _Attachements = value; }
        }

        public object blob
        {
            get { return _blob; }
            set { _blob = value; }
        }

        public string id
        {
            get { return _id; }
            set { _id = value; }
        }

        string _Sender;
        string _Body;
        string _Subject;
        excAttachement[] _Attachements;
        object _blob;
        string _id;

        MailMsg(EmailMessage msg)
        {
            this._Sender = msg.Sender.ToString();
            this._id = msg.Id.ToString();

            if (msg.HasAttachments)
            {
                List<excAttachement> oList = new List<excAttachement>();
                foreach (Attachment a in msg.Attachments)
                {
                    oList.Add(new excAttachement(null, a.Name));
                }
                this._Attachements = oList.ToArray();
            }
        }
    }
}
