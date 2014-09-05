using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Helpers;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMail
{
    class MailMsg:IMailMessage
    {
        string _User;
        string _Body;
        string _Subject;
        Attachement[] _Attachements;

        List<Attachement> attList = new List<Attachement>();
        string _id;
        public string User
        {
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

        public DateTime timestamp
        {
            get { return _timestamp; }
            set { _timestamp = value; }
        }
        DateTime _timestamp;

        public MailMsg(Outlook.MailItem msg, string user)
        {
            this._User = user;
            this._id = msg.EntryID;
            timestamp = msg.SentOn;
            this._Subject = msg.Subject;
            this._Body = msg.Body;

            if (msg.Attachments.Count>0)
            {
                //we will add all attachements to a list
                List<Attachement> oList = new List<Attachement>();
                foreach (Outlook.Attachment att in msg.Attachments)
                {
                    //need to save file temporary to get contents
                    string filename = System.IO.Path.GetTempFileName();
                    att.SaveAsFile(filename);
                    System.IO.FileStream fs=new System.IO.FileStream(filename,System.IO.FileMode.Open);
                    oList.Add(new Attachement(fs, att.FileName));
                    fs.Close();
                    System.IO.File.Delete(filename);
                }
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
