using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Helpers
{
    public enum StatusType
    {
        none,
        idle,
        busy,
        success,
        error,
        url_changed,
        validating,
        license_mail,
        other_mail,
        bulk_mail
    }

    public class StatusEventArgs:EventArgs
    {
        public StatusType eStatus;
        public string strMessage;

        public LicenseData licenseData;

        //public IMailMessage mailmsg { 
        //    get { return myMailMsg; }
        //    private set { myMailMsg = value; }
        //}
        //IMailMessage myMailMsg;
        
        public StatusEventArgs(StatusType state)
        {
            eStatus = state;
        }
        
        public StatusEventArgs(StatusType state, string msg)
        {
            eStatus = state;
            strMessage = msg;
        }
        
        public StatusEventArgs(StatusType state, LicenseData _licenseData)
        {
            eStatus = state;
            strMessage = _licenseData._deviceid;
            licenseData = _licenseData;
        }
    }
}
