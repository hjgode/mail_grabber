using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Helpers;
using System.Diagnostics;

namespace utils
{
    public class helpers
    {
        public const string filterSubject = "License Keys - Order";
        public const string filterAttachement = "xml";

        public static void addLog(string s)
        {
            System.Diagnostics.Debug.WriteLine(s);
        }
        
        public static void addExceptionLog(string s)
        {
            // Get call stack
            StackTrace stackTrace = new StackTrace();
            System.Diagnostics.Debug.WriteLine("Exception in '" + stackTrace.GetFrame(1).GetMethod().Name + ": " + s);
        }

        static string appPath = "";
        public static string getAppPath()
        {

            if (appPath.Length == 0)
            {
                appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
                if (!appPath.EndsWith("\\"))
                    appPath += "\\";
                if (!appPath.StartsWith(@"file:\\"))
                    appPath = appPath.Substring(6);            
            }

            return appPath;
        }
    }

    public class LicenseMail{
        // "License Keys - Order: 15476: [NAU-1504] CETerm for Windows CE 6.0 / 5.0 / CE .NET"
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

    }

    public class userData
    {
        public string sDomain;
        public string sUser;
        public string sPassword;
        public bool bUseProxy = false;

        public userData()
        {
            sDomain = "";
            sUser = "";
            sPassword = "";
            bUseProxy = false;
        }
        public userData(string d, string u, string p, bool bProxy)
        {
            sDomain = d;
            sUser = u;
            sPassword = p;
            bUseProxy = bProxy;
        }
    }



}
