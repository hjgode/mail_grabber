using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
    public class userData
    {
        public string sDomain;
        public string sUser;
        public string sPassword;
        public userData()
        {
            sDomain = "";
            sUser = "";
            sPassword = "";
        }
        public userData(string d, string u, string p)
        {
            sDomain = d;
            sUser = u;
            sPassword = p;
        }
    }
}
