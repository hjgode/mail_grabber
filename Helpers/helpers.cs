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
        public static bool bDoLogging = true;
        static uint maxFileSize = 1024*1024;
        public static void addLog(string s)
        {
            System.Diagnostics.Debug.WriteLine(s);
            if (bDoLogging)
            {
                try
                {
                    if (System.IO.File.Exists(LogFilename))
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(LogFilename);
                        if (fi.Length > maxFileSize)
                        {
                            //create backup file
                            System.IO.File.Copy(LogFilename, LogFilename + "bak", true);
                            System.IO.File.Delete(LogFilename);
                        }
                    }
                    System.IO.StreamWriter sw = new System.IO.StreamWriter(LogFilename, true);
                    sw.WriteLine(DateTime.Now.ToShortDateString()+ " " + DateTime.Now.ToShortTimeString() + "\t" + s);
                    sw.Flush();
                    sw.Close();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Exception in addLog FileWrite: " + ex.Message);
                }
            }
        }

        static string sLogFilename = "";
        static string LogFilename
        {
            get
            {
                if (sLogFilename.Length == 0)
                {
                    sLogFilename = getAppPath() + "ews_grabber.log.txt";
                }
                return sLogFilename;
            }
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

    public class UserData
    {
        public enum UserLogonType
        {
            none,
            exchange,
            outlook,
        }

        public string sDomain;
        public string sUser;
        public string sPassword;
        public bool bUseProxy = false;
        public string WebProxy;
        public int WebProxyPort;

        UserLogonType userLogonType = UserLogonType.exchange;
        public string sOutlookProfile = "";
        public bool bShowOutlookDialog = false;

        /// <summary>
        /// generic
        /// </summary>
        public UserData()
        {
            sDomain = "";
            sUser = "";
            sPassword = "";
            bUseProxy = false;
            WebProxy = "";
            WebProxyPort = 8080;
            userLogonType = UserLogonType.exchange;
        }

        /// <summary>
        /// outlook user logon data
        /// </summary>
        /// <param name="u">user name</param>
        /// <param name="p">password</param>
        /// <param name="b">show dialog or not</param>
        public UserData(string u, string p, bool b)
        {
            userLogonType = UserLogonType.outlook;
            sUser = u;
            sPassword = p;
            bShowOutlookDialog = b;
            userLogonType = UserLogonType.outlook;
            sOutlookProfile = "";
        }
        /// <summary>
        /// outlook user logon data
        /// </summary>
        /// <param name="bDialog">show dialog</param>
        /// <param name="profile">outlook profile name or "" for default</param>
        /// <param name="user">user name</param>
        /// <param name="pass">password</param>
        public UserData(bool bDialog, string profile, string user, string pass)
        {
            userLogonType = UserLogonType.outlook;
            bShowOutlookDialog = bDialog;
            sOutlookProfile = profile;
            sUser = user;
            sPassword = pass;
        }

        /// <summary>
        /// exchange user data
        /// </summary>
        /// <param name="d"></param>
        /// <param name="u"></param>
        /// <param name="p"></param>
        /// <param name="bProxy"></param>
        public UserData(string d, string u, string p, bool bProxy)
        {
            sDomain = d;
            sUser = u;
            sPassword = p;
            bUseProxy = bProxy;
            userLogonType = UserLogonType.exchange;
        }
        public UserData(string d, string u, string p, bool bProxy, string sWebProxy, int iWebProxyPort)
        {
            sDomain = d;
            sUser = u;
            sPassword = p;
            bUseProxy = bProxy;
            WebProxy = sWebProxy;
            WebProxyPort = iWebProxyPort;
            userLogonType = UserLogonType.exchange;
        }
    }



}
