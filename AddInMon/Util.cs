using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace AddInMon
{
    class Util
    {

        private bool debug;
      
        public bool Debug
        {
            get { return debug; }
            set { debug = value; }
        } 

        public string GetTempPath()
        {
            string path = System.Environment.GetEnvironmentVariable("TEMP");
            if (!path.EndsWith("\\")) path += "\\";
            return path;
        }

        public void LogMessageToFile(string msg)
        {
            if (debug == true)
            {
                System.IO.StreamWriter sw = System.IO.File.AppendText(
                    GetTempPath() + "AddInMonDebug.txt");
                try
                {
                    string logLine = System.String.Format(
                        "{0:G}: {1}.", System.DateTime.Now, msg);
                    sw.WriteLine(logLine);
                }
                finally
                {
                    sw.Close();
                }
            }
        }

        public string GetOfficeVersion(String sAppName)
        {

            string ver = "err";
            RegistryKey r;
            try
            {

                LogMessageToFile("GetOfficeVersion...");
                LogMessageToFile("Checking install path for " + sAppName);
                r = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default).OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + sAppName, false);
                
                if (r != null)
                {
                    ver = r.GetValue("Path","err").ToString();
                }
                if (ver.Length > 3)
                {
                    LogMessageToFile("Outlook Install Path: " + ver);
                    if (ver.Contains("12"))
                    {
                        ver = "12.0";
                    }
                    else if (ver.Contains("14"))
                    {
                        ver = "14.0";
                    }
                    else if (ver.Contains("15"))
                    {
                        ver = "15.0";
                    }
                    LogMessageToFile("Office Version is: " + ver);

                    
                }
                else
                {
                    LogMessageToFile("Install Not Found");
                    
                }

            }
            catch (Exception ex)
            {
                LogMessageToFile("GetOfficeVersion Exception:" + ex.ToString());
            }
            return ver;
        }
    }
}
