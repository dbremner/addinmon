using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Security.Principal;
using System.IO;
using Microsoft.Win32;
using System.Management;
using System.Text.RegularExpressions;
using System.Diagnostics;


namespace AddInMon
{
    public partial class Form1 : Form
    {
        private ManagementEventWatcher watcher = null;  //Used to watch registry keys
        private Util u = new Util();                    //Class with useful utilities
        private bool allowshowdisplay = false;          //Determine whether to show the form on the screen
        private bool firstLoad = false;                 //Keep track of first time the form is loaded
        private Regex regex = null;                     //Used for Regular Expression Matching

        public Form1()
        {
            InitializeComponent();
        }

        //Initialize the program on 1st Load
        void LoadProgram()
        {
            u.Debug = false;
            RegistryKey r = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default).OpenSubKey("Software\\AddInMon", false);
            if (r != null)
            {
                u.Debug = Convert.ToBoolean(r.GetValue("Debug", false));
            }
            debugModeCheck();
            u.LogMessageToFile("Started");

            lblVer.Text = "v." + this.GetType().Assembly.GetName().Version.ToString();

            WMIReceiveEvent();
            ListBox.CheckForIllegalCrossThreadCalls = false;
            chkAutoScroll.Checked = true;


        }

        //Hide the form on initial startup
        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(allowshowdisplay ? value : allowshowdisplay);
            if (firstLoad == false)
            {
                LoadProgram();
                firstLoad = true;
            }
        } 

        //Check if Debug Log should be turned on
        void debugModeCheck()
        {
            if (u.Debug)
            {
                btnDebug.Visible = true;
                tsenableDebugLogToolStripMenuItem.Text = "Disable &Debug Log";
            }
            else
            {
                tsenableDebugLogToolStripMenuItem.Text = "Enable &Debug Log";
                btnDebug.Visible = false;
            }
        }

        
        //If the form is minimized then hide it and rely on the tray icon
        private void Form1_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                this.ShowInTaskbar = false;
            }

        }

        //Open the form if the tray icon is double clicked
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            LoadFromTray();
        }


        //User has asked for the Form to become visible via the Tray Icon
        //Show the form on the screen
        void LoadFromTray()
        {
            this.allowshowdisplay = true;
            this.Visible = !this.Visible;
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            GetRegEx();
        }


        
        //Look in the registry to determine what the regex pattern will be when
        //monitoring registry keys for addin disables
        //1st Check to see if regex pattern is set via GPO reg key
        //if not look in local registry for regex pattern
        //if it doesn't exist then this is the 1st time the program has run
        //and the user has not configured a pattern to search for
        void GetRegEx()
        {
            RegistryKey r;
            try
            {
                btnSave.Visible = true;
                u.LogMessageToFile("Load RegEx string from Registry");
                u.LogMessageToFile("Check GPO RegKey...Software\\Policies\\AddInMon");
                r = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default).OpenSubKey("Software\\Policies\\AddInMon",false);
                String sRegEx = "";
                if (r != null)
                {
                   sRegEx= r.GetValue("RegEx", "").ToString();
                }
                if (sRegEx == "")
                {
                    txtRegEx.Enabled = true;
                    lblGPO.Visible = false;

                    u.LogMessageToFile("GPO settings not found, try local system Software\\AddInMon...");
                    r = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default).OpenSubKey("Software\\AddInMon", false);
                    sRegEx = r.GetValue("RegEx", "").ToString().ToLower();
                    if (sRegEx == "")
                    {
                        u.LogMessageToFile("Local RegEx not found in Registry, will not auto delete reg values");
                    }
                    else
                    {
                        u.LogMessageToFile("Local RegEx settings found: " + sRegEx);
                        regex = new Regex(sRegEx);
                        txtRegEx.Text = sRegEx;
                    
                    }
                }
                else
                {
                    sRegEx = sRegEx.ToLower();
                    u.LogMessageToFile("GPO settings found: " + sRegEx);
                    regex = new Regex(sRegEx);
                    txtRegEx.Text = sRegEx;
                    txtRegEx.Enabled = false;
                    lblGPO.Visible = true;
                    btnSave.Visible = false;
                }
                
            }
            catch (Exception ex)
            {
                u.LogMessageToFile("RegEx Load Exception:" + ex.ToString());
            }

            btnSave.Enabled = false;
        }
      

        //Monitor for Registry Events 
        //This function will Monitor the Ressiliency key of each Office product
        //When the Office App writes to this regkey this program will delete the value
        //from the resillency key if it matches the regex pattern
        //This will stop the Office App from asking to disable the add-in on next launch
        //if the office app happens to crash while the addin is active.
        void WMIReceiveEvent()
        {
            try
            {

                //Determine version of Office. If it returns "err" then Office is not on this system
                //Try to find Outlook version, if not installed attempt to determine based on Word version
                String ver = u.GetOfficeVersion("outlook.exe");
                if (ver == "err")
                {
                    ver = u.GetOfficeVersion("winword.exe");
                }


                if (ver != "err")
                {

                    //This will create the Resillency Key if it doesn't already exist
                    CheckOfficeApps(ver, "Word");
                    CheckOfficeApps(ver, "Excel");
                    CheckOfficeApps(ver, "Powerpoint");
                    CheckOfficeApps(ver, "Outlook");
                    
                    //Get Pattern we are looking for to block particular add-ins from being disabled
                    GetRegEx();

                    this.Text = "Office " + (ver == "15.0" ? "2013" : (ver == "14.0" ? "2010" : "2007")) + " Add-In Monitor";

                  
                    
                    //Construct the query. Keypath specifies the key in the registry to watch.
                    //Note the KeyPath  must have backslashes escaped. Otherwise you 
                    //will get ManagementException.
                    WindowsIdentity currentUser = WindowsIdentity.GetCurrent();
//                    WqlEventQuery query = new WqlEventQuery("SELECT * FROM RegistryTreeChangeEvent WHERE " +
//                                    "Hive = 'HKEY_USERS' " +
//                                     @"AND (RootPath = '" + currentUser.User.Value + @"\\Software\\Microsoft\\Office\\" + ver + @"\\Word\\Resiliency' or 
//                                        RootPath = '" + currentUser.User.Value + @"\\Software\\Microsoft\\Office\\" + ver + @"\\Outlook\\Resiliency' or 
//                                        RootPath = '" + currentUser.User.Value + @"\\Software\\Microsoft\\Office\\" + ver + @"\\PowerPoint\\Resiliency' or 
//                                        RootPath = '" + currentUser.User.Value + @"\\Software\\Microsoft\\Office\\" + ver + @"\\Excel\\Resiliency')");

                    WqlEventQuery query = new WqlEventQuery("SELECT * FROM RegistryTreeChangeEvent WHERE " +
                                    "Hive = 'HKEY_USERS' " +
                                     @"AND (RootPath = '" + currentUser.User.Value + @"\\Software\\Microsoft\\Office\\" + ver + @"\\Word' or 
                                                            RootPath = '" + currentUser.User.Value + @"\\Software\\Microsoft\\Office\\" + ver + @"\\Outlook' or 
                                                            RootPath = '" + currentUser.User.Value + @"\\Software\\Microsoft\\Office\\" + ver + @"\\PowerPoint' or 
                                                            RootPath = '" + currentUser.User.Value + @"\\Software\\Microsoft\\Office\\" + ver + @"\\Excel')");

                    //Active the watcher for the specified registry keys
                    watcher = new ManagementEventWatcher(query);

                    watcher.EventArrived += new EventArrivedEventHandler(HandleEvent);

                    u.LogMessageToFile("Start listening for events...");
                    
                    //Make this work on XP
                    watcher.Scope.Path.NamespacePath = @"root\default";

                    // Start listening for events
                    watcher.Start();

                    //Update UI to allow user to stop monitoring via GUI
                    tsMonitor.Text = "Stop Monitor";
                    btnMon.Text = "Stop Monitoring";
                    return;
                }
                else
                {
                    u.LogMessageToFile("Office Install Not Found, will not monitor registry");
                    MessageBox.Show("Error unable to determine Office Version, turn on debug log and try again.","Office Version Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (ManagementException err)
            {
                u.LogMessageToFile("WMIReceiveEvent:  An error occurred while trying to receive an event: " + err.Message);
            }
        }


        //Create the resilliency key if it does not exist; otherwise WMIReceiveEvent will bomb
        private void CheckOfficeApps(string ver,string AppName)
        {

            RegistryKey r;
            r = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default).CreateSubKey("Software\\Microsoft\\Office\\" + ver + "\\" + AppName + "\\Resiliency");
            r.Flush();
            r.Close();
            

        }


        //This function is called whenever a registry key or value changes from the list
        //of registry keys we are monitoring
        //If the resillency key in the Office App contains the StartupItems folder review the 
        //list of values and see if any value matches our regex pattern.  
        //If it does delete it so the Office App doesn't ask to disable the item when starting up
        private void HandleEvent(object sender,
                EventArrivedEventArgs e)
        {
            try
            {
                RegistryKey r;
                try
                {
                    r = RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Default).OpenSubKey(e.NewEvent.Properties["RootPath"].Value.ToString() + "\\Resiliency\\StartupItems",true);
                    if (r != null)
                    {
                        foreach (string value in r.GetValueNames())
                        {
                            object objValue = r.GetValue(value); 
                            string path = null;

                            u.LogMessageToFile("RootPath:" + e.NewEvent.Properties["RootPath"].Value.ToString());
                            u.LogMessageToFile("Value:" + value);

                            try
                            {
                                if (objValue != null)
                                {

                                    byte[] bytes = (byte[])objValue;
                                    BinaryReader binaryReader = new BinaryReader(new MemoryStream(bytes));
                                    int pathLength = (int)binaryReader.BaseStream.Length;
                                    u.LogMessageToFile("PathLength:" + pathLength);
                                    binaryReader.Close();
                                    if (pathLength > 32)
                                    {
                                        path = Encoding.Unicode.GetString(bytes, 32, pathLength - 32);
                                    }
                                   
                                    u.LogMessageToFile("Data:" + path);
                                }
                                
                            }
                            catch (Exception ex)
                            {
                                u.LogMessageToFile("Data: Error Value No Longer Exists");
                                u.LogMessageToFile(ex.Message.ToString());
                                u.LogMessageToFile(ex.StackTrace);
                               
                            }


                            //If we found data still in the StartupItems folder check to see if we should delete it
                            if (path != null)
                            {
                                string appName = e.NewEvent.Properties["RootPath"].Value.ToString();
                                appName = appName.Split('\\')[5];
                                u.LogMessageToFile("AppName:" + appName);

                                lstWord.Items.Add(appName + ": " + path);

                                if (chkAutoScroll.Checked == true && lstWord.Items.Count > 2)
                                {
                                    lstWord.TopIndex = lstWord.Items.Count - 1;

                                }

                                if (regex != null)
                                {
                                    
                                    if (regex.IsMatch(path.ToLower()))
                                    {
                                        u.LogMessageToFile("Match Found: Deleted StartupItem Registry Value");
                                        r.DeleteValue(value, false);
                                    }
                                }
                            }

                        }

                    }
                }
                catch (Exception ex)
                {
                    u.LogMessageToFile("HandleEvent1 Exception:" + ex.Message + "\n" + ex.StackTrace);
                }


            }
            catch (Exception ex)
            {
                u.LogMessageToFile("HandleEvent2 Exception:" + ex.Message);
            }


         

        }

      
        //User has selected the Quit Button from the tray menu
        private void tsQuit_Click(object sender, EventArgs e)
        {
            ExitCheck();
        }

        //User has toggled the registry monitor from the tray menu
        private void tsMonitor_Click(object sender, EventArgs e)
        {
            MonitorCheck();
        }

        //Determine whether to stop or start the registry monitor
        private void MonitorCheck()
        {
            if (watcher == null)
            {
                WMIReceiveEvent();  //Start the Monitor
            }
            else
            {
                u.LogMessageToFile("Stop monitoring Registry");
                watcher.Stop();  //Stop the Monitor
                watcher = null;
                tsMonitor.Text = "Start Monitoring";
                btnMon.Text = "Start Monitoring";
            }
        }

        //If the User has selected the X to close the form ask if they really wanted to do that
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (e.CloseReason == CloseReason.UserClosing)
            {

                u.LogMessageToFile("Form Closing Reason:" + e.CloseReason);

                System.Windows.Forms.DialogResult dRes = MessageBox.Show("Are you sure you want to quit the program?  Click Yes to Quit, No to Minimize, or Cancel to go back.", "Quit Program", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

                if (dRes == System.Windows.Forms.DialogResult.Yes)
                {
                    ExitCheck();
                }
                else if (dRes == System.Windows.Forms.DialogResult.No)
                {
                    this.Hide();
                    this.ShowInTaskbar = false;
                    e.Cancel = true;
                }
                else
                {
                    e.Cancel = true;
                }

            }
        }

        //Cleanup when the program is closing
        private void ExitCheck()
        {
            if (watcher != null)
            {
                u.LogMessageToFile("Stop monitoring Registry");
                watcher.Stop();
            }
            u.LogMessageToFile("Application Quit");
            Application.Exit();
        }

        //Clears the list showing registry changes
        private void btnClear_Click(object sender, EventArgs e)
        {
            lstWord.Items.Clear();
        }


        private void chkAutoScroll_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAutoScroll.Checked == true && lstWord.Items.Count > 0)
            {
                lstWord.TopIndex = lstWord.Items.Count - 1;
            }
        }

        private void txtRegEx_TextChanged(object sender, EventArgs e)
        {
            btnSave.Enabled = true;
            btnSave.Visible = true;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                btnSave.Enabled = false;
                regex = new Regex(txtRegEx.Text);
                RegistryKey r = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default).OpenSubKey("Software", true);
                RegistryKey a = r.CreateSubKey("AddInMon");
                a.SetValue("RegEx", txtRegEx.Text);
            }
            catch (Exception ex)
            {
                
                String sErr = "Error Saving Key:" + ex.Message.ToString();
                MessageBox.Show(sErr,"Error Saving",MessageBoxButtons.OK,MessageBoxIcon.Error);
                u.LogMessageToFile(sErr);
                GetRegEx();

            }
            

        }


        //Listbox override method to perform highlighting of reg deletes
        private void lstWord_DrawItem(object sender, DrawItemEventArgs e)
        {
            try
            {
                if (e.Index >= 0)
                {

                    //
                    // Draw the background of the ListBox control for each item.
                    // Create a new Brush and initialize to a Black colored brush
                    // by default.
                    //
                    e.DrawBackground();

                    Brush myBrush = Brushes.White;

                    //
                    // Determine the color of the brush to draw each item based on 
                    // the index of the item to draw.
                    //
                    if (regex != null)
                    {
                        if (regex.IsMatch(((ListBox)sender).Items[e.Index].ToString()))
                        {
                            myBrush = Brushes.Goldenrod;
                        }
                    }
                    e.Graphics.FillRectangle(myBrush, e.Bounds);

                    //
                    // Draw the current item text based on the current 
                    // Font and the custom brush settings.
                    //
                    e.Graphics.DrawString(((ListBox)sender).Items[e.Index].ToString(),
                         e.Font, Brushes.Black, e.Bounds, StringFormat.GenericDefault);
                    //
                    // If the ListBox has focus, draw a focus rectangle 
                    // around the selected item.
                    //
                    e.DrawFocusRectangle();
                }
            }
            catch (Exception ex)
            {
                u.LogMessageToFile("Listbox Draw Exception:" + ex.Message);
            
            }

        }


        private void btnMon_Click(object sender, EventArgs e)
        {
            MonitorCheck();
        }

        private void tsEditItems_Click(object sender, EventArgs e)
        {
            LoadFromTray();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnDebug_Click(object sender, EventArgs e)
        {
            Process notePad = new Process();

            notePad.StartInfo.FileName = "notepad.exe";
            notePad.StartInfo.Arguments = u.GetTempPath() + "AddInMonDebug.txt";
            notePad.Start();

        }

        private void tsenableDebugLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (u.Debug == true)
            {
                u.LogMessageToFile("Diable Debugging. Goodbye.");
                u.Debug = false;
            }
            else
            {
                u.Debug = true;
                u.LogMessageToFile("Debugging Enabled. Hello.");
            }
            
            debugModeCheck();

        }

       

     

        
        
      




    }
}
