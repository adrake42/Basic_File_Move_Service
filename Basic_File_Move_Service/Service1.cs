///////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                   //
// Author:		Andrew C Drake                                                                       //
// Company:     Gray Matter Systems                                                                  //
// Client:      GE Transportation - Grove City                                                       //
// Create Date: 07/22/16        Version: 1.3      Last Update: 1/25/17                               //
// Description:	Service for moving files from one server to another (from a third server)            //
//              Seperate user impersonation for each server                                          //
//                                                                                                   //
// Methodology: Application to move files from one server to another server. Service timer moves     //
//              all files from one folder to the other through a temp folder on the service's server //
//                                                                                                   //
// Note:        V 1.1  Origional versions. Windows Service                                           //
//              V 1.2  Supports seperate destination locations. Option for copy or move.             //  
//              V 1.3  Supports entire directory copies and does not copy if date is less than saved //
///////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Timers;

// For User Impersonation
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Permissions;
using System.Security.Principal;
using Microsoft.Win32.SafeHandles;

namespace Basic_File_Move_Service
{
    public partial class Service1 : ServiceBase
    {
        // Initialize Variables for Data Parsing
        FileInfo[] files1; // List of files in Input directory
        FileInfo[] files2; // List of files in Temp directory

        // Timer
        private Timer MainTimer = null;

        public Service1()
        {
            InitializeComponent();
        }

        // When service is started
        protected override void OnStart(string[] args)
        {
            // Read Config File
            string tempSkip;
            System.IO.StreamReader file = new System.IO.StreamReader("C:\\Basic_File_Move_Service\\config.txt");
            tempSkip = file.ReadLine();
            Global.MACHINEOUTPUT_PATHNAME[0] = file.ReadLine(); 
            tempSkip = file.ReadLine(); 
            while (tempSkip.Length > 2)
            {
                Global.NUMBEROFMACHINES++;
                Global.MACHINEOUTPUT_PATHNAME[Global.NUMBEROFMACHINES - 1] = tempSkip;
                tempSkip = file.ReadLine();
            }
            tempSkip = file.ReadLine(); // Skip 1 lines
            
            for (int cnt = 0; cnt < Global.NUMBEROFMACHINES; cnt++) Global.MACHINENETWORKDOMAIN[cnt] = file.ReadLine();
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            for (int cnt = 0; cnt < Global.NUMBEROFMACHINES; cnt++) Global.MACHINENETWORKUSER[cnt] = file.ReadLine();
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            for (int cnt = 0; cnt < Global.NUMBEROFMACHINES; cnt++) Global.MACHINENETWORKPASSWORD[cnt] = file.ReadLine();
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            for (int cnt = 0; cnt < Global.NUMBEROFMACHINES; cnt++) Global.DESTINATIONIMPORT_PATHNAME[cnt] = file.ReadLine();
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            for (int cnt = 0; cnt < Global.NUMBEROFMACHINES; cnt++) Global.DESTINATIONNETWORKDOMAIN[cnt] = file.ReadLine();
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            for (int cnt = 0; cnt < Global.NUMBEROFMACHINES; cnt++) Global.DESTINATIONNETWORKUSER[cnt] = file.ReadLine();
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            for (int cnt = 0; cnt < Global.NUMBEROFMACHINES; cnt++) Global.DESTINATIONNETWORKPASSWORD[cnt] = file.ReadLine();
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            for (int cnt = 0; cnt < Global.NUMBEROFMACHINES; cnt++) Global.MOVEORCOPY[cnt] = file.ReadLine().Equals("true") ? true : false;
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            for (int cnt = 0; cnt < Global.NUMBEROFMACHINES; cnt++) Global.ENTIREDIRECTORY[cnt] = file.ReadLine().Equals("true") ? true : false;
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            for (int cnt = 0; cnt < Global.NUMBEROFMACHINES; cnt++) Global.TEMP_PATHNAME[cnt] = file.ReadLine();
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            Global.LOGFILE_PATHNAME = file.ReadLine();
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            Global.LASTACCESSDATE_PATHNAME = file.ReadLine();
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            Global.TIMERDURATION = int.Parse(file.ReadLine());
            tempSkip = file.ReadLine(); tempSkip = file.ReadLine(); // Skip 2 lines
            Global.DEBUGMODE = file.ReadLine().Equals("true") ? true : false;
            file.Close();

            WriteToLogFile("CONFIG FILE OPENED (DEBUG)", true);

            var readLastDateFile = new StreamReader(Global.LASTACCESSDATE_PATHNAME);
            for (int CNT = 0; CNT < Global.NUMBEROFMACHINES; CNT++)
            {
                Global.FileLastWriteDate[CNT] = DateTime.Parse(readLastDateFile.ReadLine());
                WriteToLogFile("Last File Copy Date (DEBUG) - " + Global.FileLastWriteDate[CNT].ToString(), true);
            }
            readLastDateFile.Close();

            // Start Timer
            MainTimer = new Timer();
            this.MainTimer.Interval = Global.TIMERDURATION;
            this.MainTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.MainTimer_Tick);
            MainTimer.Enabled = true;

            WriteToLogFile("Service Started Successfully", false);
            MainTimer.Start();
        }

        // When Service is stopped
        protected override void OnStop()
        {
            MainTimer.Enabled = false;
            MainTimer.Close();
            MainTimer.Dispose();
            WriteToLogFile("Service Stopped Successfully", false);
        }

        // Main Timer
        private void MainTimer_Tick(object sender, EventArgs e)
        {
            WriteToLogFile("Timer Triggered (DEBUG)", true);

            for (int cnt = 0; cnt < Global.NUMBEROFMACHINES; cnt++)
            {
                using (new Impersonation(Global.MACHINENETWORKDOMAIN[cnt], Global.MACHINENETWORKUSER[cnt], Global.MACHINENETWORKPASSWORD[cnt]))
                {
                    WriteToLogFile("USER IMPERSONATED:" + Global.MACHINENETWORKDOMAIN[cnt] + "\\" + Global.MACHINENETWORKUSER[cnt], true);
                    try
                    {
                        var folder1 = new DirectoryInfo(GetUNCPath(Global.MACHINEOUTPUT_PATHNAME[cnt]));
                        files1 = folder1.GetFiles().OrderBy(f => f.LastWriteTime).ToArray(); // Get file list sorted by last write time (decending)    
                        WriteToLogFile("Copying Machine Data to Temp folder - " + Global.MACHINEOUTPUT_PATHNAME[cnt] + " to " + Global.TEMP_PATHNAME[cnt], true);

                        if (Global.ENTIREDIRECTORY[cnt]) // Copy entire directory (cannot move)
                        {
                            if (Global.MOVEORCOPY[cnt]) //copy
                                DirectoryCopy(Global.MACHINEOUTPUT_PATHNAME[cnt], Global.TEMP_PATHNAME[cnt], cnt, false, true, true, false);
                            else //move
                                DirectoryCopy(Global.MACHINEOUTPUT_PATHNAME[cnt], Global.TEMP_PATHNAME[cnt], cnt, false, true, true, true);
                        }
                        else // single directory
                        {
                            if (Global.MOVEORCOPY[cnt]) //copy
                                DirectoryCopy(Global.MACHINEOUTPUT_PATHNAME[cnt], Global.TEMP_PATHNAME[cnt], cnt, false, false, true, false);
                            else //move
                                DirectoryCopy(Global.MACHINEOUTPUT_PATHNAME[cnt], Global.TEMP_PATHNAME[cnt], cnt, false, false, true, true);
                        }

                        // Write all last access dates to text file
                        Global.FileLastWriteDate[cnt] = DateTime.Now;                     
                        File.WriteAllText(Global.LASTACCESSDATE_PATHNAME, Global.FileLastWriteDate[0].ToString());
                        for (int i = 1; i < Global.NUMBEROFMACHINES; i++)
                        {
                            File.AppendAllText(Global.LASTACCESSDATE_PATHNAME, Environment.NewLine + Global.FileLastWriteDate[i].ToString());
                        }
                        WriteToLogFile("File Last Access Time Updated (DEBUG)", true);
                    }
                    catch
                    {
                        WriteToLogFile("Error moving from source directory ::: " + files1[0].FullName + " -> " + Global.TEMP_PATHNAME[cnt] + files1[0].Name, false);
                    }
                }


                System.Threading.Thread.Sleep(1000); // Wait for 1 second

                using (new Impersonation(Global.DESTINATIONNETWORKDOMAIN[cnt], Global.DESTINATIONNETWORKUSER[cnt], Global.DESTINATIONNETWORKPASSWORD[cnt]))
                {
                    // Move file to Destination folder
                    try
                    {
                        var folder2 = new DirectoryInfo(Global.TEMP_PATHNAME[cnt]);
                        files2 = folder2.GetFiles().OrderBy(f => f.LastWriteTime).ToArray(); // Get file list sorted by last write time (decending) 
                        WriteToLogFile("Moving Temp Data to DESTINATION folder - " + Global.DESTINATIONIMPORT_PATHNAME[cnt], true);
                        DirectoryCopy(Global.TEMP_PATHNAME[cnt], Global.DESTINATIONIMPORT_PATHNAME[cnt], cnt, true, true, true, true);
                        WriteToLogFile("SUCCESS!", true);
                    }
                    catch
                    {
                        WriteToLogFile("Error moving from source directory ::: " + files2[0].FullName + " -> " + Global.DESTINATIONIMPORT_PATHNAME[cnt] + files2[0].Name, false);
                    }
                }
            }
        }

        // User Impseronation
        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
        public class Impersonation : IDisposable
        {
            private readonly SafeTokenHandle _handle;
            private readonly WindowsImpersonationContext _context;

            const int LOGON32_LOGON_NEW_CREDENTIALS = 9;

            public Impersonation(string domain, string username, string password)
            {
                var ok = LogonUser(username, domain, password,
                               LOGON32_LOGON_NEW_CREDENTIALS, 0, out this._handle);
                if (!ok)
                {
                    var errorCode = Marshal.GetLastWin32Error();
                    throw new ApplicationException(string.Format("Could not impersonate the elevated user.  LogonUser returned error code {0}.", errorCode));
                }

                this._context = WindowsIdentity.Impersonate(this._handle.DangerousGetHandle());
            }

            public void Dispose()
            {
                this._context.Dispose();
                this._handle.Dispose();
            }

            [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
            private static extern bool LogonUser(String lpszUsername, String lpszDomain, String lpszPassword, int dwLogonType, int dwLogonProvider, out SafeTokenHandle phToken);

            public sealed class SafeTokenHandle : SafeHandleZeroOrMinusOneIsInvalid
            {
                private SafeTokenHandle()
                    : base(true) { }

                [DllImport("kernel32.dll")]
                [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
                [SuppressUnmanagedCodeSecurity]
                [return: MarshalAs(UnmanagedType.Bool)]
                private static extern bool CloseHandle(IntPtr handle);

                protected override bool ReleaseHandle()
                {
                    return CloseHandle(handle);
                }
            }
        }

        // Log File
        public static void WriteToLogFile(string newLine, bool DebugOnlyLine)
        {
            // If debug mode enabled and is a debug line, or is not a debug line
            if ((Global.DEBUGMODE && DebugOnlyLine) || !DebugOnlyLine)
            {
                using (StreamWriter sw = File.AppendText(Global.LOGFILE_PATHNAME))
                {
                    sw.WriteLine("\n" + DateTime.Now + " - " + newLine);
                    sw.Close();
                }
            }
        }

        // Copy Entire directory, found on Microsoft website
        public static void DirectoryCopy(string sourceDirName, string destDirName, int machNum, bool forceDateOverride, bool copySubDirs, bool overWrite, bool isMove)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            WriteToLogFile("Directory copy initiated: " + sourceDirName + " to " + destDirName, true);

            if (!dir.Exists)
            {
                WriteToLogFile("Source directory does not exist or could not be found: " + sourceDirName, true);
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            // If the destination directory doesn't exist, create it.
            if (!Directory.Exists(destDirName))
            {
                WriteToLogFile("Destination directory DNE, creating dir" + destDirName, true);
                Directory.CreateDirectory(destDirName);
                WriteToLogFile("Directory Created", true);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                if (Global.FileLastWriteDate[machNum] < file.LastWriteTime || forceDateOverride) //do not move file if last write time older than last copy time. This is to save network traffic
                {
                    string temppath = Path.Combine(destDirName, file.Name);
                    if (isMove)
                    {
                        file.CopyTo(temppath, overWrite);
                        file.Delete();
                        WriteToLogFile("Move Success", true);
                    }
                    else
                    {
                        file.CopyTo(temppath, overWrite);
                        WriteToLogFile("Copy Success", true);
                    }
                }
            }

            // If copying subdirectories, copy them and their contents to new location.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {             
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    WriteToLogFile("Directory recursion initiated for " + subdir.FullName + " to " + temppath, true);
                    DirectoryCopy(subdir.FullName, temppath, machNum, forceDateOverride, copySubDirs, overWrite, isMove);
                    if (isMove)
                        subdir.Delete(true);
                }
            }
        }

        // Determines Mapped network drive's UNC path
        [DllImport("mpr.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int WNetGetConnection(
            [MarshalAs(UnmanagedType.LPTStr)] string localName,
            [MarshalAs(UnmanagedType.LPTStr)] StringBuilder remoteName,
            ref int length);

        // Determines Mapped network drive's UNC path
        public static string GetUNCPath(string originalPath)
        {
            StringBuilder sb = new StringBuilder(512);
            int size = sb.Capacity;

            // look for the {LETTER}: combination ...
            if (originalPath.Length > 2 && originalPath[1] == ':')
            {
                // don't use char.IsLetter here - as that can be misleading
                // the only valid drive letters are a-z && A-Z.
                char c = originalPath[0];
                if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
                {
                    int error = WNetGetConnection(originalPath.Substring(0, 2),
                        sb, ref size);
                    if (error == 0)
                    {
                        DirectoryInfo dir = new DirectoryInfo(originalPath);

                        string path = Path.GetFullPath(originalPath)
                            .Substring(Path.GetPathRoot(originalPath).Length);
                        return Path.Combine(sb.ToString().TrimEnd(), path);
                    }
                }
            }

            return originalPath;
        }
    }
}
