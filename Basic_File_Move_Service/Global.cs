using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Basic_File_Move_Service
{
    public static class Global
    {
        //public static int TimerIntervalMS = 60000; // Service Timer Interval In Milliseconds // 1 min = 60 sec
        public static int NUMBEROFMACHINES = 1;
        public static int MAXNUMBEROFMACHINES = 20;

        public static DateTime[] FileLastWriteDate = new DateTime[MAXNUMBEROFMACHINES]; // Keeps track of what the last size of the import file was
        // Initialize Variables for Import Data Paths
        public static string[] MACHINEOUTPUT_PATHNAME = new string[MAXNUMBEROFMACHINES];
        public static string[] DESTINATIONIMPORT_PATHNAME = new string[MAXNUMBEROFMACHINES];
        public static string[] MACHINENETWORKDOMAIN = new string[MAXNUMBEROFMACHINES];
        public static string[] MACHINENETWORKUSER = new string[MAXNUMBEROFMACHINES];
        public static string[] MACHINENETWORKPASSWORD = new string[MAXNUMBEROFMACHINES];
        public static string[] DESTINATIONNETWORKDOMAIN = new string[MAXNUMBEROFMACHINES];
        public static string[] DESTINATIONNETWORKUSER = new string[MAXNUMBEROFMACHINES];
        public static string[] DESTINATIONNETWORKPASSWORD = new string[MAXNUMBEROFMACHINES];
        public static bool[] MOVEORCOPY = new bool[MAXNUMBEROFMACHINES]; //false=move, true=copy
        public static bool[] ENTIREDIRECTORY = new bool[MAXNUMBEROFMACHINES];
        public static string LOGFILE_PATHNAME = "";
        public static string[] TEMP_PATHNAME = new string[MAXNUMBEROFMACHINES];
        public static string LASTACCESSDATE_PATHNAME = "";
        public static int TIMERDURATION = 60000;
        public static bool DEBUGMODE = false;
    }
}
