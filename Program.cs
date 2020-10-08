using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Net.NetworkInformation;
using System.IO.Compression;

namespace TLogs
{
    class Program
    {
        // Various locations for Teams log files...
        static string strTeams = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\microsoft\\teams";
        static string strAddin = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\microsoft\\teams\\meeting-addin";
        static string strMedia = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\microsoft\\teams\\media-stack";
        static string strDownloads = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
        static string strTLogs = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads\\TLogs";
        // Location for the executable...
        static string strApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\microsoft\\teams\\current\\teams.exe";

        static bool bGotDiagFiles = false;

        static void Main(string[] args)
        {
            Console.WriteLine("");
            Console.WriteLine("==================");
            Console.WriteLine("TLogs - Teams Logs");
            Console.WriteLine("==================");
            Console.WriteLine("Gets all Teams logs and places them in the TLogs folder in Downloads.\r\n");
            
            string[] strFolders = new string[]
            {
                strTeams,
                strAddin,
                strMedia,
                strDownloads,
            };
            bool bGetDiags = false;

            Console.WriteLine("Ensuring there is a TLogs directory under Downloads.");
            CreateTLogsDir();
            Console.WriteLine("Removing old logs if needed.");
            DeleteOldLogs();

            Console.WriteLine("Getting Teams Diagnostic Logs...");
            bGetDiags = GetTeamsDiag();
            if (bGetDiags == false)
            {
                Console.WriteLine("Can't find the correct Teams window. It might be minimized so attempting to make it active.");
                LaunchTeams(); // might be that Teams is minimized - so launch it to get the UI and try again.
                System.Threading.Thread.Sleep(1000);
                bGetDiags = GetTeamsDiag();
            }
            
            if (bGetDiags == false)
            {
                Console.WriteLine("Please make sure Teams is running and not minimized to the taskbar.");
                Console.WriteLine("Then run TLogs again to get the logs.");
                Console.WriteLine("Done!");

                return;
            }
             

            System.Threading.Thread.Sleep(1000);   // A pause to ensure the diag files are where they need to be...

            Console.WriteLine("Copying logs to the Downloads\\TLogs folder.");
            // now copy other logs we want to the Downloads folder
            foreach (string strFold in strFolders)
            {
                GetTeamsFiles(strFold); 
            }

            if (bGotDiagFiles == true)
            {
                Console.WriteLine("Zipping the files to TLogs.zip.");
                ZipDiagLogs();

                Console.WriteLine("Removing log files now that they have been zipped up.");
                DeleteTLogsFiles();
            }

            Console.WriteLine("Opening Explorer to the TLogs folder.");
            Process.Start(strTLogs);

            Console.WriteLine("\r\nDone!");
            
            return;
        }

        static void ZipDiagLogs()
        {
            string strFileName = "";
            var files = Directory.GetFiles(strTLogs);

            using (FileStream zipStream = new FileStream(strTLogs + "\\TLogs.zip", FileMode.Create))
            {
                using (ZipArchive zipFile = new ZipArchive(zipStream, ZipArchiveMode.Create))
                {
                    foreach (var file in files)
                    {
                        strFileName = Path.GetFileName(file);
                        if (strFileName.ToLower() != "desktop.ini")
                        {
                            zipFile.CreateEntryFromFile(file, strFileName);
                        }
                    }
                }
            }
        }

        // Create the TLogs folder under Downloads
        static void CreateTLogsDir()
        {
            if (Directory.Exists(strTLogs))
            {
                return;
            }
            else
            {
               Directory.CreateDirectory(strTLogs);
            }
        }

        static void GetTeamsFiles(string strPath)
        {
            /*
             * Want to get:  "MSTeams Diagnostics Log xxxxx.txt" in Downloads - they will be there from GetTeamsDiag function
             *               "logs.txt" from AppData\Roaming\microsoft\teams
             *               "teams-meeting-addin*.*" from AppData\Roaming\microsoft\teams\meeting-addin
             *               All files from AppData\Roaming\microsoft\teams\media-stack
            */
            string strFile = "";
            string strDestFile = "";
            string[] logFiles = Directory.GetFiles(strPath);

            if(strPath == strTeams)
            {
                Console.WriteLine("Getting logs.txt from the Teams folder.");
                File.Copy(strTeams + "\\logs.txt", strTLogs + "\\logs.txt", true);
            }
            else if (strPath == strDownloads)
            {
                Console.WriteLine("Getting MSTeams Diagnostic Log files from the Downloads folder.");
                foreach (string file in logFiles)
                {
                    if (file.Contains("MSTeams Diagnostics Log"))
                    {
                        strFile = Path.GetFileName(file);
                        strDestFile = Path.Combine(strTLogs, strFile);
                        File.Move(file, strDestFile);
                    }
                }
            }
            else
            {
                if (strPath == strAddin)
                    Console.WriteLine("Getting log files from the Addin folder.");
                else if (strPath == strMedia)
                    Console.WriteLine("Getting log files from the Media folder.");

                foreach (string file in logFiles)
                {
                    strFile = Path.GetFileName(file);
                    strDestFile = Path.Combine(strTLogs, strFile);
                    File.Copy(file, strDestFile, true);
                }
            }
        }

        // Delete anything in the TLogs folder if anything is there, and delete any old MSTeams Diag files from Downloads
        static void DeleteOldLogs()
        {
            string[] strFolders = new string[]
            {
                strTLogs,
                strDownloads,
            };

            foreach (string folder in strFolders)
            {
                string[] strFiles = Directory.GetFiles(folder);
                if (folder.Contains("TLogs"))
                {
                    foreach (string file in strFiles)
                    {
                        File.Delete(file);
                    }
                }
                else
                {
                    foreach (string file in strFiles)
                    {
                        if (file.Contains("MSTeams Diagnostics Log"))
                        {
                            File.Delete(file);
                        }
                    }
                }
            }
        }

        static void DeleteTLogsFiles()
        {
            string[] strFiles = Directory.GetFiles(strTLogs);
            foreach (string file in strFiles)
            {
                if (file.Contains("TLogs.zip"))
                    continue;
                else
                    File.Delete(file);
            }
        }

        // Get the current Teams Diagnostic Log files placed in the "Downloads" folder
        static bool GetTeamsDiag()
        {
            Console.WriteLine("Putting Teams windows in the foreground in order to get the diagnostic logs.");
            var hWnds = FindWindowsWithText("| Microsoft Teams");  // Will find the Teams UI windows
            
            // now iterate through them - one will get us our logs...
            foreach (var handle in hWnds)
            {
                SetForegroundWindow(handle); // Make it foreground so that the below keystrokes will work to get the files
                System.Threading.Thread.Sleep(500); // Pause half a second to ensure foreground
                SendKeys.SendWait("^%+1");   // Send - CRTL(^) + ALT(%) + SHIFT(+) + 1 

                System.Threading.Thread.Sleep(1500);  //pause for a bit to let the diag files to get into the folder...

                string[] strFiles = Directory.GetFiles(strDownloads);
                foreach (string file in strFiles)
                {
                    if (file.Contains("MSTeams Diagnostics Log"))
                    {
                        Console.WriteLine("Successfully generated MSTeams Diagnostics Logs.");
                        bGotDiagFiles = true;
                        break;
                    }
                }

                if (bGotDiagFiles == true)
                    break;
            }

            if (bGotDiagFiles == true)
                return true;
            else
                return false;
        }

        /*
         * Credit for finding the Teams window:
         * https://stackoverflow.com/questions/19867402/how-can-i-use-enumwindows-to-find-windows-with-a-specific-caption-title
        */

        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("USER32.DLL")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        // Delegate to filter which windows to include 
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        public static string GetWindowText(IntPtr hWnd)
        {
            int size = GetWindowTextLength(hWnd);
            if (size > 0)
            {
                var builder = new StringBuilder(size + 1);
                GetWindowText(hWnd, builder, builder.Capacity);
                return builder.ToString();
            }

            return String.Empty;
        }

        public static IEnumerable<IntPtr> FindWindows(EnumWindowsProc filter)
        {
            IntPtr found = IntPtr.Zero;
            List<IntPtr> windows = new List<IntPtr>();

            EnumWindows(delegate (IntPtr wnd, IntPtr param)
            {
                if (filter(wnd, param))
                {
                    // only add the windows that pass the filter
                    windows.Add(wnd);
                }

                // but return true here so that we iterate all windows
                return true;
            }, IntPtr.Zero);

            return windows;
        }

        public static IEnumerable<IntPtr> FindWindowsWithText(string titleText)
        {
            return FindWindows(delegate (IntPtr wnd, IntPtr param)
            {
                return GetWindowText(wnd).Contains(titleText);
            });
        }

        static IntPtr LaunchTeams()
        {
            IntPtr hTeams;

            Process pTeams = new Process();
            pTeams.StartInfo.FileName = strApp;
            pTeams.Start();
            hTeams = pTeams.Handle;

            return hTeams;
        }
    }
}
