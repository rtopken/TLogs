using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace TLogs
{
    class Program
    {
        static void Main(string[] args)
        {
            string strDownloads = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            strDownloads += "\\Downloads";

            DirectoryInfo dInfo = new DirectoryInfo(strDownloads);

            FileInfo[] dlFiles = dInfo.GetFiles();

        }
    }
}
