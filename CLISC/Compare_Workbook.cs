using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.ComponentModel;

namespace CLISC
{
    public partial class Compare
    {
        public int Compare_Workbook(string Results_Directory, string file_folder, string org_filepath, string conv_filepath)
        {
            // Use Beyond Compare 4 command line for comparison
            Process app = new Process();
            app.StartInfo.FileName = "C:\\Program Files\\Beyond Compare 4\\BCompare.exe";
            app.StartInfo.Arguments = $"\"{org_filepath}\" \"{conv_filepath}\" /silent /qc=<crc> /ro";
            app.Start();
            app.WaitForExit();
            int return_code = app.ExitCode;
            app.Close();
            return return_code;
        }
    }
}
