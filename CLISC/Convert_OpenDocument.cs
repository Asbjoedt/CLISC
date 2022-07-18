using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.ComponentModel;

namespace CLISC
{
    
    public partial class Spreadsheet
    {
        
        // Convert spreadsheets in OpenDocument file formats
        public bool Convert_OpenDocument(string org_filepath, string copy_new_filepath, string conv_new_filepath)
        {

            try
            {
                convert_success = true;

                // Use LibreOffice command line for conversion
                Process app = new Process();
                app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
                app.StartInfo.Arguments = "--headless --convert-to xlsx " + copy_new_filepath + " --outdir " + conv_new_filepath;
                app.Start();
                app.WaitForExit();
                app.Close();

                // Inform user
                Console.WriteLine(org_filepath);
                Console.WriteLine($"--> Conversion {convert_success}");

                return convert_success;
            }

            catch (System.ComponentModel.Win32Exception)
            {
                convert_success = false;

                // Inform user
                Console.WriteLine(org_filepath);
                Console.WriteLine($"--> Conversion {convert_success} - {convert_error_message[3]}");

                return convert_success;
            }

        }

    }

}
