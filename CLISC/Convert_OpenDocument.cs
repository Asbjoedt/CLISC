using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    
    public partial class Spreadsheet
    {
        
        // Convert spreadsheets in OpenDocument file formats
        public bool Convert_OpenDocument(string org_filepath, string copy_new_filepath, string conv_new_filepath)
        {
            
            bool success;

            try
            {
                success = true;

                // Use LibreOffice command line for conversion
                System.Diagnostics.Process app = new System.Diagnostics.Process();
                app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
                app.StartInfo.Arguments = "--headless --convert-to xlsx " + copy_new_filepath + " --outdir " + conv_new_filepath;
                app.Start();
                app.WaitForExit();
                app.Close();

                // Inform user
                Console.WriteLine(org_filepath);
                Console.WriteLine($"--> Conversion {success}");

                return success;
            }

            catch (System.ComponentModel.Win32Exception)
            {
                success = false;

                // Inform user
                Console.WriteLine(org_filepath);
                Console.WriteLine($"--> Conversion {success} - {error_message[3]}");

                return success;
            }

        }

    }

}
