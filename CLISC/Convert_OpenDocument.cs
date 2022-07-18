using System;
using System.IO;
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
        public bool Convert_OpenDocument(string org_filepath, string copy_new_filepath, string conv_dir_sub)
        {

            try
            {
                convert_success = true;

                // Use LibreOffice command line for conversion
                // --> LibreOffice has bug, so direct filepath to new converted spreadsheet cannot be specified. Only the folder can be specified, which is why "conv_dir_sub" is used
                Process app = new Process();
                app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
                app.StartInfo.Arguments = "--headless --convert-to xlsx " + copy_new_filepath + " --outdir " + conv_dir_sub;
                app.Start();
                app.WaitForExit();
                app.Close();

                // Rename converted spreadsheet, because of previous bug
                string new_filename = conv_dir_sub + "\\1.xlsx";
                var conv_file = from file in
                Directory.EnumerateFiles(conv_dir_sub)
                                where Path.GetFileName(file).Contains(".xlsx")
                                select file;
                foreach (var file in conv_file)
                {
                    string incorrect_filename = file.ToString();
                    File.Move(incorrect_filename, new_filename);
                }

                // Inform user
                Console.WriteLine(org_filepath);
                Console.WriteLine($"--> Conversion {convert_success}");

                return convert_success;
            }

            catch (Win32Exception)
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
