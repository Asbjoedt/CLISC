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
        public bool Convert_OpenDocument(string argument0, string org_filepath, string file_folder)
        {
            // Use LibreOffice command line for conversion
            // --> LibreOffice has bug, so direct filepath to new converted spreadsheet cannot be specified. Only the folder can be specified
            Process app = new Process();
            app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
            app.StartInfo.Arguments = "--headless --convert-to xlsx " + org_filepath + " --outdir " + file_folder;
            app.Start();
            app.WaitForExit();
            app.Close();

            // Because of previous bug, rename converted spreadsheet according to archiving requirements
            if (argument0 == "Count&Convert&Compare&Archive")
            {
                string[] filename = Directory.GetFiles(file_folder, "*.xlsx");
                string old_filename = filename[0];
                string new_filename = file_folder + "\\1.xlsx";
                File.Move(old_filename, new_filename);
            }

            // Ordinary use
            else
            {
                // Do nothing
            }


            bool convert_success = true;

            // Inform user
            Console.WriteLine(org_filepath);
            Console.WriteLine($"--> Conversion {convert_success}");

            return convert_success;
        }

    }

}
