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
        public bool Convert_OpenDocument(string argument0, string org_filepath, string docCollection_subdir)
        {
            // Use LibreOffice command line for conversion
            // --> LibreOffice has bug, so direct filepath to new converted spreadsheet cannot be specified. Only the folder can be specified
            Process app = new Process();
            app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
            app.StartInfo.Arguments = "--headless --convert-to xlsx " + org_filepath + " --outdir " + docCollection_subdir;
            app.Start();
            app.WaitForExit();
            app.Close();

            // Rename converted spreadsheet according to archiving requirements, because of previous bug
            if (argument0 == "Count&Convert&Compare&Archive")
            {
                string new_filename = docCollection_subdir + "\\1.xlsx";
                var conv_file = from file in
                Directory.EnumerateFiles(docCollection_subdir)
                                where Path.GetFileName(file).Contains(".xlsx")
                                select file;
                foreach (var file in conv_file)
                {
                    string incorrect_filename = file.ToString();
                    File.Move(incorrect_filename, new_filename);
                }
            }

            // Ordinary use
            else
            {

            }

            bool convert_success = true;

            // Inform user
            Console.WriteLine(org_filepath);
            Console.WriteLine($"--> Conversion {convert_success}");

            return convert_success;
        }

    }

}
