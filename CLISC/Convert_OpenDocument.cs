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
        public bool Convert_OpenDocument(string function, string org_filepath, string input_filepath, string file_folder)
        {
            // Use LibreOffice command line for conversion
            // --> LibreOffice has bug, so direct filepath to new converted spreadsheet cannot be specified. Only the folder can be specified
            Process app = new Process();
            app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
            app.StartInfo.Arguments = "--headless --convert-to xlsx " + input_filepath + " --outdir " + file_folder;
            app.Start();
            app.WaitForExit();
            app.Close();

            bool convert_success = true;

            // Because of previous bug, rename converted spreadsheet according to archiving requirements
            if (function == "count&convert&compare&archive")
            {
                string[] filename = Directory.GetFiles(file_folder, "*.xlsx");

                if (filename.Length > 0)
                {
                    string old_filename = filename[0];
                    string new_filename = file_folder + "\\1.xlsx";
                    File.Move(old_filename, new_filename);

                    // Mark conversion as succesful
                    convert_success = true;
                    // Inform user
                    Console.WriteLine(org_filepath);
                    Console.WriteLine($"--> Conversion {convert_success}");
                    Console.WriteLine($"--> Conversion saved to: {new_filename}");

                    return convert_success;
                }
                else
                {
                    // Mark conversion as unsuccesful, if converted spreadsheet was not found
                    convert_success = false;
                    // Inform user
                    Console.WriteLine(input_filepath);
                    Console.WriteLine($"--> Conversion {convert_success}");

                    return convert_success;
                }
            }
            // Bug does not exist in ordinary use
            else
            {
                // Mark conversion as succesful
                convert_success = true;
                // Inform user
                Console.WriteLine(org_filepath);
                Console.WriteLine($"--> Conversion {convert_success}");
                Console.WriteLine($"--> Conversion saved to: {file_folder}");

                return convert_success;
            }
        }
    }
}

