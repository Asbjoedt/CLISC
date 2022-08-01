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
    public partial class Conversion
    {
        // Convert spreadsheets from OpenDocument file formats
        public bool Convert_from_OpenDocument(string function, string input_filepath, string file_folder)
        {
            // Use LibreOffice command line for conversion
            // --> LibreOffice has bug, so direct filepath to new converted spreadsheet cannot be specified. Only the folder can be specified
            Process app = new Process();
            app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
            app.StartInfo.Arguments = "--headless --convert-to xlsx " + input_filepath + " --outdir " + file_folder;
            app.Start();
            app.WaitForExit();
            app.Close();

            bool convert_success;
            
            // Because of previous bug, we must rename converted spreadsheet to meet archiving requirements
            if (function == "count&convert&compare&archive")
            {
                string[] filename = Directory.GetFiles(file_folder, "*.xlsx");
                if (filename.Length > 0)
                {
                    // Rename converted spreadsheet
                    string old_filename = filename[0];
                    string new_filename = file_folder + "\\1.xlsx";
                    File.Move(old_filename, new_filename);
                    // Transform datatypes
                    xlsx_conv_extension = ".xlsx";
                    xlsx_conv_filename = "1" + xlsx_conv_extension;
                    xlsx_conv_filepath = file_folder + "\\" + xlsx_conv_filename;

                    // Report success if file exists - BUG: password protected ODS are returned as true, if not for below check
                    if (File.Exists(xlsx_conv_filepath))
                    {
                        convert_success = true;
                        return convert_success;
                    }
                    else
                    {
                        convert_success = false;
                        return convert_success;
                    }
                }
            }

            convert_success = false;
            return convert_success;
        }

        // Convert spreadsheets to OpenDocument file formats
        public bool Convert_to_OpenDocument(string function, string input_filepath, string file_folder)
        {
            // Use LibreOffice command line for conversion
            // --> LibreOffice has bug, so direct filepath to new converted spreadsheet cannot be specified. Only the folder can be specified
            Process app = new Process();
            app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
            app.StartInfo.Arguments = "--headless --convert-to ods " + input_filepath + " --outdir " + file_folder;
            app.Start();
            app.WaitForExit();
            app.Close();

            bool convert_success;

            // Because of previous bug, we must rename converted spreadsheet to meet archiving requirements
            if (function == "count&convert&compare&archive")
            {
                string[] filename = Directory.GetFiles(file_folder, "*.ods");
                if (filename.Length > 0)
                {
                    // Rename converted spreadsheet
                    string old_filename = filename[0];
                    string new_filename = file_folder + "\\1.ods";
                    File.Move(old_filename, new_filename);

                    // Report success if file exists - BUG: password protected ODS are returned as true, if not for below check
                    if (File.Exists(new_filename))
                    {
                        convert_success = true;
                        return convert_success;
                    }
                    else
                    {
                        convert_success = false;
                        return convert_success;
                    }
                }

            }

            convert_success = false;
            return convert_success;
        }
    }
}

