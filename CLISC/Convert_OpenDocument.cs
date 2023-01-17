using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace CLISC
{
    public partial class Conversion
    {
        // Convert spreadsheets from OpenDocument file formats using LibreOffice
        public bool Convert_LibreOffice(string function, string input_filepath, string output_filepath, string output_folder)
        {
            bool convert_success = false;

            // Use LibreOffice command line for conversion
            // --> LibreOffice has bug, so output filepath cannot be specified. Only the output folder can be specified
            Process app = new Process();
            string? dir = null;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // If app is run on Windows
            {
                dir = Environment.GetEnvironmentVariable("LibreOffice");
            }
            if (dir != null)
            {
                app.StartInfo.FileName = dir;
            }
            else
            {
                app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
            }
            app.StartInfo.Arguments = "--headless --convert-to xlsx " + input_filepath + " --outdir " + output_folder;
            app.Start();
            app.WaitForExit();
            app.Close();

            // Because of previous bug, we must rename converted spreadsheet to meet archiving requirements, if selected
            if (function == "CountConvertCompareArchive")
            {
                string[] filename = Directory.GetFiles(output_folder, "*.xlsx");
                if (filename.Length > 0)
                {
                    // Rename converted spreadsheet
                    string old_filename = filename[0];
                    string new_filename = output_folder + "\\1.xlsx";
                    File.Move(old_filename, new_filename);
                }
            }
            // Report success if file exists - BUG: password protected ODS are returned as true, if not for below check
            if (File.Exists(output_filepath))
            {
                convert_success = true;
            }
            return convert_success;
        }

        // Convert spreadsheets to OpenDocument file formats using LibreOffice
        public bool Convert_to_ODS(string input_filepath, string output_folder)
        {
            bool convert_success = false;

            // Use LibreOffice command line for conversion
            // --> LibreOffice has bug, so output filepath cannot be specified. Only the output folder can be specified
            Process app = new Process();
            string? dir = null;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // If app is run on Windows
            {
                dir = Environment.GetEnvironmentVariable("LibreOffice");
            }
            if (dir != null)
            {
                app.StartInfo.FileName = dir;
            }
            else
            {
                app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
            }
            app.StartInfo.Arguments = "--headless --convert-to ods " + input_filepath + " --outdir " + output_folder;
            app.Start();
            app.WaitForExit();
            app.Close();

            // Because of previous bug, we must rename converted spreadsheet to meet archiving requirements, if selected
            string[] filename = Directory.GetFiles(output_folder, "*.ods");
            if (filename.Length > 0)
            {
                // Rename converted spreadsheet
                string old_filename = filename[0];
                string new_filename = output_folder + "\\1.ods";
                File.Move(old_filename, new_filename);

                // Report success if file exists - BUG: password protected ODS are returned as true, if not for below check
                if (File.Exists(new_filename))
                {
                    convert_success = true;
                    return convert_success;
                }
            }
            return convert_success;
        }

        // Convert spreadsheets from OpenDocument file formats using Excel Interop - DOES NOT SUPPORT .FODS
        public bool Convert_from_OpenDocument_ExcelInterop(string input_filepath, string output_filepath)
        {
            bool convert_success = false;

            Excel.Application app = new Excel.Application(); // Create Excel object instance
            app.DisplayAlerts = false; // Don't display any Excel prompts
            Excel.Workbook wb = app.Workbooks.Open(input_filepath, Password: "'"); // Create workbook instance

            wb.SaveAs(output_filepath, 51); // Save workbook as .xlsx Transitional
            wb.Close(); // Close workbook
            app.Quit(); // Quit Excel application

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // If app is run on Windows
            {
                Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                Marshal.ReleaseComObject(app); // Delete Excel task in task manager
            }

            convert_success = true; // Mark as succesful
            return convert_success; // Report success
        }

        // Convert spreadsheets to OpenDocument file formats using Excel Interop - DOES NOT SUPPORT .FODS
        public bool Convert_to_OpenDocument_ExcelInterop(string input_filepath, string output_filepath)
        {
            bool convert_success = false;

            Excel.Application app = new Excel.Application(); // Create Excel object instance
            app.DisplayAlerts = false; // Don't display any Excel prompts
            Excel.Workbook wb = app.Workbooks.Open(input_filepath, Password: "'"); // Create workbook instance

            wb.SaveAs(output_filepath, 60); // Save workbook as .ods
            wb.Close(); // Close workbook
            app.Quit(); // Quit Excel application

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // If app is run on Windows
            {
                Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                Marshal.ReleaseComObject(app); // Delete Excel task in task manager
            }

            convert_success = true; // Mark as successful
            return convert_success; // Report success
        }
    }
}

