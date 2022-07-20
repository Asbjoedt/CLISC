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

        // Validate Open Office XML file formats
        public Validate_OOXML(string argument1, string argument2, string argument3)
        {
            
            // Identify CLISC subdirectory
            int results_directory_number = 1;
            string results_directory = argument2 + "\\CLISC_" + dateStamp + "_v" + results_directory_number;
            while (Directory.Exists(@results_directory))
            {
                results_directory_number++;
                results_directory = argument2 + "\\CLISC_" + dateStamp + "_v" + results_directory_number;
            }
            results_directory_number = results_directory_number - 1;
            results_directory = argument2 + "\\CLISC_" + dateStamp + "_v" + results_directory_number;

            // Filepath to XML error log
            string XML_error_log = results_directory + "\\validationErrors.xml";

            try
            {
                if (argument3 == "Recurse=Yes")
                {
                    // Use OOXML Validator command line for comparison
                    Process app = new Process();
                    app.StartInfo.FileName = $"C:\\Users\\%USERNAME%\\Desktop\\OOXMLValidatorCLI.exe";
                    app.StartInfo.Arguments = $"\"{argument1}\" --xml --recursive > {XML_error_log}";
                    app.Start();
                    app.WaitForExit();
                    app.Close();

                    // Create XML log of errors
                    string XML_errors = File.ReadAllText($"@\"{XML_error_log}\"");

                    if ()
                    { 
                        File.Delete(XML_error_log);
                    }

                    // Return string of errors
                    return XML_errors;

                    // Identify if Strict conformance



                }

                else
                {
                    Process app = new Process();
                    app.StartInfo.FileName = "C:\\Users\\%USERNAME%\\Desktop\\OOXMLValidatorCLI.exe";
                    app.StartInfo.Arguments = $"\"{argument1}\" --xml > {XML_error_log}";
                    app.Start();
                    app.WaitForExit();
                    app.Close();
                }

            }

            // If OOXML Validator cannot be found
            catch (Win32Exception)
            {
                // Inform user
                Console.WriteLine("OOXML Validator executable cannot be found. Make sure the exe is located in the directory: C:\\Users\\%USERNAME%\\Desktop");
            }

        }

    }

}
