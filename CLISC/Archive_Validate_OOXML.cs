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
        public string Validate_OOXML(string argument1, string argument2)
        {

            // Filepath to XML error log
            string XML_error_log = results_directory + "\\validationErrors.xml";
            string XML_errors = "";

            try
            {
                // Use OOXML Validator command line for comparison
                Process app = new Process();
                app.StartInfo.FileName = $"C:\\Users\\%USERNAME%\\Desktop\\OOXMLValidatorCLI.exe";
                app.StartInfo.Arguments = $"\"{argument1}\" --xml > {XML_error_log}";
                app.Start();
                app.WaitForExit();
                app.Close();

                // Create XML log of errors
                XML_errors = File.ReadAllText($"@\"{XML_error_log}\"");

                //Contains("IsStrict = "false"")

                ooxml_strict_conformance = true;

                if (ooxml_strict_conformance == false)
                {
                    File.Delete(XML_error_log);
                }
                else
                {
                    ooxml_strict_conformance = true;
                }



                // Identify if Strict conformance


                // Return string of errors
                return XML_errors;

            }

            // If OOXML Validator cannot be found
            catch (Win32Exception)
            {
                // Inform user
                Console.WriteLine("OOXML Validator CLI executable cannot be found. Make sure the exe is located in directory: C:\\Users\\%USERNAME%\\Desktop");

                // Return error message
                return "Validation was not performed";
            }

        }

    }

}
