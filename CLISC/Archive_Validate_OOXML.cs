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

        public bool ooxml_strict_conformance = false;

        // Validate Open Office XML file formats
        public string Validate_OOXML(string argument1, string argument2)
        {

            // Filepath to XML error log
            string XML_error_log = results_directory + "\\validationErrors.xml";

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
                string XML_errors = File.ReadAllText($"@\"{XML_error_log}\"");

                ooxml_strict_conformance = true;

                if (ooxml_strict_conformance == false)
                {
                    File.Delete(XML_error_log);
                }
                else
                {
                    ooxml_strict_conformance = true;
                }

                // Return string of errors
                return XML_errors;

                // Identify if Strict conformance



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
