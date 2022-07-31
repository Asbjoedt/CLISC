using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.ComponentModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace CLISC
{
    public partial class Archive
    {
        public static int valid_files = 0;
        public static int invalid_files = 0;

        // Validate Open Office XML file formats
        public string Validate_OOXML(string filepath)
        {
            string validation_message = "";

            using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                // Validate
                var validator = new OpenXmlValidator();
                var validation_errors = validator.Validate(spreadsheet).ToList();
                int error_count = validation_errors.Count;
                int error_number = 0;

                // If errors
                if (validation_errors.Any())
                {
                    // Inform user
                    Console.WriteLine($"--> File format is invalid - Spreadsheet has {error_count} validation errors");
                    foreach (var error in validation_errors)
                    {
                        error_number++;
                        Console.WriteLine($"--> Error {error_number}");
                        Console.WriteLine("----> Description: " + error.Description);
                        Console.WriteLine("----> ErrorType: " + error.ErrorType);
                        Console.WriteLine("----> Node: " + error.Node);
                        Console.WriteLine("----> Path: " + error.Path.XPath);
                        Console.WriteLine("----> Part: " + error.Part.Uri);
                        if (error.RelatedNode != null)
                        {
                            Console.WriteLine("----> Related Node: " + error.RelatedNode);
                            Console.WriteLine("----> Related Node Inner Text: " + error.RelatedNode.InnerText);
                        }
                    }
                    // Add files to number of invalid spreadsheets
                    invalid_files++;
                    // Turn list into string
                    validation_message = string.Join(",", validation_errors);
                    return validation_message;
                }
                else
                {
                    // If no errors, inform user
                    valid_files++;
                    Console.WriteLine("--> File format is valid");
                    return validation_message = "Valid";
                }
            }
        }
    }
}
