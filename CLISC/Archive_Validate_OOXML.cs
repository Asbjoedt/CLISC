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

    public partial class Spreadsheet
    {

        public bool valid_file_format = true;

        // Validate Open Office XML file formats
        public bool Validate_OOXML(string argument1)
        {

            using (var spreadsheet = SpreadsheetDocument.Open(conv_filepath, false))
            {
                var validator = new OpenXmlValidator();
                var validation_errors = validator.Validate(spreadsheet).ToList();
                var error_count = String.Format("Spreadsheet has {0} validation errors", validation_errors.Count);

                if (validation_errors.Any())
                {
                    // Open CSV file to log results
                    var csv = new StringBuilder();
                    var newLine0 = string.Format($"Convert filepath;Validation error messages");
                    csv.AppendLine(newLine0);

                    valid_file_format = false;
                    Console.WriteLine(error_count);
                    Console.WriteLine();

                    foreach (var error in validation_errors)
                    {
                        Console.WriteLine("Description: " + error.Description);
                        Console.WriteLine("ErrorType: " + error.ErrorType);
                        Console.WriteLine("Node: " + error.Node);
                        Console.WriteLine("Path: " + error.Path.XPath);
                        Console.WriteLine("Part: " + error.Part.Uri);
                        if (error.RelatedNode != null)
                        {
                            Console.WriteLine("Related Node: " + error.RelatedNode);
                            Console.WriteLine("Related Node Inner Text: " + error.RelatedNode.InnerText);
                        }
                        Console.WriteLine();
                        Console.WriteLine("==============================");
                        Console.WriteLine();

                        // Output result in open CSV file
                        var newLine1 = string.Format($"{conv_filepath};{error}");
                        csv.AppendLine(newLine1);
                    }

                    // Close CSV file to log results
                    string CSV_filepath = docCollection_subdir + "\\validationErrors.csv";
                    File.WriteAllText(CSV_filepath, csv.ToString());

                }

                return valid_file_format;

            }

        }

    }

}
