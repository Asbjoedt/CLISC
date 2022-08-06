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
    public class Validation
    {
        public string Validity { get; set; }

        public int? Error_Number { get; set; }

        public string? Error_Description { get; set; }

        public string? Error_Type { get; set; }

        public string? Error_Node { get; set; }

        public string? Error_Path { get; set; }

        public string? Error_Part { get; set; }

        public string? Error_RelatedNode { get; set; }

        public string? Error_RelatedNode_InnerText { get; set; }

        // Validate Open Office XML file formats
        public List<Validation> Validate_OOXML(string org_filepath, string filepath, string Results_Directory)
        {
            List<Validation> results = new List<Validation>();

            using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                // Validate
                var validator = new OpenXmlValidator();
                var validation_errors = validator.Validate(spreadsheet).ToList();
                int error_count = validation_errors.Count;
                int error_number = 0;

                if (validation_errors.Any()) // If errors, inform user & return results
                {
                    if (error_count == 45)
                    {
                        Console.WriteLine($"--> File format is valid - {error_count} incorrectly reported validation errors have been suppressed"); // Inform user
                    }
                    else
                    {
                        Console.WriteLine($"--> File format is invalid - Spreadsheet has {error_count} validation errors"); // Inform users
                    }

                    foreach (var error in validation_errors)
                    {
                        // Open XML SDK has 45 bugs, that is incorrectly reported as 45 errors for Strict conformant spreadsheets. The switch suppresses these
                        switch (error.Node)
                        {
                            case DocumentFormat.OpenXml.Spreadsheet.Border:
                            case DocumentFormat.OpenXml.Drawing.FillToRectangle:
                            case DocumentFormat.OpenXml.Drawing.GradientStop:
                            case DocumentFormat.OpenXml.Drawing.SaturationModulation:
                            case DocumentFormat.OpenXml.Drawing.LuminanceModulation:
                            case DocumentFormat.OpenXml.Drawing.Shade:
                            case DocumentFormat.OpenXml.Drawing.Tint:
                            case DocumentFormat.OpenXml.Drawing.Alpha:
                            case DocumentFormat.OpenXml.Drawing.Miter:
                            case DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties:
                                // Do nothing
                                break;

                            default:
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
                                break;
                        }
                    }
                    if (error_count == 45)
                    {
                        foreach (var error in validation_errors)
                        {
                            // Add validation results to list
                            results.Add(new Validation { Validity = "Valid", Error_Number = null, Error_Description = "", Error_Type = "", Error_Node = "", Error_Path = "", Error_Part = "", Error_RelatedNode = "", Error_RelatedNode_InnerText = "" });
                        }

                        return results;
                    }
                    else
                    {
                        error_number = 0;
                        foreach (var error in validation_errors)
                        {
                            error_number++;

                            string? er_rel_1 = "";
                            string? er_rel_2 = "";
                            if (error.RelatedNode != null)
                            {
                                er_rel_1 = error.RelatedNode.ToString();
                                er_rel_2 = error.RelatedNode.InnerText;
                            }
                            // Add validation results to list
                            results.Add(new Validation { Validity = "Invalid", Error_Number = error_number, Error_Description = error.Description, Error_Type = error.ErrorType.ToString(), Error_Node = error.Node.ToString(), Error_Path = error.Path.XPath.ToString(), Error_Part = error.Part.Uri.ToString(), Error_RelatedNode = er_rel_1, Error_RelatedNode_InnerText = er_rel_2});
                        }
                        return results;
                    }
                }
                else
                {
                    // Add validation results to list
                    results.Add(new Validation { Validity = "Valid", Error_Number = null, Error_Description = "", Error_Type = "", Error_Node = "", Error_Path = "", Error_Part = "", Error_RelatedNode = "", Error_RelatedNode_InnerText = "" });

                    // Inform user & return list
                    Console.WriteLine($"--> File format is valid");
                    return results;
                }
            }
        }
    }
}
