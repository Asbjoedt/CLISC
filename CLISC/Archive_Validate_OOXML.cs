﻿using System;
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

                if (validation_errors.Any()) // If errors
                {
                    if (error_count == 45)
                    {
                        valid_files++; // Add file to number of valid spreadsheets
                        Console.WriteLine($"--> File format is valid - {error_count} incorrectly reported validation errors have been suppressed"); // Inform user
                    }
                    else
                    {
                        invalid_files++; // Add file to number of invalid spreadsheets
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
                        return validation_message = "Valid";
                    }
                    else
                    {
                        return validation_message = string.Join(", ", validation_errors); // Turn list of errors into string;
                    }

                }
                else
                {
                    valid_files++; // Add file to number of valid spreadsheets
                    Console.WriteLine($"--> File format is valid"); // Inform user
                    return validation_message = "Valid";
                }
            }
        }
    }
}
