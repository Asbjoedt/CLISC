using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace CLISC
{

    public partial class Spreadsheet
    {

        public int conformance_count_fail = 0;

        // Count XLSX Transtional conformance
        public int Count_XLSX_Transitional(string argument1, string argument3)
        {

            var xlsx_enumeration = new List<string>();

            // Recurse enumeration of original spreadsheets from input directory
            if (argument3 == "Recurse=Yes")
            {
                xlsx_enumeration = (List<string>)Directory.EnumerateFiles(argument1, "*.xlsx", SearchOption.AllDirectories)
                    .ToList();

                try
                {
                    foreach (var file in xlsx_enumeration)
                    {
                        SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(file, false);
                        bool? strict = spreadsheet.StrictRelationshipFound;
                        spreadsheet.Close();

                        if (strict != true)
                        {
                            numXLSX_Transitional++;
                        }
                    }
                }

                catch (InvalidDataException)
                {
                    conformance_count_fail++;
                }

                return numXLSX_Transitional;
            }

            else
            {
                xlsx_enumeration = (List<string>)Directory.EnumerateFiles(argument1, "*.xlsx", SearchOption.TopDirectoryOnly)
                    .ToList();

                try
                {
                    foreach (var file in xlsx_enumeration)
                    {
                        SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(file, false);
                        bool? strict = spreadsheet.StrictRelationshipFound;
                        spreadsheet.Close();

                        if (strict != true)
                        {
                            numXLSX_Transitional++;
                        }
                    }
                }

                catch (InvalidDataException)
                {
                    conformance_count_fail++;
                }

                return numXLSX_Transitional;

            }

        }

        // Count XLSX Strict conformance
        public int Count_XLSX_Strict(string argument1, string argument3)
        {

            var xlsx_enumeration = new List<string>();

            // Recurse enumeration of original spreadsheets from input directory
            if (argument3 == "Recurse=Yes")
            {
                xlsx_enumeration = (List<string>)Directory.EnumerateFiles(argument1, "*.xlsx", SearchOption.AllDirectories)
                    .ToList();

                try
                {
                    foreach (var file in xlsx_enumeration)
                    {
                        SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(file, false);
                        bool strict = spreadsheet.StrictRelationshipFound;
                        spreadsheet.Close();

                        if (strict == true)
                        {
                            numXLSX_Strict++;
                        }
                    }
                }

                catch (InvalidDataException)
                {
                    conformance_count_fail++;
                }

                return numXLSX_Strict;
            }

            else
            {
                xlsx_enumeration = (List<string>)Directory.EnumerateFiles(argument1, "*.xlsx", SearchOption.TopDirectoryOnly)
                    .ToList();

                try
                {
                    foreach (var file in xlsx_enumeration)
                    {
                        SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(file, false);
                        bool strict = spreadsheet.StrictRelationshipFound;
                        spreadsheet.Close();

                        if (strict == true)
                        {
                            numXLSX_Strict++;
                        }
                    }
                }

                catch (InvalidDataException)
                {
                    conformance_count_fail++;
                }

                return numXLSX_Strict;

            }

        }

    }

}
