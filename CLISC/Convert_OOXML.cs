﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace CLISC
{
    public partial class Conversion
    {
        // Convert to .xlsx Transitional - DOES NOT SUPPORT STRICT TO TRANSITIONAL
        public bool Convert_to_OOXML_Transitional(string input_filepath, string output_filepath)
        {
            byte[] byteArray = File.ReadAllBytes(input_filepath);
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(stream, true))
                {
                    spreadsheetDoc.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                }
                File.WriteAllBytes(output_filepath, stream.ToArray());
            }
            bool convert_success = true;
            return convert_success;
        }

        // Convert .xlsx Strict to Transitional conformance
        public bool Convert_Strict_to_Transitional(string input_filepath)
        {
            string namespace_xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            string namespace_xmlns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            string namespace_app_xlmns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
            string namespace_app_xmlns_vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
            string namespace_xmlns_dc = ""; // not relevant
            string namespace_xmlns_dcterms = ""; // not relevant
            string namespace_xmlns_dcmitype = ""; // not relevant
            string namespace_xmlns_a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            using (var spreadsheet = SpreadsheetDocument.Open(input_filepath, true))
            {
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Workbook workbook = wbPart.Workbook;
                // If Strict, transform
                if (workbook.Conformance != null || workbook.Conformance != "transitional")
                {
                    // Change conformance class
                    workbook.Conformance.Value = ConformanceClass.Enumtransitional;
                    // Change namespaces in /xl/workbook.xml
                    workbook.RemoveNamespaceDeclaration("x");
                    workbook.AddNamespaceDeclaration("x", namespace_xmlns);
                    workbook.RemoveNamespaceDeclaration("r");
                    workbook.AddNamespaceDeclaration("r", namespace_xmlns_r);
                    // Change namespaces in /xl/worksheets/worksheet[n+1].xml
                    List<WorksheetPart> worksheets = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                    if (worksheets.Count > 0)
                    {
                        foreach (WorksheetPart worksheet in worksheets)
                        {
                            worksheet.Worksheet.RemoveNamespaceDeclaration("x");
                            worksheet.Worksheet.AddNamespaceDeclaration("x", namespace_xmlns);
                            worksheet.Worksheet.RemoveNamespaceDeclaration("r");
                            worksheet.Worksheet.AddNamespaceDeclaration("r", namespace_xmlns_r);
                            worksheet.Worksheet.RemoveNamespaceDeclaration("v");
                        }
                    }
                    // Change namespaces in /xl/styles.xml
                    wbPart.WorkbookStylesPart.Stylesheet.RemoveNamespaceDeclaration("x");
                    wbPart.WorkbookStylesPart.Stylesheet.AddNamespaceDeclaration("x", namespace_xmlns);

                    // Change namespaces in /xl/sharedStrings.xml
                    if (wbPart.SharedStringTablePart != null)
                    {
                        wbPart.SharedStringTablePart.SharedStringTable.RemoveNamespaceDeclaration("x");
                        wbPart.SharedStringTablePart.SharedStringTable.AddNamespaceDeclaration("x", namespace_xmlns);
                    }


                    // Change namespaces in /xl/embeddings

                    // Change namespaces in /xl/externallinks
                    List<ExternalWorkbookPart> extwbParts = wbPart.ExternalWorkbookParts.ToList();
                    if (extwbParts.Count > 0)
                    {
                        foreach (ExternalWorkbookPart extwbPart in extwbParts)
                        {

                        }
                    }
                    // Change namespaces in /docProps/app.xml


                    // Change namespaces in /docProps/core.xml


                    // Change namespaces in /xl/theme/theme[n+1].xml
                    wbPart.ThemePart.Theme.RemoveNamespaceDeclaration("a");
                    wbPart.ThemePart.Theme.AddNamespaceDeclaration("a", namespace_xmlns_a);
                }
            }
            bool convert_success = true;
            return convert_success;
        }

        // Convert .xlsx Transtional to Strict
        public void Convert_Transitional_to_Strict(string input_filepath)
        {
            string namespace_xmlns = "http://purl.oclc.org/ooxml/spreadsheetml/main";
            string namespace_xmlns_r = "http://purl.oclc.org/ooxml/officeDocument/relationships";
            string namespace_app_xlmns = "http://purl.oclc.org/ooxml/officeDocument/extendedProperties";
            string namespace_app_xmlns_vt = "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes";
            string namespace_core_xmlns_dc = "http://purl.org/dc/elements/1.1/"; // not relevant
            string namespace_core_xmlns_dcterms = "http://purl.org/dc/terms/"; // not relevant
            string namespace_core_xmlns_dcmitype = "http://purl.org/dc/dcmitype/"; // not relevant
            string namespace_xmlns_a = "http://purl.oclc.org/ooxml/drawingml/main";
            string namespace_xmlns_v = "urn:schemas-microsoft-com:vml";
            string namespace_rel_styles = "http://purl.oclc.org/ooxml/officeDocument/relationships/styles";
            string namespace_rel_themes = "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";
            string namespace_rel_worksheet = "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet";
            string namespace_rel_sharedstrings = "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings";
            string namespace_rel_externallink = "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLink";
            string namespace_rel_workbook = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
            string namespace_externallink_externallinkpath = "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLinkPath";
            string namespace_drawing_xmlns_xdr = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";

            using (var spreadsheet = SpreadsheetDocument.Open(input_filepath, true))
            {
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Workbook workbook = wbPart.Workbook;
                // If Transitional, transform
                if (workbook.Conformance == null || workbook.Conformance != "strict")
                {
                    // Change conformance class
                    workbook.Conformance.Value = ConformanceClass.Enumstrict;
                    // Change namespaces in /xl/workbook.xml
                    workbook.RemoveNamespaceDeclaration("x");
                    workbook.AddNamespaceDeclaration("x", namespace_xmlns);
                    workbook.RemoveNamespaceDeclaration("r");
                    workbook.AddNamespaceDeclaration("r", namespace_xmlns_r);
                    // Change namespaces in /xl/worksheets/worksheet[n+1].xml
                    List<WorksheetPart> worksheets = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                    if (worksheets.Count > 0)
                    {
                        foreach (WorksheetPart worksheet in worksheets)
                        {
                            worksheet.Worksheet.RemoveNamespaceDeclaration("x");
                            worksheet.Worksheet.AddNamespaceDeclaration("x", namespace_xmlns);
                            worksheet.Worksheet.RemoveNamespaceDeclaration("r");
                            worksheet.Worksheet.AddNamespaceDeclaration("r", namespace_xmlns_r);
                            worksheet.Worksheet.AddNamespaceDeclaration("v", namespace_xmlns_v);
                        }
                    }
                    // Change namespaces in /xl/styles.xml
                    if (wbPart.WorkbookStylesPart.Stylesheet != null)
                    {
                        wbPart.WorkbookStylesPart.Stylesheet.RemoveNamespaceDeclaration("x");
                        wbPart.WorkbookStylesPart.Stylesheet.AddNamespaceDeclaration("x", namespace_xmlns);
                    }
                    // Change namespaces in /xl/sharedStrings.xml
                    if (wbPart.SharedStringTablePart != null)
                    {
                        wbPart.SharedStringTablePart.SharedStringTable.RemoveNamespaceDeclaration("x");
                        wbPart.SharedStringTablePart.SharedStringTable.AddNamespaceDeclaration("x", namespace_xmlns);
                    }
                    // Change namespaces in /xl/embeddings

                    // Change namespaces in /xl/externallinks
                    List<ExternalWorkbookPart> extwbParts = wbPart.ExternalWorkbookParts.ToList();
                    if (extwbParts.Count > 0)
                    {
                        foreach (ExternalWorkbookPart extwbPart in extwbParts)
                        {
                            extwbPart.RootElement.RemoveNamespaceDeclaration("x");
                            extwbPart.RootElement.AddNamespaceDeclaration("x", namespace_xmlns);
                            extwbPart.RootElement.RemoveNamespaceDeclaration("r");
                            extwbPart.RootElement.AddNamespaceDeclaration("r", namespace_xmlns_r);
                        }
                    }
                    // Change namespaces in /docProps/app.xml


                    // Change namespaces in /docProps/core.xml


                    // Change namespaces in /xl/theme/theme[n+1].xml
                    if (wbPart.ThemePart.Theme != null)
                    {
                        wbPart.ThemePart.Theme.RemoveNamespaceDeclaration("a");
                        wbPart.ThemePart.Theme.AddNamespaceDeclaration("a", namespace_xmlns_a);
                    }
                    // Change namespaces in /xl/calcChain.xml
                    if (wbPart.CalculationChainPart.CalculationChain != null)
                    {
                        wbPart.CalculationChainPart.CalculationChain.RemoveNamespaceDeclaration("x");
                        wbPart.CalculationChainPart.CalculationChain.AddNamespaceDeclaration("x", namespace_xmlns);
                    }
                }
            }
        }

        // Convert .xlsx Transtional to Strict using Excel
        public bool Convert_Transitional_to_Strict_ExcelInterop(string input_filepath, string output_filepath)
        {
            bool convert_success = false;

            Excel.Application app = new Excel.Application(); // Create Excel object instance
            app.DisplayAlerts = false; // Don't display any Excel prompts
            Excel.Workbook wb = app.Workbooks.Open(input_filepath); // Create workbook instance

            wb.SaveAs(output_filepath, 61); // Save workbook as .xlsx Strict
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
    }
}
