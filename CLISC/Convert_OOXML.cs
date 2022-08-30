using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
        public bool Convert_Strict_to_Transitional(string input_filepath, string output_filepath)
        {
            string namespace_xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            string namespace_xmlns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            using (var spreadsheet = SpreadsheetDocument.Open(input_filepath, true))
            {
                Workbook workbook = spreadsheet.WorkbookPart.Workbook;
                // If Strict, transform
                if (workbook.Conformance != null || workbook.Conformance != "transitional")
                {
                    // Change conformance class
                    workbook.Conformance.InnerText = "transitional";
                    // Change namespaces in Workbook
                    workbook.RemoveNamespaceDeclaration("xmlns");
                    workbook.AddNamespaceDeclaration("xmlns", namespace_xmlns);
                    workbook.RemoveNamespaceDeclaration("xmlns:r");
                    workbook.AddNamespaceDeclaration("xmlns:r", namespace_xmlns_r);
                    // Change namespaces in worksheets
                    List<WorksheetPart> worksheets = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                    foreach (WorksheetPart worksheet in worksheets)
                    {
                        worksheet.Worksheet.RemoveNamespaceDeclaration("xmlns");
                        worksheet.Worksheet.AddNamespaceDeclaration("xmlns", namespace_xmlns);
                        worksheet.Worksheet.RemoveNamespaceDeclaration("xmlns:r");
                        worksheet.Worksheet.AddNamespaceDeclaration("xmlns:r", namespace_xmlns_r);
                        worksheet.Worksheet.RemoveNamespaceDeclaration("xmlns:v");
                    }
                    // Change namespaces in stylesheet

                }
                bool convert_success = true;
                return convert_success;
            }
        }

        // Convert .xlsx Transtional to Strict
        public bool Convert_Transitional_to_Strict(string input_filepath, string output_filepath)
        {
            string namespace_xmlns = "http://purl.oclc.org/ooxml/spreadsheetml/main";
            string namespace_xmlns_r = "http://purl.oclc.org/ooxml/officeDocument/relationships";

            using (var spreadsheet = SpreadsheetDocument.Open(input_filepath, true))
            {
                Workbook workbook = spreadsheet.WorkbookPart.Workbook;
                // If Transitional, transform
                if (workbook.Conformance == null || workbook.Conformance == "transitional")
                {
                    // Change conformance class
                    workbook.Conformance.InnerText = "strict";

                }
            }
            bool convert_success = true;
            return convert_success;
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
