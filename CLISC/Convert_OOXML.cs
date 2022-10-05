using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using CLISC;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlSchemaReferences;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace CLISC
{
    public partial class Conversion
    {
        // Convert to .xlsx Transitional - DOES NOT SUPPORT STRICT TO TRANSITIONAL
        public bool Convert_to_OOXML_Transitional(string input_filepath, string output_filepath)
        {
            bool convert_success = false;

            // If password-protected or reserved by another user
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(input_filepath, false))
            {
                if (spreadsheet.WorkbookPart.Workbook.WorkbookProtection != null || spreadsheet.WorkbookPart.Workbook.FileSharing != null)
                {
                    return convert_success;
                }
            }

            // Convert spreadsheet
            byte[] byteArray = File.ReadAllBytes(input_filepath);
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, true))
                {
                    spreadsheet.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                }
                File.WriteAllBytes(output_filepath, stream.ToArray());
            }

            // Use hacks because of errors in Open XML SDK
            if (System.IO.Path.GetExtension(input_filepath) == ".xlsm" || System.IO.Path.GetExtension(input_filepath) == ".XLSM" || System.IO.Path.GetExtension(input_filepath) == ".xltm" || System.IO.Path.GetExtension(input_filepath) == ".XLTM")
            {
                Repair rep = new Repair();
                rep.Repair_VBA(output_filepath);
            }

            // Return success
            convert_success = true;
            return convert_success;
        }

        // Convert .xlsx Strict to Transitional conformance
        public void Convert_Strict_to_Transitional(string input_filepath)
        {
            // Create list of namespaces
            List<namespaceIndex> namespaces = namespaceIndex.Create_Namespaces_Index();

            using (var spreadsheet = SpreadsheetDocument.Open(input_filepath, true))
            {
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Workbook workbook = wbPart.Workbook;
                // If Strict
                if (workbook.Conformance != null || workbook.Conformance != "transitional")
                {
                    // Change conformance class
                    workbook.Conformance.Value = ConformanceClass.Enumtransitional;

                    // Remove vml urn namespace from workbook.xml
                    workbook.RemoveNamespaceDeclaration("v");
                }
            }
        }

        // Convert .xlsx Transtional to Strict
        public void Convert_Transitional_to_Strict(string input_filepath)
        {
            // Create list of namespaces
            List<namespaceIndex> namespaces = namespaceIndex.Create_Namespaces_Index();

            using (var spreadsheet = SpreadsheetDocument.Open(input_filepath, true))
            {
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Workbook workbook = wbPart.Workbook;
                // If Transitional
                if (workbook.Conformance == null || workbook.Conformance != "strict")
                {
                    // Change conformance class
                    workbook.Conformance.Value = ConformanceClass.Enumstrict;

                    // Add vml urn namespace to workbook.xml
                    workbook.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
                }
            }
        }

        // Convert .xlsx Transtional to Strict using Excel
        public bool Convert_Transitional_to_Strict_ExcelInterop(string input_filepath, string output_filepath)
        {
            bool convert_success = false;

            Excel.Application app = new Excel.Application(); // Create Excel object instance
            app.DisplayAlerts = false; // Don't display any Excel prompts
            Excel.Workbook wb = app.Workbooks.Open(input_filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

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