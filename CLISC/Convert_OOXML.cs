using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlSchemaReferences;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CLISC
{
    public partial class Conversion
    {
        // Convert to .xlsx Transitional
        public bool Convert_to_OOXML_Transitional(string input_filepath, string output_filepath)
        {
            bool convert_success = false;

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

            // Return fail if spreadsheet is protected or filesharing is enabled
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(output_filepath, true))
            {
                if (spreadsheet.WorkbookPart.Workbook.WorkbookProtection != null || spreadsheet.WorkbookPart.Workbook.FileSharing != null)
                {
                    return convert_success;
                }
            }

            // Repair spreadsheet
            Repair rep = new Repair();
            rep.Repair_OOXML(output_filepath);


            // Convert Strict to Transitional using Excel Interop
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(output_filepath, true))
            {
                if (spreadsheet.StrictRelationshipFound)
                {
                    Convert_ExcelInterop(input_filepath, output_filepath);
                }
            }

            // Return success
            return convert_success = true;
        }

        // Work in progress
        // Convert .xlsx Strict to Transitional conformance
        public void Convert_Strict_to_Transitional(string input_filepath)
        {
            // Create list of namespaces
            List<namespaceIndex> namespaces = namespaceIndex.Create_Namespaces_Index();

            using (var spreadsheet = SpreadsheetDocument.Open(input_filepath, true))
            {
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                Workbook workbook = wbPart.Workbook;
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

        // Work in progress
        // Remove write or filesharing protection from spreadsheet in cases of no password
        public void Handle_Protection(SpreadsheetDocument spreadsheet, string input_filepath)
        {
            if (spreadsheet.WorkbookPart.Workbook.WorkbookProtection != null || spreadsheet.WorkbookPart.Workbook.FileSharing != null)
            {
                // Use Excel Interop to convert the spreadsheet
                Convert_ExcelInterop(input_filepath, input_filepath);
            }
        }
    }
}