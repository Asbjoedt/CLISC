using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CLISC
{
    public partial class Conversion
    {
        // Convert to Office Open XML XLSX Transitional conformance - DOES NOT SUPPORT STRICT TO TRANSITIONAL
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

        // Convert XLSX Strict to Transitional conformance - WORK IN PROGRESS
        public bool Convert_Strict_to_Transitional(string input_filepath, string output_filepath, string file_folder)
        {
            using (var spreadsheet = SpreadsheetDocument.Open(input_filepath, false))
            {
                // Check for Strict conformance class
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                var conformance = wbPart.Workbook.Conformance;

                // If Strict, transform
                if (conformance != null || conformance == "transitional")
                {

                }

                bool convert_success = true;
                return convert_success;
            }
        }
    }
}
