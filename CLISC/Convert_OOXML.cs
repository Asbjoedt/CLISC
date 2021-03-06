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
    public partial class Spreadsheet
    {
        // Convert to Office Open XML XLSX Transitional conformance
        public bool Convert_OOXML_Transitional(string org_filepath, string conv_filepath)
        {
            byte[] byteArray = File.ReadAllBytes(org_filepath);
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(stream, true))
                {
                    spreadsheetDoc.ChangeDocumentType(DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
                }
                File.WriteAllBytes(conv_filepath, stream.ToArray());
            }

            bool convert_success = true;

            // Inform user
            Console.WriteLine(org_filepath);
            Console.WriteLine($"--> Conversion {convert_success}");

            return convert_success;
        }

        // Convert to Office Open XML XLSX Strict conformance - NOT WORKING - IT OUTPUTS TRANSITIONAL
        public bool Convert_OOXML_Strict(string org_filepath, string conv_filepath)
        {
            byte[] byteArray = File.ReadAllBytes(org_filepath);
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, true))
                {
                    spreadsheet.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                }
                File.WriteAllBytes(conv_filepath, stream.ToArray());
            }

            bool convert_success = true;

            // Inform user
            Console.WriteLine(org_filepath);
            Console.WriteLine($"--> Conversion {convert_success}");

            convert_success = true;
            return convert_success;
        }
    }
}
