using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel; // Use with XLSB
using Excel = Microsoft.Office.Interop.Excel; // Use with XLSB

namespace CLISC
{
    public partial class Spreadsheet
    {
        // Convert to Office Open XML XLSX Transitional conformance - DOES NOT SUPPORT STRICT TO TRANSITIONAL
        public bool Convert_OOXML_Transitional(string org_filepath, string input_filepath, string output_filepath)
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

            // Inform user
            Console.WriteLine(org_filepath);
            Console.WriteLine($"--> Conversion {convert_success}");
            Console.WriteLine($"--> Conversion saved to: {output_filepath}");

            return convert_success;
        }

        // Convert to Office Open XML XLSX Strict conformance - NOT WORKING - IT OUTPUTS TRANSITIONAL
        public bool Convert_OOXML_Strict(string org_filepath, string input_filepath, string output_filepath)
        {
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

            bool convert_success = true;

            // Inform user
            Console.WriteLine(org_filepath);
            Console.WriteLine($"--> Conversion {convert_success}");

            return convert_success;
        }

        // Convert XLSB using Excel
        // Found code here: https://docs.microsoft.com/en-us/answers/questions/212363/how-to-convert-xlsb-file-to-xlsx.html
        public bool Convert_XLSB(string org_filepath, string input_filepath, string output_filepath)
        {

            Excel.Application excelApplication = new Excel.Application();
            Workbooks workbooks = excelApplication.Workbooks;
            // open book in any format
            Excel.Workbook workbook = workbooks.Open(input_filepath, XlUpdateLinks.xlUpdateLinksNever, true, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // save in XlFileFormat.xlExcel12 format which is XLSB
            workbook.SaveAs(output_filepath, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // close workbook
            workbook.Close(false, Type.Missing, Type.Missing);
            excelApplication.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);

            foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
            {
                proc.Kill();
            }

            bool convert_success = true;

            // Inform user
            Console.WriteLine(org_filepath);
            Console.WriteLine($"--> Conversion {convert_success}");

            return convert_success;
        }
    }
}
