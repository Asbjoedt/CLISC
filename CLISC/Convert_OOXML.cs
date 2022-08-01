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
using System.IO.Compression; // Use with Transitional to Strict

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


            bool convert_success = true;
            return convert_success;
        }

        // Convert XLSB using Excel
        // Found code here: https://docs.microsoft.com/en-us/answers/questions/212363/how-to-convert-xlsb-file-to-xlsx.html
        // NOT USED IN PROGRAM - it needs Excel installed
        public bool Convert_XLSB(string org_filepath, string input_filepath, string output_filepath)
        {
            // Create object instance
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

            //System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);

            //foreach (system.diagnostics.process proc in system.diagnostics.process.getprocessesbyname("excel"))
            //{
            //    proc.kill();
            //}

            bool convert_success = true;
            return convert_success;
        }
    }
}
