using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CLISC
{
    public partial class Conversion
    {
        // Check OOXML format validity
        public void CheckOOXMLFormat(string filepath)
        {
            // Quick ZIP/OOXML header check
            using (var fs = new FileStream(filepath, FileMode.Open, FileAccess.Read))
            {
                if (fs.Length < 4) throw new FileFormatException("File too small to be OOXML.");
                byte[] header = new byte[4];
                fs.Read(header, 0, 4);
                if (header[0] != 0x50 || header[1] != 0x4B) // 'P' 'K'
                    throw new FileFormatException("File is not a valid ZIP/OOXML package.");
            }
        }

        // Check if spreadsheet is writeable
        public void CheckWriteAbility(string filepath)
        {
            // Check OOXML format validity
            CheckOOXMLFormat(filepath);

            try
            {
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    // Validate workbook part exists
                    if (spreadsheet.WorkbookPart?.Workbook == null)
                        throw new FileFormatException("Missing workbook part; file is corrupted or not an OOXML spreadsheet.");

                    // Check if spreadsheet is protected
                    if (spreadsheet.WorkbookPart.Workbook.WorkbookProtection != null ||
                        spreadsheet.WorkbookPart.Workbook.FileSharing != null)
                    {
                        throw new FileFormatException("Workbook is protected or shared.");
                    }
                }
            }
            catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException)
            {
                throw new FileFormatException("File is not a valid Open XML package or is corrupted.");
            }
        }

        // Convert to .xlsx Transitional
        public bool ConvertToXLSX(string input_filepath, string output_filepath)
        {
            bool convert_success = false;

            // Check if file is writeable
            CheckWriteAbility(input_filepath);

            // Convert spreadsheet
            byte[] byteArray = File.ReadAllBytes(input_filepath);
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, true))
                {
                    //Convert
                    spreadsheet.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                }
                File.WriteAllBytes(output_filepath, stream.ToArray());
            }

            // Repair spreadsheet
            Repair rep = new Repair();
            rep.Repair_OOXML(output_filepath);

            // Return success
            return convert_success = true;
        }

        // Work in progress
        // Convert .xlsx Strict to Transitional conformance
        public void ConvertStrictToTransitional(string input_filepath)
        {
            // Create list of namespaces
            List<namespaceIndex> namespaces = namespaceIndex.Create_Namespaces_Index();

            using (var spreadsheet = SpreadsheetDocument.Open(input_filepath, true))
            {
                WorkbookPart? wbPart = spreadsheet.WorkbookPart;
                Workbook? workbook = wbPart?.Workbook;
                // If Strict
                if (workbook?.Conformance != null || (workbook?.Conformance != null && workbook.Conformance != "transitional"))
                {
                    // Change conformance class
                    workbook.Conformance.Value = ConformanceClass.Enumtransitional;

                    // Remove vml urn namespace from workbook.xml
                    workbook.RemoveNamespaceDeclaration("v");
                }
            }
        }
    }
}