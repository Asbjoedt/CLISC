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

        public bool Convert_OOXML(string copy_filepath)
        {
            byte[] byteArray = File.ReadAllBytes(copy_filepath);
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(stream, true))
                {
                    spreadsheetDoc.ChangeDocumentType(DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
                }
                File.WriteAllBytes(conv_filepath, stream.ToArray());
            }

            convert_success = true;

            return convert_success;
        }

    }

}
