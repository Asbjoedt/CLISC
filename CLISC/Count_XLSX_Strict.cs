using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace CLISC
{
    public partial class Spreadsheet
    {
        public int numCONFORM_fail = 0;

        // Count XLSX Strict conformance
        public int Count_XLSX_Strict(string argument1, string argument3)
        {
            DirectoryInfo count = new DirectoryInfo(argument1);
            string[] xlsx_files;
            if (argument3 == "Recurse=Yes")
            {
                xlsx_files = Directory.GetFiles(argument1,"*.xlsx", SearchOption.AllDirectories);
            }
            else
            {
                xlsx_files = Directory.GetFiles(argument1, "*.xlsx", SearchOption.TopDirectoryOnly);
            }
            try
            {
                foreach (var xlsx in xlsx_files)
                {
                    SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(xlsx, false);
                    bool? strict = spreadsheet.StrictRelationshipFound;
                    spreadsheet.Close();
                    if (strict == true)
                    {
                        numXLSX_Strict++;
                    }
                }
            }

            // Catch exceptions, when spreadsheet cannot be opened due to password protection or corruption
            catch (InvalidDataException)
            {
                numCONFORM_fail++;
            }
            catch (OpenXmlPackageException)
            {
                numCONFORM_fail++;
            }

            // Return number of Strict conformant xlsx files
            return numXLSX_Strict;
        }
    }
}
