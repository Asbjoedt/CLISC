using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CLISC
{
    public partial class Count
    {
        public static int numCONFORM_fail = 0;

        // Count XLSX Strict conformance
        public int Count_OOXML_Conformance(string inputdir, bool recurse, string conformance)
        {
            int count = 0;
            string[] xlsx_files = {""};

            // Search recursively or not
            SearchOption searchoption = SearchOption.TopDirectoryOnly;
            if (recurse == true)
            {
                searchoption = SearchOption.AllDirectories;
            }

            // Create index of xlsx files
            xlsx_files = Directory.GetFiles(inputdir, "*.xlsx", searchoption);

            // Open each spreadsheet to check for conformance
            try
            {
                // Count Transitional
                if (conformance == "transitional")
                {
                    foreach (var xlsx in xlsx_files)
                    {
                        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(xlsx, false))
                        {
                            Workbook workbook = spreadsheet.WorkbookPart.Workbook;
                            if (workbook.Conformance == null || workbook.Conformance == "transitional")
                            {
                                count++;
                            }
                        }
                    }
                }
                // Count Strict
                else if (conformance == "strict")
                {
                    foreach (var xlsx in xlsx_files)
                    {
                        SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(xlsx, false);
                        bool? strict = spreadsheet.StrictRelationshipFound;
                        spreadsheet.Close();
                        if (strict == true)
                        {
                            count++;
                        }
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

            // Return count
            return count;
        }

        // Count Strict conformance
        public int Count_Strict(string inputdir, bool recurse)
        {
            int count = 0;
            string[] xlsx_files = { "" };

            // Search recursively or not
            SearchOption searchoption = SearchOption.TopDirectoryOnly;
            if (recurse == true)
            {
                searchoption = SearchOption.AllDirectories;
            }

            // Create index of xlsx files
            xlsx_files = Directory.GetFiles(inputdir, "*.xlsx", searchoption);

            try
            {
                foreach (var xlsx in xlsx_files)
                {
                    SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(xlsx, false);
                    bool? strict = spreadsheet.StrictRelationshipFound;
                    spreadsheet.Close();
                    if (strict == true)
                    {
                        count++;
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
            return count;
        }
    }
}
