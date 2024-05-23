using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CLISC
{
    public partial class Count
    {
        // Count XLSX Strict conformance
        public Tuple<int, int, int> Count_OOXML_Conformance(string inputdir, bool recurse)
        {
            int transitional_count = 0;
            int strict_count = 0;
            int unknown_count = 0;
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
                    using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(xlsx, false))
                    {
                        bool? strict = spreadsheet.StrictRelationshipFound;
                        if (strict == true)
                        {
                            strict_count++;
                        }
                    }
                }
            }
            // Catch exceptions, when spreadsheet cannot be opened due to password protection or corruption
            catch (InvalidDataException)
            {
                unknown_count++;
            }
            catch (OpenXmlPackageException)
            {
                unknown_count++;
            }
            catch (System.IO.FileFormatException)
            {
                unknown_count++;
            }

            // Calculate transitional
            transitional_count = xlsx_files.Count() - strict_count - unknown_count;

            // Return counts
            return System.Tuple.Create(transitional_count, strict_count, unknown_count);
        }

        // Alternative method for counting OOXML conformance
        public Tuple<int, int, int> Count_OOXML_Conformance_Alt(string inputdir, bool recurse, string conformance)
        {
            int transitional_count = 0;
            int strict_count = 0;
            int unknown_count = 0;
            string[] xlsx_files = { "" };

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
                                transitional_count++;
                            }
                        }
                    }
                }
                // Count Strict
                else if (conformance == "strict")
                {
                    foreach (var xlsx in xlsx_files)
                    {
                        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(xlsx, false))
                        {
                            bool? strict = spreadsheet.StrictRelationshipFound;
                            if (strict == true)
                            {
                                strict_count++;
                            }
                        }
                    }
                }

            }

            // Catch exceptions, when spreadsheet cannot be opened due to password protection or corruption
            catch (InvalidDataException)
            {
                unknown_count++;
            }
            catch (OpenXmlPackageException)
            {
                unknown_count++;
            }
            catch (System.IO.FileFormatException)
            {
                unknown_count++;
            }

            // Calculate transitional
            transitional_count = xlsx_files.Count() - strict_count - unknown_count;

            // Return count
            return System.Tuple.Create(transitional_count, strict_count, unknown_count);
        }
    }
}
