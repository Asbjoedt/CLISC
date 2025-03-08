using System.Text;

namespace CLISC
{
    public partial class Count
    {
        // Public data types
        public static int numTOTAL;

        // Count spreadsheets
        public string Count_Spreadsheets(string inputdir, string outputdir, bool recurse)
        {
            int numCONFORM_fail = 0;

            Console.WriteLine("COUNT");
            Console.WriteLine("---");

            //Object reference
            DirectoryInfo count = new DirectoryInfo(inputdir);
            fileFormatIndex index = new fileFormatIndex();
            List<fileFormatIndex> fileformats = index.Create_fileFormatIndex();

            // Search recursively or not
            SearchOption searchoption = SearchOption.TopDirectoryOnly;
            if (recurse == true)
				searchoption = SearchOption.AllDirectories;

			foreach (fileFormatIndex fileformat in fileformats)
            {
                // Count
                int total = count.GetFiles($"*{fileformat.Extension}", searchoption).Length;

                // Handle XLSX and detect OOXML conformance
                if (fileformat.Extension == ".xlsx")
                {
                    Tuple<int, int, int> counted_OOXML_conformance = Count_OOXML_Conformance(inputdir, recurse);

                    if (fileformat.Conformance == "transitional")
                        total = counted_OOXML_conformance.Item1;
                    else if (fileformat.Conformance == "strict")
                        total = counted_OOXML_conformance.Item2;
                    else if (fileformat.Conformance == "unknown") 
                    {
                        total = counted_OOXML_conformance.Item3;
                        numCONFORM_fail = counted_OOXML_conformance.Item3;
                    }
                }

                // Add counted value to the file format list
                fileformat.Count = total;

                if (fileformat.Conformance == null)
                {
                    // Add counted value to sum of all files
                    numTOTAL = numTOTAL + total;
                }
            }

            // Inform user if no spreadsheets identified
            if (numTOTAL == 0 && numCONFORM_fail == 0)
            {
                Console.WriteLine("No spreadsheets identified");
                Console.WriteLine("CLISC ended");
                Console.WriteLine("---");
                throw new Exception();
            }
            else
            {
                // Show count to user
                Console.WriteLine("# Extension - Name");
                foreach (fileFormatIndex fileformat in fileformats)
                {
                    if (fileformat.Conformance == null)
                        Console.WriteLine($"{fileformat.Count} {fileformat.Extension} - {fileformat.Description}");
                    else if (fileformat.Conformance != null)
                        Console.WriteLine($"--> {fileformat.Count} {fileformat.Extension} have {fileformat.Conformance} conformance");
                }
                Console.WriteLine($"{numTOTAL} spreadsheets in total");

                // Create new directory to output results in CSV
                Results res = new Results();
                string Results_Directory = res.Create_Results_Directory(outputdir);

                // Output results in CSV
                var csv = new StringBuilder();
                var newLine0 = string.Format("#;Extension;Name");
                csv.AppendLine(newLine0);
                foreach (fileFormatIndex fileformat in fileformats)
                {
                    var newLine1 = string.Format($"{fileformat.Count};{fileformat.Extension};{fileformat.Description}");
                    csv.AppendLine(newLine1);
                }
                var newLine2 = string.Format($"{numTOTAL};spreadshets in total;");
                csv.AppendLine(newLine2);

                // Close CSV
                Results.CSV_filepath = Results_Directory + "\\1_Count_Results.csv";
                File.WriteAllText(Results.CSV_filepath, csv.ToString(), Encoding.UTF8);

                return Results_Directory;
            }
        }
    }
}
