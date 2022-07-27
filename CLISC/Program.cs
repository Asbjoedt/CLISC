using System;
using System.IO;
using System.CommandLine;


namespace CLISC
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Inform user of beginning of program
            Console.WriteLine("CLISC - Command Line Interface Spreadsheet Count Convert & Compare");
            Console.WriteLine("@Asbjørn Skødt, web: https://github.com/Asbjoedt/CLISC");
            Console.WriteLine("---");

            try
            {
                // Data types
                string argument0 = Convert.ToString(args[0]);
                string argument1 = Convert.ToString(args[1]);
                string argument2 = Convert.ToString(args[2]);
                string argument3 = Convert.ToString(args[3]);
                string Results_Directory = "";

                // Object reference
                Spreadsheet process = new Spreadsheet();
                Program report = new Program();

                // Validate recurse and archive arguments
                if (argument3 == "Recurse=Yes" || argument3 == "Recurse=No")
                {
                    // Method references
                    switch (args[0])
                    {
                        case "Count":
                            process.Count(argument1, argument2, argument3);
                            report.Final_Results();
                            break;

                        case "Count&Convert":
                            Results_Directory = process.Count(argument1, argument2, argument3);
                            process.Convert(argument0, argument1, argument3, Results_Directory);
                            report.Final_Results();
                            break;

                        case "Count&Convert&Compare":
                            Results_Directory = process.Count(argument1, argument2, argument3);
                            List<fileIndex> File_List = process.Convert(argument0, argument1, argument3, Results_Directory);
                            process.Compare(argument0, argument1, Results_Directory, File_List);
                            report.Final_Results();
                            break;

                        case "Count&Convert&Compare&Archive":
                            Results_Directory = process.Count(argument1, argument2, argument3);
                            File_List = process.Convert(argument0, argument1, argument3, Results_Directory);
                            process.Compare(argument0, argument1, Results_Directory, File_List);
                            process.Archive(Results_Directory, File_List);
                            report.Final_Results();
                            break;

                        default:
                            Console.WriteLine("Invalid first argument. First argument must be one these: Count, Count&Convert, Count&Convert&Compare, Count&Convert&Compare&Archive");
                            break;
                    }
                }
                // Inform user of invalid recurse argument
                else
                {
                    Console.WriteLine("Invalid recurse argument. It must be one of these: Resurse=Yes or Recurse=No");
                }
            }
            // Inform user of argument errors
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("The number of arguments are invalid");
            }
        }

        public void Final_Results()
        {
            Console.WriteLine("CLISC FINAL RESULTS");
            Console.WriteLine("---");
            Console.WriteLine($"{Spreadsheet.numTOTAL} spreadsheets counted");
            Console.WriteLine($"{Spreadsheet.numCOMPLETE} out of {Spreadsheet.numTOTAL} spreadsheets completed conversion");
            Console.WriteLine($"--> {Spreadsheet.numFAILED} spreadsheets failed conversion");
            Console.WriteLine($"{Spreadsheet.numTOTAL_compare} out of {Spreadsheet.numTOTAL_conv} converted spreadsheets were compared");
            Console.WriteLine("CLISC ended");
            Console.WriteLine("---");
        }
    }
}