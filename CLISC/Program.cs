using System;
using System.IO;


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

            // Data types
            string argument0 = Convert.ToString(args[0]);
            string argument1 = Convert.ToString(args[1]);
            string argument2 = Convert.ToString(args[2]);
            string argument3 = Convert.ToString(args[3]);
            string Results_Directory = "";
 
            // Object reference
            Spreadsheet process = new Spreadsheet();
            List<fileIndex> File_List = new List<fileIndex>();

            try
            {
                // Validate recurse and archive arguments
                if (argument3 == "Recurse=Yes" || argument3 == "Recurse=No")
                {
                    // Method references
                    switch (args[0])
                    {
                        case "Count":
                            process.Count(argument1, argument2, argument3);
                            break;

                        case "Count&Convert":
                            Results_Directory = process.Count(argument1, argument2, argument3);
                            process.Convert(argument0, argument1, argument3, Results_Directory);
                            break;

                        case "Count&Convert&Compare":
                            Results_Directory = process.Count(argument1, argument2, argument3);
                            File_List = process.Convert(argument0, argument1, argument3, Results_Directory);
                            process.Compare(argument0, argument1, Results_Directory, File_List);
                            break;

                        case "Count&Convert&Compare&Archive":
                            Results_Directory = process.Count(argument1, argument2, argument3);
                            File_List = process.Convert(argument0, argument1, argument3, Results_Directory);
                            process.Compare(argument0, argument1, Results_Directory, File_List);
                            process.Archive(argument0, argument1, argument2, Results_Directory, File_List);
                            break;

                        default:
                            Console.WriteLine("Invalid first argument. First argument must be one these: Count, Count&Convert, Count&Convert&Compare, Count&Convert&Compare&Archive");
                            break;
                    }
                }
                // Inform user of invalid recurse argument
                else
                {
                    Console.WriteLine("Invalid recurse argument. It must be one of these Resurse=Yes or Recurse=No");
                }
            }
            // Inform user of argument errors
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("The number of arguments are invalid. Consult GitHub documentation");
            }
            // Inform user of end of CLISC
            finally
            {
                Console.WriteLine("CLISC has finished");
                Console.WriteLine("---");
            }
        }
    }
}