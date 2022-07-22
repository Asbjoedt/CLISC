using System;
using System.IO;


namespace CLISC
{

    public class Program
    {

        public static void Main(string[] args)
        {

            Console.WriteLine("CLISC - Command Line Interface Spreadsheet Count Convert & Compare");
            Console.WriteLine("@Asbjørn Skødt, web: https://github.com/Asbjoedt/CLISC");
            Console.WriteLine("---");

            string argument1 = Convert.ToString(args[1]);
            string argument2 = Convert.ToString(args[2]);
            string argument3 = Convert.ToString(args[3]);
            string argument4 = Convert.ToString(args[4]);
            string results_directory = "";

            // Object reference
            Spreadsheet process = new Spreadsheet();

            try
            {

                // Validate recurse and archive arguments
                if (argument3 == "Recurse=Yes" || argument3 == "Recurse=No")
                {
                    if (argument4 == "Archive=Yes" || argument4 == "Archive=No")
                    {

                        // Method references
                        switch (args[0])
                        {
                            case "Count":
                                results_directory = process.Count(argument1, argument2, argument3);
                                break;

                            case "Count&Convert":
                                results_directory = process.Count(argument1, argument2, argument3);
                                process.Convert(argument1, results_directory, argument3, argument4);
                                break;

                            case "Count&Convert&Compare":
                                results_directory = process.Count(argument1, argument2, argument3);
                                process.Convert(argument1, results_directory, argument3, argument4);
                                process.Compare(argument1, results_directory, argument3, argument4);
                                break;

                            case "Count&Convert&Compare&Archive":
                                results_directory = process.Count(argument1, argument2, argument3);
                                process.Convert(argument1, results_directory, argument3, argument4);
                                process.Compare(argument1, results_directory, argument3, argument4);
                                process.Archive(argument1, argument2, results_directory);
                                break;

                            default:
                                Console.WriteLine("Invalid first argument. First argument must be one these: Count, Count&Convert, Count&Convert&Compare, Count&Convert&Compare&Archive");
                                break;
                        }

                    }

                    // Inform user of invalid archive argument
                    else
                    {
                        Console.WriteLine("Invalid archive argument. It must be one of these Archive=Yes or Archive=No");
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

            finally
            {
                // Inform user of end of CLISC
                Console.WriteLine("CLISC has finished");
                Console.WriteLine("---");
            }

        }

    }

}