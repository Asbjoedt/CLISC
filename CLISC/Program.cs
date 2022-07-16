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

            if (argument1 == argument2)
            {
                Console.WriteLine("Error: Input and output directories cannot be the same");
            }
            else
            {
                // Object reference
                Spreadsheet process = new Spreadsheet();

                // Validate recursive argument
                if (argument3 == "Recursive=Yes" || argument3 == "Recursive=No")

                    // Method references
                    switch (args[0])
                    {
                        case "Count":
                            process.Count(argument1, argument2, argument3);
                            break;
                        case "Count&Convert":
                            process.Count(argument1, argument2, argument3);
                            process.Convert(argument1, argument2, argument3);
                            break;
                        case "Count&Convert&Compare":
                            process.Count(argument1, argument2, argument3);
                            process.Convert(argument1, argument2, argument3);
                            process.Compare(argument1, argument2, argument3);
                            break;
                        default:
                            Console.WriteLine("Invalid first argument. First argument must be one these: Count, Count&Convert, Count&Convert&Compare");
                            break;
                    }

                // Inform user of invalid recursive argument
                else
                {
                    Console.WriteLine("Invalid recursive argument");
                }

            }

        }

    }

}