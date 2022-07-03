using System;
using System.IO;

namespace CLISC
{

    public class Program
    {

        public static void Main(string[] args)
        {
            
            string argument1 = Convert.ToString(args[1]);
            string argument2 = Convert.ToString(args[2]);
            string argument3 = Convert.ToString(args[3]);

            // Object reference
            Spreadsheet process = new Spreadsheet();

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

        }

    }

}