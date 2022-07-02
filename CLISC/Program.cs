using System;
using System.IO;

namespace CLISC
{

    class Program
    {

        public static void Main(string[] args)
        {

            // Object reference
            Spreadsheet process = new Spreadsheet();

            // Method references
            switch (args[0])
            {
                case "Count":
                    process.Count();
                    break;
                case "Count&Convert":
                    process.Count();
                    process.Convert();
                    break;
                case "Count&Convert&Compare":
                    process.Count();
                    process.Convert();
                    process.Compare();
                    break;
                default: throw new ArgumentException("Unknown argument", args[0]);
            }

        }

    }

}