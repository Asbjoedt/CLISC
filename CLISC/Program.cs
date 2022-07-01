﻿using System;
using System.IO;

namespace CLISC
{

    class Program
    {

        public static void Main(string[] args)
        {
            // User guidance
            Console.WriteLine("CLISC - Command Line Interface Spreadsheet Count, Convert & Compare");
            Console.WriteLine();
            Console.WriteLine("Program behavior:");
            Console.WriteLine("\tCount Excel spreadsheets in directory by file format");
            Console.WriteLine("\tConvert XLS, XLT, XLAM, XLSB, XLTX, XLSM, XLTM to XLSX (OOXML Transitional conformance)");
            Console.WriteLine("\tOutput all conversions in a new directory with new subdirectories named n+1");
            Console.WriteLine("\tRename all conversions n+1.xlsx");
            Console.WriteLine("\tCompare the results to log workbook and checksum differences between input and output file formats");
            Console.WriteLine("\tOutput log in CSV");
            Console.WriteLine();
            Console.WriteLine("Available arguments:");
            Console.WriteLine("\tCount 'Filepath to directory' -Recursive");
            Console.WriteLine("\tCount&Convert 'Filepath to directory' -Recursive");
            Console.WriteLine("\tCount&Convert&Compare 'Filepath to directory' -Recursive");
            Console.WriteLine();
            Console.WriteLine("Input your argument:");
            // Object reference
            Spreadsheet process = new Spreadsheet();
            // Method reference
            foreach (string arg in args)
            {
                switch (args[0])
                {
                    case "Count":
                        process.Count(args[2]);
                        break;
                    case "Count&Convert":
                        process.Count(args[2]);
                        process.Convert(args[2]);
                        break;
                    case "Count&Convert&Compare":
                        process.Count(args[2]);
                        process.Convert(args[2]);
                        process.Compare(args[2]);
                        break;
                    default: throw new ArgumentException("Unknown argument", args[0]);
                }
            }

        }

    }

}