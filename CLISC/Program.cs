using System;
using System.IO;
using System.Collections.Generic;
using CommandLine;

namespace CLISC
{
    class Program
    {
        static void Main(string[] args)
        {
            // Inform user of beginning of program
            Console.WriteLine("CLISC - Command Line Interface Spreadsheet Count Convert & Compare");
            Console.WriteLine("@Asbjørn Skødt, web: https://github.com/Asbjoedt/CLISC");
            Console.WriteLine("---");

            // Parse user arguments
            var parse_args = Parser.Default.ParseArguments<Program_Args>(args)
                .WithParsed(Run)
                .WithNotParsed(HandleParseError);
        }

        // Handle errors if args are not parsed
        static void HandleParseError(IEnumerable<Error> errs)
        {
            Console.ReadLine();
        }

        // Run program if args are parsed
        static void Run(Program_Args Arg)
        {
            Console.WriteLine($"Function: {Arg.function}");
            Console.WriteLine($"Inputdir: {Arg.inputdir}");
            Console.WriteLine($"Outputdir: {Arg.outputdir}");
            Console.WriteLine($"Recurse: {Arg.recurse}");
            Console.WriteLine("---");

            // Execute the real program with arguments
            Program_Real.Execute(Arg.function, Arg.inputdir, Arg.outputdir, Arg.recurse);
        }
    }
}