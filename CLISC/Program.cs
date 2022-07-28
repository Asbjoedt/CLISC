using System;
using System.Collections.Generic;
using CommandLine;
using CommandLine.Text;

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
            var parser = new Parser(with => with.HelpWriter = null);
            var parse_args = parser.ParseArguments<Program_Args>(args);
                parse_args
                .WithParsed(Run)
                .WithNotParsed(errs => Help(parse_args, errs));
        }

        // Show help dialog to user
        static void Help<T>(ParserResult<T> result, IEnumerable<Error> errs)
        {
            var helpText = HelpText.AutoBuild(result, h =>
            {
                h.AdditionalNewLineAfterOption = false;
                h.Heading = "Program could not run.";
                h.Copyright = "";
                h.AutoHelp = false;
                h.AutoVersion = false;
                h.MaximumDisplayWidth = 90;
                h.AddPostOptionsLine("Input new argument:");
                return HelpText.DefaultParsingErrorsHandler(result, h);
            }, e => e);
            Console.WriteLine(helpText);
        }

        // Run program if args are parsed
        static void Run(Program_Args Arg)
        {
            // Inform user of their input values
            Console.WriteLine("YOUR INPUT");
            Console.WriteLine("---");
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