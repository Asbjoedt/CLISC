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

        // Show help dialog to users, if errors in arguments
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
                h.AddPostOptionsLine("Don't let the errors weigh you down. Here's a cool song: https://youtu.be/4Is3D_VpX1Y");
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
            Console.WriteLine($"Function: {Arg.Function}");
            Console.WriteLine($"Inputdir: {Arg.Inputdir}");
            Console.WriteLine($"Outputdir: {Arg.Outputdir}");
            Console.WriteLine($"Recurse: {Arg.Recurse}");
            Console.WriteLine("---");

            // Execute the real program with arguments
            Program_Real.Execute(Arg.Function, Arg.Inputdir, Arg.Outputdir, Arg.Recurse);
        }
    }
}