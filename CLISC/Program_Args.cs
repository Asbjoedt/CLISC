using System;
using System.Collections.Generic;
using CommandLine;

namespace CLISC
{
    public class Program_Args
    {
        [Option('f', "function", Required = true, HelpText = "Specify the function you want to use: count, count&convert, count&convert&compare or count&convert&compare&archive")]
        public string function { get; set; }

        [Option('i', "inputdir", Required = true, HelpText = "Specify the input directory you want to use with function. Do not end string with \\ .")]
        public string inputdir { get; set; }

        [Option('o', "outputdir", Required = true, HelpText = "Specify the output directory you want to use with function. Do not end string with \\ .")]
        public string outputdir { get; set; }

        [Option('r', "recurse", Required = false, HelpText = "Specify true if input directory should include subdirectories.", Default = false)]
        public bool recurse { get; set; }
    }
}
