using System;
using System.Collections.Generic;
using CommandLine;

namespace CLISC
{
    public class Program_Args
    {
        [Option('f', "function", Required = true, HelpText = "Specify the function you wish to use.")]
        public string function { get; set; }

        [Option('i', "inputdir", Required = true, HelpText = "Specify the input directory you wish to use with function.")]
        public string inputdir { get; set; }

        [Option('o', "outputdir", Required = true, HelpText = "Specify the output directory you wish to use with function.")]
        public string outputdir { get; set; }

        [Option('r', "recurse", Required = false, HelpText = "Specify if input directory should include subdirectories.", Default = false)]
        public bool recurse { get; set; }
    }
}
