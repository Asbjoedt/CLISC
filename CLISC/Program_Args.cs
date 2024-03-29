﻿using CommandLine;

namespace CLISC
{
    public class Program_Args
    {
        // Parameter function
        [Option('f', "function", Required = true, HelpText = "Specify the function you want to use: Count, CountConvert, CountConvertCompare or CountConvertCompareArchive")]
        public string Function { get; set; }

        // Parameter input directory
        [Option('i', "inputdir", Required = true, HelpText = "Specify the input directory you want to use with function. Do not end string with \\ .")]
        public string Inputdir { get; set; }

        // Parameter output directory
        [Option('o', "outputdir", Required = true, HelpText = "Specify the output directory you want to use with function. Do not end string with \\ .")]
        public string Outputdir { get; set; }

        // Parameter to include subdirectories in input directory
        [Option('r', "recurse", Required = false, HelpText = "Specify if input directory should include subdirectories.", Default = false)]
        public bool Recurse { get; set; }

        // Parameter to set fullcompliance for archiving
        [Option('c', "fullcompliance", Required = false, HelpText = "Specify if archiving should create full compliance with OPF specification.", Default = false)]
        public bool FullCompliance { get; set; }
    }
}
