﻿using System;
using System.Collections.Generic;
using CommandLine;

namespace CLISC
{
    public class Program_Args
    {
        // Parameter function
        [Option('f', "function", Required = true, HelpText = "Specify the function you want to use: count, count&convert, count&convert&compare or count&convert&compare&archive")]
        public string function { get; set; }

        // Parameter input directory
        [Option('i', "inputdir", Required = true, HelpText = "Specify the input directory you want to use with function. Do not end string with \\ .")]
        public string inputdir { get; set; }

        // Parameter output directory
        [Option('o', "outputdir", Required = true, HelpText = "Specify the output directory you want to use with function. Do not end string with \\ .")]
        public string outputdir { get; set; }

        // Parameter to include subdirectories in input directory
        [Option('r', "recurse", Required = false, HelpText = "Specify true if input directory should include subdirectories.", Default = false)]
        public bool recurse { get; set; }
    }
}