using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC.Classes
{
    internal class Test
    {
        //Declare public variables
        public string directory;

        // User input
        public void UserInput()
        {
            // Input directory
            Console.WriteLine("Input directory path:");
            directory = Console.ReadLine();
            Console.WriteLine();
            // Include subdirectories
            Console.WriteLine("Include subdirectories? Input 'true' or 'false'");
            string recursiveString = Console.ReadLine();
            bool recursive = recursiveString == "true";
            if (recursiveString == "true")
            {
                Console.WriteLine("Subdirectories will be included");
            }
            else if (recursiveString == "false")
            {
                Console.WriteLine("Subdirectories will be excluded");
            }
            else
            {
                Console.WriteLine("Input not valid");
                // Restart method or create another kind of loop?
            }
            //return (directory);
        }

        // User confirmation prompt
        public void Confirm()
        {
            Console.WriteLine("Continue to next process y/n");
            string continue_conversion = Console.ReadLine();
            if (continue_conversion == "y")
            {
                Console.WriteLine();
                Console.WriteLine("Funktion på vej");
            }
            else
            {
                Environment.Exit(0);
            }
        }
    }
}
