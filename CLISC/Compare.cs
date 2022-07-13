using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;

namespace CLISC
{

    public partial class Spreadsheet
    {
        
        // Compare spreadsheets
        public void Compare(string argument1, string argument2, string argument3)
        {
            Console.WriteLine("COMPARE");
            Console.WriteLine("---");

            if (true)
            {
                // Author: Kamil Niklasinski
                // This script is provided under GNU license -see license file for details.
                // Make sure you add to system path folder with SPREADSHEETCOMPARE.EXE
                // C:\Program Files(x86)\Microsoft Office\Office15\DCF\

                //excomp.bat Book1.xlsx Book2.xlsx
                //dir % 1 / B / S > temp.txt
                //dir % 2 / B / S >> temp.txt
                //SPREADSHEETCOMPARE temp.txt
            }
            else
            {
                Console.WriteLine("Error: The program Microsoft Spreadsheet Compare is necessary for compare function to run.");
            }

            //Calculate checksums


            // Log
            Console.WriteLine();
            //Console.WriteLine($"{} out of {numTOTAL} conversions have workbook differences");
            Console.WriteLine("Results saved to log in CSV file format");
            Console.WriteLine("Comparison finished");
            Console.WriteLine();
        }
   
    }

}
