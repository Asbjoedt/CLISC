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
            
            Console.WriteLine("Compare");
            Console.WriteLine("---");


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
