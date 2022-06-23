using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC.Classes
{
    public class Count
    {
        public static void CountSpreadsheets()
        {
            DirectoryInfo di = new DirectoryInfo(@directory);
            int numXLS = di.GetFiles("*.xls", SearchOption.AllDirectories).Length;
            int numXLT = di.GetFiles("*.xlt", SearchOption.AllDirectories).Length;
            int numXLAM = di.GetFiles("*.xlam", SearchOption.AllDirectories).Length;
            int numXLSB = di.GetFiles("*.xlsb", SearchOption.AllDirectories).Length;
            int numXLSM = di.GetFiles("*.xlsm", SearchOption.AllDirectories).Length;
            int numXLSX = di.GetFiles("*.xlsx", SearchOption.AllDirectories).Length;
            int numXLTM = di.GetFiles("*.xltm", SearchOption.AllDirectories).Length;
            int numXLTX = di.GetFiles("*.xltx", SearchOption.AllDirectories).Length;
            // Show count to user
            Console.WriteLine();
            Console.WriteLine($"{numXLS} XLS");
            Console.WriteLine($"{numXLT} XLT");
            Console.WriteLine($"{numXLAM} XLAM");
            Console.WriteLine($"{numXLSB} XLSB");
            Console.WriteLine($"{numXLSM} XLSM");
            Console.WriteLine($"{numXLSX} XLSX");
            Console.WriteLine($"{numXLTM} XLTM");
            Console.WriteLine($"{numXLTX} XLTX");
            Console.WriteLine();
            if (numXLS == 0 && numXLT == 0 && numXLAM == 0 && numXLSB == 0 && numXLSM == 0 && numXLSX == 0 && numXLTM == 0 && numXLTX == 0)
            {
                Console.WriteLine("No spreadsheets in the directory. You can close the application");
                return;
            }
        }
    }
}
