using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Timers;
using Excel = Microsoft.Office.Interop.Excel; 

namespace CLISC
{
    public partial class Conversion
    {
        // Convert legacy Excel files to .xlsx Transitional using Microsoft Office Interop Excel
        public bool Convert_Legacy_ExcelInterop(string input_filepath, string output_filepath)
        {
            // Start timeout counter
            System.Timers.Timer timeout = new System.Timers.Timer();
            timeout.Interval = 300000;
            timeout.Elapsed += Timeout_Elapsed;
            timeout.AutoReset = false;
            timeout.Enabled = true;

            bool convert_success = false;

            Excel.Application app = new Excel.Application(); // Create Excel object instance
            app.DisplayAlerts = false; // Don't display any Excel prompts
            Excel.Workbook wb = app.Workbooks.Open(input_filepath); // Create workbook instance

            wb.SaveAs(output_filepath, 51); // Save workbook as .xlsx Transitional
            wb.Close(); // Close workbook
            app.Quit(); // Quit Excel application
           
            // Stop timer
            timeout.Stop();
            timeout.Dispose();

            convert_success = true; // Mark as succesful
            return convert_success; // Report success

        }
        private void Timeout_Elapsed(object? sender, ElapsedEventArgs e)
        {
            new TimeoutException("Conversion of file has exceeded 5 min. Handle file manually");
        }
    }
}