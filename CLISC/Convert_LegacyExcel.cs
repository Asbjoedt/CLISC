﻿using System.IO;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel; 

namespace CLISC
{
    public partial class Conversion
    {
        // Convert legacy Excel files to .xlsx Transitional using Microsoft Office Interop Excel
        public bool Convert_Legacy_ExcelInterop(string input_filepath, string output_filepath)
        {
          bool convert_success = false;

            Excel.Application app = new Excel.Application(); // Create Excel object instance
            app.DisplayAlerts = false; // Don't display any Excel prompts
            Excel.Workbook wb = app.Workbooks.Open(input_filepath, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

            wb.SaveAs(output_filepath, 51); // Save workbook as .xlsx Transitional
            wb.Close(); // Close workbook
            app.Quit(); // Quit Excel application

            convert_success = true; // Mark as succesful
            return convert_success; // Report success
        }
    }
}