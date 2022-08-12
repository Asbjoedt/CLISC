using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public class FileFormats
    {
        // Public arrays
        public static string[] Extension = { ".gsheet", ".fods", ".numbers", ".ods", ".ots", ".xla", ".xlam", ".xls", ".xlsb", ".xlsm", ".xlsx", ".xlt", ".xltm", ".xltx" };

        public static string[] Extension_Upper = { ".GSHEET", ".FODS", ".NUMBERS", ".ODS", ".OTS", ".XLA", ".XLAM", ".XLS", ".XLSB", ".XLSM", ".XLSX", ".XLT", ".XLTM", ".XLTX" };

        public static string[] Description = { "Google Sheets hyperlink", "OpenDocument Flat XML Spreadsheet", "Apple Numbers Spreadsheet", "OpenDocument Spreadsheet", "OpenDocument Spreadsheet Template", "Legacy Microsoft Excel Spreadsheet Add-In", "Office Open XML Macro-Enabled Add-In", "Legacy Microsoft Excel Spreadsheet", "Office Open XML Binary Spreadsheet", "Office Open XML Macro-Enabled Spreadsheet", "Office Open XML Spreadsheet (Transitional and Strict conformance)", "Legacy Microsoft Excel Spreadsheet Template", "Office Open XML Macro-Enabled Spreadsheet Template", "Office Open XML Spreadsheet Template" };
    }
}
