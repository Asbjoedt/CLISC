using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    // Index sorted alphabetically by extension
    public class fileFormatIndex
    {
        // Public arrays
        public static string[] Extension_Array = { ".gsheet", ".fods", ".numbers", ".ods", ".ots", ".xla", ".xlam", ".xls", ".xlsb", ".xlsm", ".xlsx", ".xlt", ".xltm", ".xltx" };

        public static string[] Extension_Upper_Array = { ".GSHEET", ".FODS", ".NUMBERS", ".ODS", ".OTS", ".XLA", ".XLAM", ".XLS", ".XLSB", ".XLSM", ".XLSX", ".XLT", ".XLTM", ".XLTX" };

        public static string[] Description_Array = { "Google Sheets hyperlink", "OpenDocument Flat XML Spreadsheet", "Apple Numbers Spreadsheet", "OpenDocument Spreadsheet", "OpenDocument Spreadsheet Template", "Legacy Microsoft Excel Spreadsheet Add-In", "Office Open XML Macro-Enabled Add-In", "Legacy Microsoft Excel Spreadsheet", "Office Open XML Binary Spreadsheet", "Office Open XML Macro-Enabled Spreadsheet", "Office Open XML Spreadsheet (Transitional and Strict conformance)", "Legacy Microsoft Excel Spreadsheet Template", "Office Open XML Macro-Enabled Spreadsheet Template", "Office Open XML Spreadsheet Template" };

        public string Extension { get; protected set; }

        public string Extension_Upper { get; protected set; }

        public string Description { get; protected set; }

        public int? Count { get; set; }

        public string? Conformance { get; protected set; }

        public List<fileFormatIndex> Create_fileFormatIndex()
        {
            List<fileFormatIndex> list = new List<fileFormatIndex>();

            // GSHEET
            list.Add(new fileFormatIndex() { Extension = ".gsheet", Extension_Upper = ".GSHEET", Description = "Google Sheets hyperlink", });
            // FODS
            list.Add(new fileFormatIndex() { Extension = ".fods", Extension_Upper = ".FODS", Description = "OpenDocument Flat XML Spreadsheet" });
            // NUMBERS
            list.Add(new fileFormatIndex() { Extension = ".numbers", Extension_Upper = ".NUMBERS", Description = "Apple Numbers Spreadsheet" });
            // ODS
            list.Add(new fileFormatIndex() { Extension = ".ods", Extension_Upper = ".ODS", Description = "OpenDocument Spreadsheet" });
            // OTS
            list.Add(new fileFormatIndex() { Extension = ".ots", Extension_Upper = ".OTS", Description = "OpenDocument Spreadsheet Template" });
            // XLA
            list.Add(new fileFormatIndex() { Extension = ".xla", Extension_Upper = ".XLA", Description = "Legacy Microsoft Excel Spreadsheet Add-In" });
            // XLAM
            list.Add(new fileFormatIndex() { Extension = ".xlam", Extension_Upper = ".XLAM", Description = "Office Open XML Macro-Enabled Add-In" });
            // XLS
            list.Add(new fileFormatIndex() { Extension = ".xls", Extension_Upper = ".XLS", Description = "Legacy Microsoft Excel Spreadsheet" });
            // XLSB
            list.Add(new fileFormatIndex() { Extension = ".xlsb", Extension_Upper = ".XLSB", Description = "Office Open XML Binary Spreadsheet" });
            // XLSM
            list.Add(new fileFormatIndex() { Extension = ".xlsm", Extension_Upper = ".XLSM", Description = "Office Open XML Macro-Enabled Spreadsheet" });
            // XLSX - Transitional and Strict conformance
            list.Add(new fileFormatIndex() { Extension = ".xlsx", Extension_Upper = ".XLSX", Description = "Office Open XML Spreadsheet (transitional and strict conformance)" });
            // XLSX - Transitional conformance
            list.Add(new fileFormatIndex() { Extension = ".xlsx", Extension_Upper = ".XLSX", Description = "Office Open XML Spreadsheet (transitional conformance)", Conformance = "transitional" });
            // XLSX - Strict conformance
            list.Add(new fileFormatIndex() { Extension = ".xlsx", Extension_Upper = ".XLSX", Description = "Office Open XML Spreadsheet (strict conformance)", Conformance = "strict" });
            // XLT
            list.Add(new fileFormatIndex() { Extension = ".xlt", Extension_Upper = ".XLT", Description = "Legacy Microsoft Excel Spreadsheet Template" });
            // XLTM
            list.Add(new fileFormatIndex() { Extension = ".xltm", Extension_Upper = ".XLTM", Description = "Office Open XML Macro-Enabled Spreadsheet Template" });
            // XLTX
            list.Add(new fileFormatIndex() { Extension = ".XLTX", Extension_Upper = ".XLTX", Description = "Office Open XML Spreadsheet Template" });

            return list;
        }
    }
}
