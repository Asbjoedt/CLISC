using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

namespace CLISC
{
    public partial class Archive
    {
        public void Transform_XLSX_Requirements(string filepath)
        {
            int connections = Check_DataConnections(filepath);
            if (connections > 0)
            {
                Remove_DataConnections(filepath);
            }
            int extrels = Check_ExternalRelationships(filepath);
            if (extrels > 0)
            {
                Handle_ExternalRelationships(filepath);
            }
            int rtdfunctions = Check_RTDFunctions(filepath);
            if (rtdfunctions > 0)
            {
                Remove_RTDFunctions(filepath);
            }
            int printersettings = Check_PrinterSettings(filepath);
            if (printersettings > 0)
            {
                Remove_PrinterSettings(filepath);
            }
        }

        // Remove data connections
        public void Remove_DataConnections(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                ConnectionsPart conn = spreadsheet.WorkbookPart.ConnectionsPart;
                if (conn != null)
                {
                    //conn.Connections.RemoveAllChildren();
                    bool success = conn.DeletePart(conn);
                }
            }
        }

        // Remove RTD functions
        public int Remove_RTDFunctions(string filepath) // Remove RTD functions
        {
            int rtd_functions = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Sheets allSheets = wbPart.Workbook.Sheets;
                foreach (Sheet aSheet in allSheets)
                {
                    WorksheetPart wsp = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(aSheet.Id);
                    DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = wsp.Worksheet;
                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                    foreach (var row in rows)
                    {
                        var cells = row.Elements<Cell>();
                        foreach (Cell cell in cells)
                        {
                            if (cell.CellFormula != null)
                            {
                                string formula = cell.CellFormula.InnerText;
                                if (formula.Length > 2)
                                {
                                    string hit = formula.Substring(0, 3); // Transfer first 3 characters to string
                                    if (hit == "RTD")
                                    {
                                        var cellvalue = cell.CellValue;
                                        cell.CellFormula = null;
                                        cell.CellValue = cellvalue;
                                        Console.WriteLine($"--> RTD function in sheet \"{aSheet.Name}\" cell {cell.CellReference} removed");
                                        rtd_functions++;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return rtd_functions;
        }

        // Remove printersettings
        public void Remove_PrinterSettings(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                var list = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (var item in list)
                {
                    var printerList = item.SpreadsheetPrinterSettingsParts.ToList();
                    foreach (var part in printerList)
                    {

                    }
                }
            }
        }


        public void Transform_Requirements_ExcelInterop(string filepath)  // Use Excel Interop
        {
            Excel.Application app = new Excel.Application(); // Create Excel object instance
            app.DisplayAlerts = false; // Don't display any Excel window prompts
            Excel.Workbook wb = app.Workbooks.Open(filepath); // Create workbook instance

            // Find and delete data connections
            int count_conn = wb.Connections.Count;
            if (count_conn > 0)
            {
                for (int i = 1; i <= wb.Connections.Count; i++)
                {
                    wb.Connections[i].Delete();
                    i = i - 1;
                }
                count_conn = wb.Connections.Count;
                wb.Save(); // Save workbook
            }

            // Find and replace RTD functions with cell values
            bool hasRTD = false;
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                try
                {
                    Excel.Range range = (Excel.Range)sheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                    //Excel.Range range = sheet.get_Range("A1", "XFD1048576").SpecialCells(Excel.XlCellType.xlCellTypeFormulas); // Alternative range

                    foreach (Excel.Range cell in range.Cells)
                    {
                        var value = cell.Value2;
                        string formula = cell.Formula.ToString();
                        string hit = formula.Substring(0, 4); // Transfer first 4 characters to string

                        if (hit == "=RTD")
                        {
                            hasRTD = true;
                            cell.Formula = "";
                            cell.Value2 = value;
                        }
                    }
                    if (hasRTD == true)
                    {
                        Console.WriteLine("--> RTD function formulas detected and replaced with cell values"); // Inform user
                        wb.Save(); // Save workbook
                    }
                }
                catch (System.Runtime.InteropServices.COMException) // Catch if no formulas in range
                {
                    // Do nothing
                }
                catch (System.ArgumentOutOfRangeException) // Catch if formula has less than 4 characters
                {
                    // Do nothing
                }
            }

            // Find and replace external cell chains with cell values
            bool hasChain = false;
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                try
                {
                    Excel.Range range = (Excel.Range)sheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                    //Excel.Range range = sheet.get_Range("A1", "XFD1048576").SpecialCells(Excel.XlCellType.xlCellTypeFormulas); // Alternative range

                    foreach (Excel.Range cell in range.Cells)
                    {
                        var value = cell.Value2;
                        string formula = cell.Formula.ToString();
                        string hit = formula.Substring(0, 2); // Transfer first 2 characters to string

                        if (hit == "='")
                        {
                            hasChain = true;
                            cell.Formula = "";
                            cell.Value2 = value;
                        }
                    }
                    if (hasChain == true)
                    {
                        Console.WriteLine("--> External cell chains detected and replaced with cell values"); // Inform user
                        wb.Save(); // Save workbook
                    }
                }
                catch (System.Runtime.InteropServices.COMException) // Catch if no formulas in range
                {
                    // Do nothing
                }
                catch (System.ArgumentOutOfRangeException) // Catch if formula has less than 2 characters
                {
                    // Do nothing
                }
            }

            try
            {
                // Make first cell in first sheet active
                if (app.Sheets.Count > 0)
                {
                    Excel.Worksheet firstSheet = (Excel.Worksheet)app.ActiveWorkbook.Sheets[1];
                    firstSheet.Activate();
                    firstSheet.Select();
                }
            }
            // For some reason an exception is thrown in some spreadsheets when trying to make the first sheet active
            catch (System.Runtime.InteropServices.COMException)
            {
                // Do nothing
            }

            wb.Save(); // Save workbook
            wb.Close(); // Close the workbook
            app.Quit(); // Quit Excel application
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // If app is run on Windows
            {
                Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                Marshal.ReleaseComObject(app); // Delete Excel task in task manager
            }
        }
    }
}
