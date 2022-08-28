using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Office2013.ExcelAc;

namespace CLISC
{
    public partial class Archive_Requirements
    {
        // Remove data connections
        public void Remove_DataConnections(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                // Delete connection
                ConnectionsPart conn = spreadsheet.WorkbookPart.ConnectionsPart;
                spreadsheet.WorkbookPart.DeletePart(conn);

                // Delete querytable
                List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart part in worksheetparts)
                {
                    Console.WriteLine(part.QueryTableParts.Count());
                    if (part.QueryTableParts == null)
                    {
                        Console.WriteLine("du er dejlig");
                    }
                }
            }
        }

        // Remove RTD functions
        public void Remove_RTDFunctions(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart part in worksheetparts)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = part.Worksheet;
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
                                        CellValue cellvalue = cell.CellValue; // Save current cell value
                                        cell.CellFormula = null; // Remove RTD formula
                                        // If cellvalue does not have a real value
                                        if (cellvalue.Text == "#N/A")
                                        {
                                            cell.DataType = CellValues.String;
                                            cell.CellValue = new CellValue("Invalid data removed");
                                        }
                                        else
                                        {
                                            cell.CellValue = cellvalue; // Insert saved cell value
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                // Delete calculation chain
                CalculationChainPart calc = spreadsheet.WorkbookPart.CalculationChainPart;
                spreadsheet.WorkbookPart.DeletePart(calc);

                // Delete volatile dependencies
                VolatileDependenciesPart vol = spreadsheet.WorkbookPart.VolatileDependenciesPart;
                spreadsheet.WorkbookPart.DeletePart(vol);
            }
        }

        // Remove printer settings
        public void Remove_PrinterSettings(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                var list = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (var item in list)
                {
                    var printerList = item.SpreadsheetPrinterSettingsParts.ToList();
                    foreach (var part in printerList)
                    {
                        item.DeletePart(part);
                    }
                }
            }
        }

        // Remove external cell references
        public void Remove_CellReferences(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                List<ExternalWorkbookPart> externalworkbookparts = spreadsheet.WorkbookPart.ExternalWorkbookParts.ToList();
                foreach (ExternalWorkbookPart externalworkbookpart in externalworkbookparts)
                {
                    if (externalworkbookpart.RelationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" || externalworkbookpart.RelationshipType == "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLinkPath")
                    {
                        // Delete cell reference formula

                        // Delete cell references part
                        spreadsheet.WorkbookPart.DeletePart(externalworkbookpart);
                    }
                }
                // Delete calculation chain
                CalculationChainPart calc = spreadsheet.WorkbookPart.CalculationChainPart;
                spreadsheet.WorkbookPart.DeletePart(calc);
            }
        }

        // Remove external object references
        public void Remove_ExternalObjects(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {

            }
        }

        // Make first sheet active sheet
        public void Activate_FirstSheet(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                BookViews bookViews = spreadsheet.WorkbookPart.Workbook.GetFirstChild<BookViews>();
                bookViews.Remove(); // Remove bookview and thereby remove custom active tab
            }
        }

        // Remove absolute path to local directory
        public void Remove_AbsolutePath(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                if (spreadsheet.WorkbookPart.Workbook.AbsolutePath.Url != null)
                {
                    spreadsheet.WorkbookPart.Workbook.AbsolutePath.Url = "";
                }
            }
        }

        // Transform data using Excel Interop
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