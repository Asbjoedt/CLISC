using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Office2013.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace CLISC
{
    public partial class Archive_Requirements
    {
        // Remove data connections
        public void Remove_DataConnections(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                // Delete all connections
                ConnectionsPart conn = spreadsheet.WorkbookPart.ConnectionsPart;
                spreadsheet.WorkbookPart.DeletePart(conn);

                List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart part in worksheetparts)
                {
                    int qtables = part.QueryTableParts.Count();
                    Console.WriteLine(qtables);

                    List<QueryTablePart> queryTables = part.QueryTableParts.ToList();
                    foreach (QueryTablePart qtp in queryTables)
                    {
                        spreadsheet.WorkbookPart.DeletePart(qtp);
                    }

                    qtables = part.QueryTableParts.Count();
                    Console.WriteLine(qtables);
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
                    Worksheet worksheet = part.Worksheet;
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
                // Delete all cell references in worksheet
                List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart part in worksheetparts)
                {
                    Worksheet worksheet = part.Worksheet;
                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                    foreach (var row in rows)
                    {
                        var cells = row.Elements<Cell>();
                        foreach (Cell cell in cells)
                        {
                            if (cell.CellFormula != null)
                            {
                                string formula = cell.CellFormula.InnerText;
                                if (formula.Length > 0)
                                {
                                    string hit = formula.Substring(0, 1); // Transfer first 1 characters to string
                                    if (hit == "[")
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
                // Delete all external link references
                List<ExternalWorkbookPart> extwbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts.ToList();
                if (extwbParts.Count > 0)
                {
                    foreach (ExternalWorkbookPart extpart in extwbParts)
                    {
                        var elements = extpart.ExternalLink.ChildElements.ToList();
                        foreach (var element in elements)
                        {
                            if (element.LocalName == "externalBook")
                            {
                                spreadsheet.WorkbookPart.DeletePart(extpart);
                            }
                        }
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
                List<ExternalWorkbookPart> extwbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts.ToList();
                if (extwbParts.Count > 0)
                {
                    foreach (ExternalWorkbookPart extpart in extwbParts)
                    {
                        if (extpart.ExternalLink.ChildElements != null)
                        {
                            var elements = extpart.ExternalLink.ChildElements.ToList();
                            foreach (var element in elements)
                            {
                                if (element.LocalName == "oleLink")
                                {
                                    spreadsheet.WorkbookPart.DeletePart(extpart);
                                }
                            }
                        }
                    }
                }
            }
        }

        public void Remove_EmbeddedObjects(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                var worksheetParts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart worksheetPart in worksheetParts)
                {
                    List<EmbeddedObjectPart> embedobj_ole_list = worksheetPart.EmbeddedObjectParts.ToList();
                    List<EmbeddedPackagePart> embedobj_package_list = worksheetPart.EmbeddedPackageParts.ToList();
                    List<ImagePart> embedobj_image_list = worksheetPart.ImageParts.ToList();
                    List<ImagePart> embedobj_drawing_image_list = new List<ImagePart>();
                    if (worksheetPart.DrawingsPart != null)
                    {
                        embedobj_drawing_image_list = worksheetPart.DrawingsPart.ImageParts.ToList();
                    }
                    List<Model3DReferenceRelationshipPart> embedobj_3d_list = worksheetPart.Model3DReferenceRelationshipParts.ToList();

                    if (embedobj_ole_list.Count() > 0)
                    {
                        foreach (EmbeddedObjectPart ole in embedobj_ole_list)
                        {
                            worksheetPart.DeletePart(ole);
                        }
                    }

                    if (embedobj_package_list.Count() > 0)
                    {
                        foreach (EmbeddedPackagePart package in embedobj_package_list)
                        {
                            worksheetPart.DeletePart(package);
                        }
                    }
                    if (embedobj_image_list.Count() > 0)
                    {
                        foreach (ImagePart image in embedobj_image_list)
                        {
                            worksheetPart.DeletePart(image);
                        }
                    }
                    if (embedobj_drawing_image_list.Count() > 0)
                    {
                        foreach (ImagePart drawing_image in embedobj_drawing_image_list)
                        {
                            worksheetPart.DrawingsPart.DeletePart(drawing_image);
                        }
                    }
                    if (embedobj_3d_list.Count() > 0)
                    {
                        foreach (Model3DReferenceRelationshipPart threeD in embedobj_3d_list)
                        {
                            worksheetPart.DeletePart(threeD);
                        }
                    }
                }
            }
        }

        // Make first sheet active sheet
        public void Activate_FirstSheet(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                if (spreadsheet.WorkbookPart.Workbook.BookViews != null)
                {
                    BookViews bookViews = spreadsheet.WorkbookPart.Workbook.GetFirstChild<BookViews>();
                    bookViews.Remove(); // Remove bookview and thereby remove custom active tab
                }
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
    }
}