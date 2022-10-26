using System;
using System.IO.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2013.ExcelAc;

namespace CLISC
{
    public partial class Archive_Requirements
    {
        // Change .xlsx according to archival requirements
        public void Change_XLSX_Requirements(List<Archive_Requirements> arcReq, string filepath)
        {
            foreach (var item in arcReq)
            {
                if (item.Metadata == true)
                {
                    //Remove_Metadata(filepath);
                    //Console.WriteLine("--> Change: File property information was removed and saved to sidecar file");
                }
                if (item.Conformance == true)
                {
                    Change_Conformance_ExcelInterop(filepath);
                    Console.WriteLine("--> Change: Conformance was changed to Strict");
                }
                if (item.Connections > 0)
                {
                    Remove_DataConnections(filepath);
                    Console.WriteLine($"--> Change: {item.Connections} data connections were removed");
                }
                if (item.CellReferences > 0)
                {
                    Remove_CellReferences(filepath);
                    Console.WriteLine($"--> Change: {item.CellReferences} cell references were removed");
                }
                if (item.RTDFunctions > 0)
                {
                    Remove_RTDFunctions(filepath);
                    Console.WriteLine($"--> Change: {item.RTDFunctions} RTD functions were removed");
                }
                if (item.PrinterSettings > 0)
                {
                    Remove_PrinterSettings(filepath);
                    Console.WriteLine($"--> Change: {item.PrinterSettings} printer settings were removed");
                }
                if (item.ExternalObj > 0)
                {
                    Remove_ExternalObjects(filepath);
                    Console.WriteLine($"--> Change: {item.ExternalObj} external objects were removed");
                }
                if (item.ActiveSheet == true)
                {
                    Activate_FirstSheet(filepath);
                    Console.WriteLine("--> Change: First sheet was activated");
                }
                if (item.AbsolutePath == true)
                {
                    Remove_AbsolutePath(filepath);
                    Console.WriteLine("--> Change: Absolute path to local directory was removed");
                }
                if (item.EmbedObj > 0)
                {
                    //Remove_EmbeddedObjects(filepath);
                    //Console.WriteLine($"--> Change: {item.EmbedObj} embedded objects were removed");
                }
                if (item.Hyperlinks > 0)
                {
                    //Change_Hyperlinks(filepath);
                    //Console.WriteLine($"--> Change: {item.Hyperlinks} hyperlinks were converted to Wayback Machine hyperlinks");
                }
            }
        }

        // Change conformance to Strict
        public void Change_Conformance(string filepath)
        {
            // Work in progress

            // Create list of namespaces
            List<namespaceIndex> namespaces = namespaceIndex.Create_Namespaces_Index();

            using (var spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Workbook workbook = wbPart.Workbook;
                // If Transitional
                if (workbook.Conformance == null || workbook.Conformance != "strict")
                {
                    // Change conformance class
                    workbook.Conformance.Value = ConformanceClass.Enumstrict;

                    // Add vml urn namespace to workbook.xml
                    workbook.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
                }
            }
        }

        // Remove data connections
        public void Remove_DataConnections(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                // Delete all connections
                ConnectionsPart conn = spreadsheet.WorkbookPart.ConnectionsPart;
                spreadsheet.WorkbookPart.DeletePart(conn);

                // Delete all query tables
                List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart part in worksheetparts)
                {
                    List<QueryTablePart> queryTables = part.QueryTableParts.ToList();
                    foreach (QueryTablePart qtp in queryTables)
                    {
                        part.DeletePart(qtp);
                    }
                }

                // If spreadsheet contains a custom XML Map, delete databinding
                if (spreadsheet.WorkbookPart.CustomXmlMappingsPart != null)
                {
                    CustomXmlMappingsPart xmlMap = spreadsheet.WorkbookPart.CustomXmlMappingsPart;
                    List<Map> maps = xmlMap.MapInfo.Elements<Map>().ToList();
                    foreach (Map map in maps)
                    {
                        if (map.DataBinding != null)
                        {
                            map.DataBinding.Remove();
                        }
                    }
                }
            }
            // Repair spreadsheet
            Repair rep = new Repair();
            //rep.Repair_QueryTables(filepath);
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
                List<WorksheetPart> wsParts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart wsPart in wsParts)
                {
                    List<SpreadsheetPrinterSettingsPart> printerList = wsPart.SpreadsheetPrinterSettingsParts.ToList();
                    foreach (SpreadsheetPrinterSettingsPart printer in printerList)
                    {
                        wsPart.DeletePart(printer);
                    }
                }
            }
        }

        // Remove external cell references
        public void Remove_CellReferences(string filepath)
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
                                if (formula.Length > 1)
                                {
                                    string hit = formula.Substring(0, 1); // Transfer first 1 characters to string
                                    string hit2 = formula.Substring(0, 2); // Transfer first 2 characters to string
                                    if (hit == "[" || hit2 == "'[")
                                    {
                                        CellValue cellvalue = cell.CellValue; // Save current cell value
                                        cell.CellFormula = null;
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

                // Delete external book references
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

                // Delete defined names that includes external cell references
                DefinedNames definedNames = spreadsheet.WorkbookPart.Workbook.DefinedNames;
                if (definedNames != null)
                {
                    var definedNamesList = definedNames.ToList();
                    foreach (DefinedName definedName in definedNamesList)
                    {
                        if (definedName.InnerXml.StartsWith("["))
                        {
                            definedName.Remove();
                        }
                    }
                }
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
                List<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
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
                BookViews bookViews = spreadsheet.WorkbookPart.Workbook.GetFirstChild<BookViews>();
                WorkbookView workbookView = bookViews.GetFirstChild<WorkbookView>();
                if (workbookView.ActiveTab != null)
                {
                    var activeSheetId = workbookView.ActiveTab.Value;
                    if (activeSheetId > 0)
                    {
                        workbookView.ActiveTab.Value = 0;

                        List<WorksheetPart> worksheets = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                        foreach (WorksheetPart worksheet in worksheets)
                        {
                            var sheetviews = worksheet.Worksheet.SheetViews.ToList();
                            foreach (SheetView sheetview in sheetviews)
                            {
                                sheetview.TabSelected = null;
                            }
                        }
                    }
                }
            }
        }

        // Remove absolute path to local directory
        public void Remove_AbsolutePath(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                if (spreadsheet.WorkbookPart.Workbook.AbsolutePath != null)
                {
                    AbsolutePath absPath = spreadsheet.WorkbookPart.Workbook.GetFirstChild<AbsolutePath>();
                    absPath.Remove();
                }
            }
        }

        // Remove metadata in file properties
        public void Remove_Metadata(string filepath)
        {
            string folder = Path.GetDirectoryName(filepath);

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                PackageProperties property = spreadsheet.Package.PackageProperties;

                // Create metadata file
                using (StreamWriter w = File.AppendText($"{folder}\\orgFile_Metadata.txt"))
                {
                    w.WriteLine("STRIPPED FILE PROPERTIES INFORMATION");
                    w.WriteLine("---");
                }

                if (property.Creator != null)
                {
                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"CREATOR: {property.Creator}");
                    }

                    // Remove information
                    property.Creator = null;
                }
                if (property.Title != null)
                {
                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"TITLE: {property.Title}");
                    }

                    // Remove information
                    property.Title = null;
                }
                if (property.Subject != null)
                {
                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"SUBJECT: {property.Subject}");
                    }

                    // Remove information
                    property.Subject = null;
                }
                if (property.Description != null)
                {
                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"DESCRIPTION: {property.Description}");
                    }

                    // Remove information
                    property.Description = null;
                }
                if (property.Keywords != null)
                {

                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"KEYWORDS: {property.Keywords}");
                    }

                    // Remove information
                    property.Keywords = null;
                }
                if (property.Category != null)
                {
                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"CATEGORY: {property.Category}");
                    }

                    // Remove information
                    property.Category = null;
                }
                if (property.LastModifiedBy != null)
                {
                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"LAST MODIFIED BY: {property.LastModifiedBy}");
                    }

                    // Remove information
                    property.LastModifiedBy = null;
                }
            }
        }

        // Change hyperlinks to link to Wayback Machine
        public void Change_Hyperlinks(string filepath)
        {
            string old_hyperlink = "";
            string new_hyperlink = "";

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                List<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart worksheetPart in worksheetParts)
                {
                    Worksheet worksheet = worksheetPart.Worksheet;
                    IEnumerable<Hyperlink> hyperlinks = worksheet.GetFirstChild<Hyperlinks>().Elements<Hyperlink>();
                    foreach(Hyperlink hyperlink in hyperlinks)
                    {
                        Console.WriteLine(hyperlink.Id);
                        ReferenceRelationship refRel = worksheetPart.GetReferenceRelationship(hyperlink.Id);
                    }
                }
            }
        }
    }
}