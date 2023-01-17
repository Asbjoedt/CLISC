using System;
using System.IO;
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
                    //int success = Remove_Metadata(filepath);
                    //Console.WriteLine($"--> Change: {success} file property information were removed and saved to sidecar file");
                }
                if (item.Conformance == true)
                {
                    Change_Conformance_ExcelInterop(filepath);
                    Console.WriteLine("--> Change: Conformance was changed to Strict");
                }
                if (item.Connections > 0)
                {
                    int success = Remove_DataConnections(filepath);
                    Console.WriteLine($"--> Change: {success} data connections were removed");
                }
                if (item.CellReferences > 0)
                {
                    int success = Remove_ExternalCellReferences(filepath);
                    Console.WriteLine($"--> Change: {success} cell references were removed");
                }
                if (item.RTDFunctions > 0)
                {
                    int success = Remove_RTDFunctions(filepath);
                    Console.WriteLine($"--> Change: {success} RTD functions were removed");
                }
                if (item.PrinterSettings > 0)
                {
                    int success = Remove_PrinterSettings(filepath);
                    Console.WriteLine($"--> Change: {success} printer settings were removed");
                }
                if (item.ExternalObj > 0)
                {
                    int success = Remove_ExternalObjects(filepath);
                    Console.WriteLine($"--> Change: {success} external objects were removed");
                }
                if (item.EmbedObj > 0)
                {
                    int success = Convert_EmbeddedObjects(filepath);
                    Console.WriteLine($"--> Change: {success} embedded objects were converted");
                }
                if (item.ActiveSheet == true)
                {
                    bool success = Activate_FirstSheet(filepath);
                    if (success)
                    {
                        Console.WriteLine("--> Change: First sheet was activated");
                    }
                }
                if (item.AbsolutePath == true)
                {
                    bool success = Remove_AbsolutePath(filepath);
                    if (success)
                    {
                        Console.WriteLine("--> Change: Absolute path to local directory was removed");
                    }
                }
                if (item.Hyperlinks > 0)
                {
                    //Change_Hyperlinks(filepath);
                    //Console.WriteLine($"--> Change: {item.Hyperlinks} hyperlinks were converted to Wayback Machine hyperlinks");
                }
            }
        }

        // Work in progress
        // Change conformance to Strict
        public void Change_Conformance(string filepath)
        {
            // Create list of namespaces
            List<namespaceIndex> namespaces = namespaceIndex.Create_Namespaces_Index();

            using (var spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                Workbook workbook = wbPart.Workbook;
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
        public int Remove_DataConnections(string filepath)
        {
            int success = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                ConnectionsPart conn = spreadsheet.WorkbookPart.ConnectionsPart;

                // Count connections
                success = conn.Connections.Count();

                // Delete all connections
                spreadsheet.WorkbookPart.DeletePart(conn);

                // Delete all query tables
                IEnumerable<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart part in worksheetparts)
                {
                    List<QueryTablePart> queryTables = part.QueryTableParts.ToList(); // Must be a list
                    foreach (QueryTablePart qtp in queryTables)
                    {
                        part.DeletePart(qtp);
                    }
                }

                // If spreadsheet contains a custom XML Map, delete databinding
                if (spreadsheet.WorkbookPart.CustomXmlMappingsPart != null)
                {
                    CustomXmlMappingsPart xmlMap = spreadsheet.WorkbookPart.CustomXmlMappingsPart;
                    List<Map> maps = xmlMap.MapInfo.Elements<Map>().ToList(); // Must be a list
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
            //Repair rep = new Repair();
            //rep.Repair_QueryTables(filepath);

            return success;
        }

        // Remove RTD functions
        public int Remove_RTDFunctions(string filepath)
        {
            int success = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart part in worksheetparts)
                {
                    Worksheet worksheet = part.Worksheet;
                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                    foreach (var row in rows)
                    {
                        IEnumerable<Cell> cells = row.Elements<Cell>();
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
                                        // Add to success
                                        success++;
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
            return success;
        }

        // Remove printer settings
        public int Remove_PrinterSettings(string filepath)
        {
            int success = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<WorksheetPart> wsParts = spreadsheet.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart wsPart in wsParts)
                {
                    List<SpreadsheetPrinterSettingsPart> printers = wsPart.SpreadsheetPrinterSettingsParts.ToList(); // Must be a list
                    foreach (SpreadsheetPrinterSettingsPart printer in printers)
                    {
                        wsPart.DeletePart(printer);
                        success++;
                    }
                }
            }
            return success;
        }

        // Remove external cell references
        public int Remove_ExternalCellReferences(string filepath)
        {
            int success = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart part in worksheetparts)
                {
                    Worksheet worksheet = part.Worksheet;
                    IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                    foreach (Row row in rows)
                    {
                        IEnumerable<Cell> cells = row.Elements<Cell>();
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
                                        // Add to success
                                        success++;
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
            return success;
        }

        // Remove external object references
        public int Remove_ExternalObjects(string filepath)
        {
            int success = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<ExternalWorkbookPart> extWbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts;
                foreach (ExternalWorkbookPart extWbPart in extWbParts)
                {
                    List<ExternalRelationship> extrels = extWbPart.ExternalRelationships.ToList(); // Must be a list
                    foreach (ExternalRelationship extrel in extrels)
                    {
                        // Change external target reference
                        Uri uri = new Uri("External reference was removed", UriKind.Relative);
                        extWbPart.DeleteExternalRelationship("rId1");
                        extWbPart.AddExternalRelationship(relationshipType: "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject", externalUri: uri, id: "rId1");

                        // Add to success
                        success++;
                    }
                }
            }
            return success;
        }

        // Embed external objects
        public int Embed_ExternalObjects(string filepath)
        {
            int success = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<ExternalWorkbookPart> extWbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts;
                foreach (ExternalWorkbookPart extWbPart in extWbParts)
                {
                    List<ExternalRelationship> extrels = extWbPart.ExternalRelationships.ToList(); // Must be a list
                    foreach (ExternalRelationship extrel in extrels)
                    {
                        // Change external target reference
                        Uri uri = new Uri("External reference was removed", UriKind.Relative);
                        extWbPart.DeleteExternalRelationship("rId1");
                        extWbPart.AddExternalRelationship(relationshipType: "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject", externalUri: uri, id: "rId1");

                        // Add to success
                        success++;
                    }



                    // Do a null check to see if the external object is available
                    if (extWbPart != null)
                    {
                        // Embed object


                        // Delete external relationship
                        extWbPart.DeleteExternalRelationship("rId1");

                        // Different approach to deleting external relationship
                        ExternalRelationship extrel = extWbPart.ExternalRelationships.FirstOrDefault();
                        extWbPart.DeleteExternalRelationship(extrel.Id);

                        // Add to success
                        success++;
                    }
                }
            }
            return success;
        }

        public int Remove_EmbeddedObjects(string filepath)
        {
            int success = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
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
                            success++;
                        }
                    }
                    if (embedobj_package_list.Count() > 0)
                    {
                        foreach (EmbeddedPackagePart package in embedobj_package_list)
                        {
                            worksheetPart.DeletePart(package);
                            success++;
                        }
                    }
                    if (embedobj_image_list.Count() > 0)
                    {
                        foreach (ImagePart image in embedobj_image_list)
                        {
                            worksheetPart.DeletePart(image);
                            success++;
                        }
                    }
                    if (embedobj_drawing_image_list.Count() > 0)
                    {
                        foreach (ImagePart drawing_image in embedobj_drawing_image_list)
                        {
                            worksheetPart.DrawingsPart.DeletePart(drawing_image);
                            success++;
                        }
                    }
                    if (embedobj_3d_list.Count() > 0)
                    {
                        foreach (Model3DReferenceRelationshipPart threeD in embedobj_3d_list)
                        {
                            worksheetPart.DeletePart(threeD);
                            success++;
                        }
                    }
                }
            }
            return success;
        }

        // Make first sheet active sheet
        public bool Activate_FirstSheet(string filepath)
        {
            bool success = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                BookViews bookViews = spreadsheet.WorkbookPart.Workbook.GetFirstChild<BookViews>();
                WorkbookView workbookView = bookViews.GetFirstChild<WorkbookView>();
                if (workbookView.ActiveTab != null)
                {
                    var activeSheetId = workbookView.ActiveTab.Value;
                    if (activeSheetId > 0)
                    {
                        // Set value in workbook.xml to first sheet
                        workbookView.ActiveTab.Value = 0;

                        // Iterate all worksheets to detect if sheetview.Tabselected exists and change it
                        IEnumerable<WorksheetPart> worksheets = spreadsheet.WorkbookPart.WorksheetParts;
                        foreach (WorksheetPart worksheet in worksheets)
                        {
                            SheetViews sheetviews = worksheet.Worksheet.SheetViews;
                            foreach (SheetView sheetview in sheetviews)
                            {
                                sheetview.TabSelected = null;
                            }
                        }
                        success = true;
                    }
                }
            }
            return success;
        }

        // Remove absolute path to local directory
        public bool Remove_AbsolutePath(string filepath)
        {
            bool success = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                if (spreadsheet.WorkbookPart.Workbook.AbsolutePath != null)
                {
                    AbsolutePath absPath = spreadsheet.WorkbookPart.Workbook.GetFirstChild<AbsolutePath>();
                    absPath.Remove();
                    success = true;
                }
            }
            return success;
        }

        // Remove metadata in file properties
        public int Remove_Metadata(string filepath)
        {
            string folder = Path.GetDirectoryName(filepath);
            int success = 0;

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

                    // Add to success
                    success++;
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

                    // Add to success
                    success++;
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

                    // Add to success
                    success++;
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

                    // Add to success
                    success++;
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

                    // Add to success
                    success++;
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

                    // Add to success
                    success++;
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

                    // Add to success
                    success++;
                }
            }
            return success;
        }

        // Work in progress
        // Change hyperlinks to link to Wayback Machine
        public int Change_Hyperlinks(string filepath)
        {
            int success = 0;
            string wayback = "https://web.archive.org/web/20230000000000*/";

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart worksheetPart in worksheetParts)
                {
                    Worksheet worksheet = worksheetPart.Worksheet;
                    IEnumerable<Hyperlink> hyperlinks = worksheet.GetFirstChild<Hyperlinks>().Elements<Hyperlink>();
                    foreach (Hyperlink hyperlink in hyperlinks)
                    {
                        Console.WriteLine(hyperlink.Id);
                        ReferenceRelationship refRel = worksheetPart.GetReferenceRelationship(hyperlink.Id);

                        // Create new hyperlink string
                        string new_Hyperlink = wayback + hyperlink.Id;

                        // Change hyperlink


                        // Add to success
                        success++;
                    }
                }
            }
            return success;
        }
    }
}