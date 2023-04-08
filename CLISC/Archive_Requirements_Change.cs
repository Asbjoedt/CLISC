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
        public void Change_XLSX_Requirements(List<Archive_Requirements> arcReq, string filepath, bool fullcompliance)
        {
            foreach (var item in arcReq)
            {
                if (item.Conformance == true)
                {
                    Change_ConformanceToStrict_ExcelInterop(filepath);
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
                if (item.ExternalObj > 0)
                {
                    Tuple<int, int> success = CopyAndRemove_ExternalObjects(filepath);
                    if (success.Item1 > 0)
                    {
                        Console.WriteLine($"--> Change: {success.Item1} external object references were copied to new subfolder and removed");
                    }
                    if (success.Item2 > 0)
                    {
                        Console.WriteLine($"--> ChangeError: {success.Item2} external object references were removed but NOT copied to new subfolder \"External objects\"");
                    }
                }
                if (item.EmbedObj > 0)
                {
                    Tuple<int, int, int> success = Convert_EmbeddedObjects(filepath);
                    if (success.Item1 > 0)
                    {
                        Console.WriteLine($"--> Extract: {success.Item1} original embedded objects were saved to new \"Embedded objects\" subfolder.");
                    }
                    if (success.Item2 > 0)
                    {
                        Console.WriteLine($"--> Change: {success.Item2} embedded objects were converted");
                    }
                    if (success.Item3 > 0)
                    {
                        Console.WriteLine($"--> ChangeError: {success.Item3} embedded objects could not be processed");
                    }
                }
                if (fullcompliance)
                {
                    if (item.Metadata == true)
                    {
                        int success = Remove_Metadata(filepath);
                        Console.WriteLine($"--> Change: {success} file property information were removed and saved to sidecar file");
                    }
                    if (item.PrinterSettings > 0)
                    {
                        int success = Remove_PrinterSettings(filepath);
                        Console.WriteLine($"--> Change: {success} printer settings were removed");
                    }
                    if (item.ActiveSheet == true)
                    {
                        bool success = Activate_FirstSheet(filepath);
                        if (success)
                        {
                            Console.WriteLine("--> Change: First sheet was activated");
                        }
                        else
                        {
                            Console.WriteLine("--> ChangeError: First sheet was NOT activated");
                        }
                    }
                    if (item.AbsolutePath == true)
                    {
                        bool success = Remove_AbsolutePath(filepath);
                        if (success)
                        {
                            Console.WriteLine("--> Change: Absolute path to local directory was removed");
                        }
                        else
                        {
                            Console.WriteLine("--> ChangeError: Absolute path to local directory was NOT removed");
                        }
                    }
                    if (item.Hyperlinks > 0)
                    {
                        int success = Extract_Hyperlinks(filepath);
                        Console.WriteLine($"--> Extract: {success} cell hyperlinks were extracted");
                    }
                }
            }
        }

        // Work in progress
        // Change conformance to Strict
        public void Change_ConformanceToStrict(string filepath)
        {
            // Create list of namespaces
            List<namespaceIndex> namespaces = namespaceIndex.Create_Namespaces_Index();

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
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
                
                // Delete all QueryTableParts
                IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart worksheetPart in worksheetParts)
                {
                    // Delete all QueryTableParts in WorksheetParts
                    List<QueryTablePart> queryTables = worksheetPart.QueryTableParts.ToList(); // Must be a list
                    foreach (QueryTablePart queryTablePart in queryTables)
                    {
                        worksheetPart.DeletePart(queryTablePart);
                    }

                    // Delete all QueryTableParts, if they are not registered in a WorksheetPart
                    List<TableDefinitionPart> tableDefinitionParts = worksheetPart.TableDefinitionParts.ToList();
                    foreach (TableDefinitionPart tableDefinitionPart in tableDefinitionParts)
                    {
                        List<IdPartPair> idPartPairs = tableDefinitionPart.Parts.ToList();
                        foreach (IdPartPair idPartPair in idPartPairs)
                        {
                            if (idPartPair.OpenXmlPart.ToString() == "DocumentFormat.OpenXml.Packaging.QueryTablePart")
                            {
                                // Delete QueryTablePart
                                tableDefinitionPart.DeletePart(idPartPair.OpenXmlPart);
                                // The TableDefinitionPart must also be deleted
                                worksheetPart.DeletePart(tableDefinitionPart);
                                // And the reference to the TableDefinitionPart in the WorksheetPart must be deleted
                                List<TablePart> tableParts = worksheetPart.Worksheet.Descendants<TablePart>().ToList();
                                foreach (TablePart tablePart in tableParts)
                                {
                                    if (idPartPair.RelationshipId == tablePart.Id)
                                    {
                                        tablePart.Remove();
                                    }
                                }
                            }
                        }
                    }
                }

                // If spreadsheet contains a CustomXmlMappingsPart, delete databinding
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
        public Tuple<int, int> CopyAndRemove_ExternalObjects(string filepath)
        {
            int success = 0;
            int fail = 0;

            // Create new subfolder for external objects
            int backslash = filepath.LastIndexOf("\\");
            string file_folder = filepath.Substring(0, backslash);
            string new_folder = file_folder + "\\External objects";
            Directory.CreateDirectory(new_folder);

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<ExternalWorkbookPart> extWbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts;
                foreach (ExternalWorkbookPart extWbPart in extWbParts)
                {
                    List<ExternalRelationship> extrels = extWbPart.ExternalRelationships.ToList(); // Must be a list
                    foreach (ExternalRelationship extrel in extrels)
                    {
                        // Copy external file to subfolder
                        string output_filepath = new_folder + "\\" + extrel.Uri.ToString().Split("/").Last();
                        try
                        {
                            File.Copy(extrel.Uri.ToString(), output_filepath);
                            success++;
                        }
                        catch(System.IO.IOException)
                        {
                            fail++;
                        }

                        // Remove external object reference
                        Uri uri = new Uri($"External reference {extrel.Uri} was removed", UriKind.Relative);
                        extWbPart.DeleteExternalRelationship(extrel.Id);
                        extWbPart.AddExternalRelationship(relationshipType: "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject", externalUri: uri, id: extrel.Id);
                    }
                }
            }
            // Delete new subfolder, if no objects were copied to it
            if (Directory.GetFiles(new_folder).Length == 0)
            {
                Directory.Delete(new_folder);
            }
            return System.Tuple.Create(success, fail);
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
            string folder = System.IO.Path.GetDirectoryName(filepath);
            int success = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                PackageProperties property = spreadsheet.PackageProperties;

                // Create metadata file
                using (StreamWriter w = File.AppendText($"{folder}\\orgFile_Metadata.txt"))
                {
                    w.WriteLine("---");
                    w.WriteLine("EXTRACTED METADATA");
                    w.WriteLine("---");

                    if (property.Creator != null)
                    {
                        // Write information to metadata file
                        w.WriteLine($"CREATOR: {property.Creator}");

                        // Remove information
                        property.Creator = null;

                        // Add to success
                        success++;
                    }
                    if (property.Title != null)
                    {
                        // Write information to metadata file
                        w.WriteLine($"TITLE: {property.Title}");

                        // Remove information
                        property.Title = null;

                        // Add to success
                        success++;
                    }
                    if (property.Subject != null)
                    {
                        // Write information to metadata file
                        w.WriteLine($"SUBJECT: {property.Subject}");

                        // Remove information
                        property.Subject = null;

                        // Add to success
                        success++;
                    }
                    if (property.Description != null)
                    {
                        // Write information to metadata file
                        w.WriteLine($"DESCRIPTION: {property.Description}");

                        // Remove information
                        property.Description = null;

                        // Add to success
                        success++;
                    }
                    if (property.Keywords != null)
                    {
                        // Write information to metadata file
                        w.WriteLine($"KEYWORDS: {property.Keywords}");

                        // Remove information
                        property.Keywords = null;

                        // Add to success
                        success++;
                    }
                    if (property.Category != null)
                    {
                        // Write information to metadata file
                        w.WriteLine($"CATEGORY: {property.Category}");

                        // Remove information
                        property.Category = null;

                        // Add to success
                        success++;
                    }
                    if (property.LastModifiedBy != null)
                    {
                        // Write information to metadata file
                        w.WriteLine($"LAST MODIFIED BY: {property.LastModifiedBy}");

                        // Remove information
                        property.LastModifiedBy = null;

                        // Add to success
                        success++;
                    }
                }
            }
            return success;
        }

        // Extract all cell hyperlinks to an external file
        public int Extract_Hyperlinks(string filepath)
        {
            string folder = System.IO.Path.GetDirectoryName(filepath);
            int hyperlinks_count = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                List<HyperlinkRelationship> hyperlinks = spreadsheet
                    .GetAllParts()
                    .SelectMany(p => p.HyperlinkRelationships)
                    .ToList();

                // Create metadata file
                using (StreamWriter w = File.AppendText($"{folder}\\orgFile_Metadata.txt"))
                {
                    w.WriteLine("---");
                    w.WriteLine("EXTRACTED HYPERLINKS");
                    w.WriteLine("---");

                    foreach (HyperlinkRelationship hyperlink in hyperlinks)
                    {
                        // Write information to metadata file
                        w.WriteLine(hyperlink.Uri);
                        // Add to count
                        hyperlinks_count++;
                    }
                }
            }
            return hyperlinks_count;
        }
    }
}