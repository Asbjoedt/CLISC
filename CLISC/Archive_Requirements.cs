using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

namespace CLISC
{
    public partial class Archive
    {
        public Tuple<bool, int, int, int, int, int> Check_XLSX_Requirements(string filepath) 
        {
            bool data = Check_Value(filepath);
            int connections = Check_DataConnections(filepath);
            int extrels = Check_ExternalRelationships(filepath);
            int rtdfunctions = Check_RTDFunctions(filepath);
            int embedobj = Check_EmbeddedObjects(filepath);
            int hyperlinks = Check_Hyperlinks(filepath);

            (bool, int, int, int, int, int) pidgeon = (data, connections, extrels, rtdfunctions, embedobj, hyperlinks);
            return pidgeon.ToTuple();
        }

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
        }

        // Check for any values by checking if sheets and cell values exist
        public bool Check_Value(string filepath)
        {
            bool hascellvalue = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                //Check if worksheets exist
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Sheets allSheets = wbPart.Workbook.Sheets;
                if (allSheets == null)
                {
                    Console.WriteLine("--> No cell values detected");
                    return hascellvalue;
                }
                // Check if any cells have any value
                foreach (Sheet aSheet in allSheets)
                {
                    WorksheetPart wsp = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(aSheet.Id);
                    DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = wsp.Worksheet;
                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                    int row_count = rows.Count(); // Count number of rows
                    if (row_count > 0) // If any rows exist, this means cells exist
                    {
                        hascellvalue = true;
                        return hascellvalue;
                    }
                }
            }
            Console.WriteLine("--> No cell values detected");
            return hascellvalue;
        }

        // Check for data connections
        public int Check_DataConnections(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                int conn_count = wbPart.ConnectionsPart.Connections.Count();
                if (conn_count > 0)
                {
                    Console.WriteLine($"--> {conn_count} data connections detected and removed");
                }
                return conn_count;
            }
        }

        // Remove data connections
        public void Remove_DataConnections(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                var conn_list = wbPart.ConnectionsPart.Connections.ToList();
                if (conn_list.Any())
                {
                    foreach (Connection conn in conn_list)
                    {
                        wbPart.DeleteReferenceRelationship(conn.Id);
                        conn.Deleted = true;
                        conn.Remove();
                    }
                }
            }
        }

        // Check for external relationships
        public static bool Simple_Check_ExternalRelationships(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                List<ExternalRelationship> extRels = spreadsheet // Find all external relationships
                .GetAllParts()
                .SelectMany(p => p.ExternalRelationships)
                .ToList();

                bool check = false;
                if (extRels.Count > 0)
                {
                    check = true;
                }
                return check;
            }
        }

        public int Check_ExternalRelationships(string filepath) // Find all external relationships
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                List<ExternalRelationship> extrels = spreadsheet 
                .GetAllParts()
                .SelectMany(p => p.ExternalRelationships)
                .ToList();
                int extrels_count = extrels.Count;

                if (extrels.Count > 0) // If external relationships
                {
                    int extrel_number = 0;
                    Console.WriteLine($"--> {extrels.Count} external relationships detected");
                    foreach (ExternalRelationship rel in extrels)
                    {
                        extrel_number++;
                        Console.WriteLine($"--> External relationship {extrel_number}");
                        Console.WriteLine($"----> ID: {rel.Id}");
                        Console.WriteLine($"----> Target URI: {rel.Uri}");
                        Console.WriteLine($"----> Relationship type: {rel.RelationshipType}");
                        Console.WriteLine($"----> External: {rel.IsExternal}");
                    }
                    return extrels_count;
                }
                return extrels_count;
            }
        }

        public void Handle_ExternalRelationships(string filepath) // Handle external relationships
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                List<ExternalRelationship> extRels = spreadsheet // Find all external relationships
                .GetAllParts()
                .SelectMany(p => p.ExternalRelationships)
                .ToList();

                foreach (ExternalRelationship rel in extRels)
                {
                    if (rel.IsExternal == true)
                    {
                        switch (rel.RelationshipType)
                        {
                            // Embed linked cell values and remove relationship
                            case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath":
                            case "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLinkPath":
                                // Replace formula values with cell values 

    

                                // Remember to finish with removing relationshipId


                                // Inform user
                                Console.WriteLine("--> External cell values detected. All cell values were embedded and the relationship removed");
                                break;

                            // Alert if the relationship is an OLE object
                            case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject":
                            case "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject":
                                // Inform user
                                Console.WriteLine("--> External OLE object detected. Handle OLE object manually");
                                break;
                        }
                    }
                }
                // Save and close spreadsheet
                spreadsheet.Save();
                spreadsheet.Close();
            }
        }

        public static bool Simple_Check_RTDFunctions(string filepath) // Check for RTD functions
        {
            bool check = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
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
                                        check = true;
                                        return check;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return check;
        }

        public static int Check_RTDFunctions(string filepath) // Check for RTD functions
        {
            int rtd_functions = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
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
                                        rtd_functions++;
                                        Console.WriteLine($"--> RTD function in sheet \"{aSheet.Name}\" cell {cell.CellReference} detected");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return rtd_functions;
        }

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

        public bool Simple_Check_EmbeddedObjects(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                var list = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (var item in list)
                {
                    int count_ole = item.EmbeddedObjectParts.Count(); // Register the number of OLE
                    int count_image = item.ImageParts.Count(); // Register number of images
                    int count = count_ole + count_image; // Sum
                    if (count == 0) // If no embedded objects, inform user
                    {
                        string embedobj_message = $"--> {count} embedded objects detected. Extract objects manually";
                        Console.WriteLine(embedobj_message);
                        return false;
                    }
                    else
                    {
                        string embedobj_message = $"--> {count} embedded objects detected. Extract objects manually";
                        Console.WriteLine(embedobj_message);
                        return true;
                    }
                }
                return false;
            }
        }

        public int Check_EmbeddedObjects(string filepath) // Check for embedded objects and return alert
        {
            using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                var list = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (var item in list)
                {
                    int count_ole = item.EmbeddedObjectParts.Count(); // Register the number of OLE
                    int count_image = item.ImageParts.Count(); // Register number of images
                    int count_3d = item.Model3DReferenceRelationshipParts.Count(); // Register number of 3D models
                    int count_embedobj = count_ole + count_image + count_3d; // Sum

                    if (count_embedobj > 0) // If embedded objects
                    {
                        Console.WriteLine($"--> {count_embedobj} embedded objects detected");
                        var embed_ole = item.EmbeddedObjectParts.ToList(); // Register each OLE to a list
                        var embed_image = item.ImageParts.ToList(); // Register each image to a list
                        var embed_3d = item.Model3DReferenceRelationshipParts.ToList(); // Register each 3D model to a list
                        int embedobj_number = 0;
                        foreach (var part in embed_ole) // Inform user of each OLE object
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: {part.ContentType.ToString()}");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");

                        }
                        foreach (var part in embed_image) // Inform user of each image object
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: {part.ContentType.ToString()}");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                        }
                        foreach (var part in embed_3d) // Inform user of each 3D object
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: {part.ContentType.ToString()}");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                        }
                    }
                    return count_embedobj;
                }
                return 0;
            }
        }

        public int Check_Hyperlinks(string filepath) // Find all hyperlinks
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                List<HyperlinkRelationship> hyperlinks = spreadsheet
                    .GetAllParts()
                    .SelectMany(p => p.HyperlinkRelationships)
                    .ToList();
                int hyperlinks_count = hyperlinks.Count;

                if (hyperlinks_count > 0) // If hyperlinks
                {
                    Console.WriteLine($"--> {hyperlinks_count} hyperlinks detected");
                    int hyperlink_number = 0;
                    foreach (HyperlinkRelationship hyperlink in hyperlinks)
                    {
                        hyperlink_number++;
                        Console.WriteLine($"--> Hyperlink: {hyperlink_number}");
                        Console.WriteLine($"----> Address: {hyperlink.Uri}");
                    }

                    return hyperlinks_count;
                }
                return hyperlinks_count;
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
                    i = i-1;
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

            // Make first sheet active sheet
            ((Excel.Worksheet)app.ActiveWorkbook.Sheets[1]).Select();

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