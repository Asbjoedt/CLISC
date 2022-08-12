using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel = Microsoft.Office.Interop.Excel;

namespace CLISC
{
    public partial class Archive
    {
        public static int extrels_files = 0;
        public static int rtdfunctions_files = 0;
        public static int embedobj_files = 0;

        public Tuple<bool, int, int, int, int, int> Check_XLSX_Requirements(string filepath) 
        {
            bool data = Check_Value(filepath);
            int connections = Simple_Check_DataConnections(filepath);
            int extrels = Check_ExternalRelationships(filepath);
            int rtdfunctions = Check_RTDFunctions(filepath);
            int embedobj = Check_EmbeddedObjects(filepath);
            int hyperlinks = Check_Hyperlinks(filepath);

            (bool, int, int, int, int, int) pidgeon = (data, connections, extrels, rtdfunctions, embedobj, hyperlinks);
            return pidgeon.ToTuple();
        }

        // Get all worksheets in a spreadsheet
        public bool Check_Value(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                //Check if worksheets exist
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                Sheets theSheets = wbPart.Workbook.Sheets;
                if (theSheets == null)
                {
                    Console.WriteLine("--> Spreadsheet has no cell information");
                    return false;
                }

                // Check if any cells have any value


                return true;
            }
        }

        // Check for data connections
        public int Simple_Check_DataConnections(string filepath) // Using Excel interop
        {
            Excel.Application app = new Excel.Application();
            app.DisplayAlerts = false;
            Excel.Workbook wb = app.Workbooks.Open(filepath);
            int count_conn = wb.Connections.Count;
            wb.Close();
            app.Quit();
            return count_conn;
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
                                Console.WriteLine("--> Linked cell values detected. All cell values were embedded and the relationship removed");
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

        public static bool Simple_Check_RTDFunctions(string filepath) // Check for RTD functions and return alert
        {

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                var rtd_functions = "";

                bool check = false;
                if (rtd_functions != "")
                {
                    check = true;
                }
                return check;
            }
        }

        public static int Check_RTDFunctions(string filepath) // Check for RTD functions
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                return 0;
            }
        }

        public void Remove_RTDFunctions(string filepath) // Remove RTD functions
        {
            string rtdfunctions_message = "";

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {

                Console.WriteLine($"--> RTD functions removed");
            }
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

        public void Transform_Requirements(string filepath)  // Use Excel Interop
        {
            Excel.Application app = new Excel.Application(); // Create Excel object instance
            app.DisplayAlerts = false; // Don't display any Excel window prompts
            Excel.Workbook wb = app.Workbooks.Open(filepath); // Create workbook instance

            // Find any cell value
            int used_cells_count = 0;
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                try
                {
                    Excel.Range range = (Excel.Range)sheet.UsedRange;
                    foreach (Excel.Range cell in range.Cells)
                    {
                        var value = cell.Value2;
                        if (value != null)
                        {
                            used_cells_count++;
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException) // Catch if cell has no value
                {
                    // Do nothing
                }
                finally
                {
                    if (used_cells_count == 0)
                    {
                        Console.WriteLine("--> No cell values detected. Exempt spreadsheet from archiving");
                    }
                }
            }

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
                Console.WriteLine("--> Data connections detected and removed");
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
                        string address = cell.Address;
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
            }

            wb.Save(); // Save workbook
            wb.Close(); // Close the workbook
            app.Quit(); // Quit Excel application
        }

        // Retrieve the value of a cell, given a file name, sheet name, and address name.
        // Source: https://docs.microsoft.com/en-us/office/open-xml/how-to-retrieve-the-values-of-cells-in-a-spreadsheet
        public static string GetCellValue(string fileName, string sheetName, string addressName)
        {
            string value = null;

            // Open the spreadsheet document for read-only access.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                // Retrieve a reference to the workbook part.
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetName).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                // Use its Worksheet property to get a reference to the cell whose address matches the address you supplied.
                Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                  Where(c => c.CellReference == addressName).FirstOrDefault();

                // If the cell does not exist, return an empty string.
                if (theCell.InnerText.Length > 0)
                {
                    value = theCell.InnerText;

                    // If the cell represents an integer number, you are done. For dates, this code returns the serialized value that represents the date. The code handles strings and Booleans individually. For shared strings, the code looks up the corresponding value in the shared string table. For Booleans, the code converts the value into the words TRUE or FALSE.
                    if (theCell.DataType != null)
                    {
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:

                                // For shared strings, look up the value in the shared strings table.
                                var stringTable =
                                    wbPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault();

                                // If the shared string table is missing, something is wrong. Return the index that is in the cell. Otherwise, look up the correct text in the table.
                                if (stringTable != null)
                                {
                                    value =
                                        stringTable.SharedStringTable
                                        .ElementAt(int.Parse(value)).InnerText;
                                }
                                break;

                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }
                }
            }
            return value;
        }
    }
}