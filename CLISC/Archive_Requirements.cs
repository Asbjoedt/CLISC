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

        
        public string Check_Requirements(string filepath) // Method ONLY checks archival requirements
        {
            string dataquality_message = "";

            try // call the methods
            {
                string extrels_message = Check_ExternalRelationships(filepath);
                string embedobj_message = Check_EmbeddedObjects(filepath);
                bool rtdfunctions = Simple_Check_RTDFunctions(filepath);
                string hyperlinks_message = Check_Hyperlinks(filepath);

                string messages_combined = extrels_message + ", " + embedobj_message + ", " + rtdfunctions + ", " + hyperlinks_message;

                return messages_combined;
            }

            catch (ArgumentNullException)
            {
                // BUG: Method cannot handle null filepaths. Must handle exception to it
                dataquality_message = "";

                return dataquality_message;
            }
        }

        public void Simple_Check_and_Transform_Requirements(string filepath)
        {
            bool extrels = Simple_Check_ExternalRelationships(filepath); // Check for external relationships
            if (extrels == true)
            {
                Handle_ExternalRelationships(filepath);
            }

            bool rtdfunctions = Simple_Check_RTDFunctions(filepath); // Check for RTD functions
            if (rtdfunctions == true)
            {
                Remove_RTDFunctions(filepath);
            }

            Simple_Interop(filepath); // Use Excel Interop to cheat

            Mark_ReadOnly(filepath); // Mark spreadsheet with read only prompt before enabling editing
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

        public string Check_ExternalRelationships(string filepath) // Find all external relationships
        {
            string extrels_message = "";

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                List<ExternalRelationship> extRels = spreadsheet 
                .GetAllParts()
                .SelectMany(p => p.ExternalRelationships)
                .ToList();

                if (extRels.Count > 0) // If external relationships
                {
                    int extrel_number = 0;
                    Console.WriteLine($"--> {extRels.Count} external relationships detected");
                    foreach (ExternalRelationship rel in extRels)
                    {
                        extrel_number++;
                        Console.WriteLine($"--> External relationship {extrel_number}");
                        Console.WriteLine($"----> ID: {rel.Id}");
                        Console.WriteLine($"----> Target URI: {rel.Uri}");
                        Console.WriteLine($"----> Relationship type: {rel.RelationshipType}");
                        Console.WriteLine($"----> External: {rel.IsExternal}");
                    }
                    extrels_message = extRels.Count + " external relationships detected";
                    return extrels_message;
                }

                return extrels_message;
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
                        string embedobj_message = $"--> {count} embedded objects detected";
                        Console.WriteLine(embedobj_message);
                        return false;
                    }
                    else
                    {
                        string embedobj_message = $"--> {count} embedded objects detected";
                        Console.WriteLine(embedobj_message);
                        return true;
                    }
                }
                return false;
            }
        }

        public string Check_EmbeddedObjects(string filepath) // Check for embedded objects and return alert
        {
            string embedobj_message = "";

            using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                var list = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (var item in list)
                {
                    int count_ole = item.EmbeddedObjectParts.Count(); // Register the number of OLE
                    int count_image = item.ImageParts.Count(); // Register number of images
                    int count_3d = item.Model3DReferenceRelationshipParts.Count(); // Register number of 3D models
                    int count = count_ole + count_image + count_3d; // Sum

                    if (count > 0) // If no embedded objects
                    {
                        Console.WriteLine($"--> {count} embedded objects detected");
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
                        embedobj_message = count + " embedded objects detected";
                        return embedobj_message;
                    }
                }
            }
            return embedobj_message;
        }

        public string Check_Hyperlinks(string filepath) // Find all hyperlinks
        {
            string hyperlinks_message = "";

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                List<HyperlinkRelationship> hyperlinks = spreadsheet
                    .GetAllParts()
                    .SelectMany(p => p.HyperlinkRelationships)
                    .ToList();
                int hyperlinks_count = hyperlinks.Count;

                if (hyperlinks.Count > 0) // If hyperlinks
                {
                    Console.WriteLine($"--> {hyperlinks_count} hyperlinks detected");
                    int hyperlink_number = 0;
                    foreach (HyperlinkRelationship hyperlink in hyperlinks)
                    {
                        hyperlink_number++;
                        Console.WriteLine($"--> Hyperlink: {hyperlink_number}");
                        Console.WriteLine($"----> Address: {hyperlink.Uri}");
                    }
                    hyperlinks_message = hyperlinks_count + " external relationships detected";
                    return hyperlinks_message;
                }
                return hyperlinks_message;
            }
        }

        public void Mark_ReadOnly(string filepath)
        {

        }

        public void Simple_Interop(string filepath)
        {
            Excel.Application app = new Excel.Application(); // Create Excel object instance
            app.DisplayAlerts = false; // Don't display any Excel prompts
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
                Console.WriteLine("--> Data connections detected and removed");
            }

            // Find and delete file property details
            if (wb.Author != "" || wb.Subject != "" || wb.Comments != "") // Inform user of removal of file property details
            {
                Console.WriteLine("--> Removed file property details");
            }

            wb.Author = ""; // Remove author information
            wb.Subject = ""; // Remove subject information
            wb.Comments = ""; // Remove comments information

            wb.Save(); // Save workbook
            wb.Close(); // Close the workbook
            app.Quit(); // Quit Excel application
        }

        // Get all worksheets in a spreadsheet
        // Source: https://docs.microsoft.com/en-us/office/open-xml/how-to-retrieve-a-list-of-the-worksheets-in-a-spreadsheet
        public static Sheets GetAllWorksheets(string filepath)
        {
            Sheets theSheets = null;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filepath, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                theSheets = wbPart.Workbook.Sheets;
            }
            return theSheets;
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