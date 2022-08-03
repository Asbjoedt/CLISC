using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CLISC
{
    public partial class Archive
    {
        public static int extrels_files = 0;
        public static int rtdfunctions_files = 0;
        public static int embedobj_files = 0;

        // Perform data quality actions
        public string Check_DataQuality(string filepath) // Method ONLY checks
        {
            string dataquality_message = "";

            try // call the methods
            {
                string extrels_message = Check_ExternalRelationships(filepath);
                string embedobj_message = Alert_EmbeddedObjects(filepath);
                bool rtdfunctions = Simple_Check_RTDFunctions(filepath);

                string messages_combined = extrels_message + ", " + embedobj_message + ", " + rtdfunctions;

                return messages_combined;
            }

            catch (ArgumentNullException)
            {
                // BUG: Method cannot handle null filepaths. Must handle exception to it
                dataquality_message = "";

                return dataquality_message;
            }
        }

        public void Simple_Check_and_Remove_DataQuality(string filepath)
        {
            // Check for data to change
            bool extrels = Simple_Check_ExternalRelationships(filepath);
            bool rtdfunctions = Simple_Check_RTDFunctions(filepath);

            // If true, change data
            if (extrels == true)
            {
                //Remove_ExternalRelationships(filepath);
                Console.WriteLine($"--> External relationships removed - To prevent data loss, manually handle data and reconvert from original");
            }
            if (rtdfunctions == true)
            {
                Remove_RTDFunctions(filepath);
                Console.WriteLine($"--> RTD functions removed");
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

        public string Check_ExternalRelationships(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                List<ExternalRelationship> extRels = spreadsheet // Find all external relationships
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
                        Console.WriteLine($"--> #{extrel_number} external relationship detected");
                        Console.WriteLine($"----> ID: {rel.Id}");
                        Console.WriteLine($"----> Target URI: {rel.Uri}");
                        Console.WriteLine($"----> Relationship type: {rel.RelationshipType}");
                        Console.WriteLine($"----> External: {rel.IsExternal}");
                        Console.WriteLine($"----> Container: {rel.Container}");
                    }
                }
                else // If no external relationships, inform user
                {
                    Console.WriteLine($"--> {extRels.Count} external relationships");
                }

                string extrels_message = extRels.Count + "external relationships detected";
                return extrels_message;
            }
        }

        // Remove external relationships
        public void Remove_ExternalRelationships(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                List<ExternalRelationship> extRels = spreadsheet // Find all external relationships
                .GetAllParts()
                .SelectMany(p => p.ExternalRelationships)
                .ToList();

                foreach (ExternalRelationship rel in extRels) // Remove each external relationship
                {
                    spreadsheet.DeleteExternalRelationship(rel.Id);
                }

                // Save and close spreadsheet
                spreadsheet.Save();
                spreadsheet.Close();
            }
        }

        // Check for RTD functions and return alert
        public static bool Simple_Check_RTDFunctions(string filepath)
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

        public void Remove_RTDFunctions(string filepath)
        {
            string rtdfunctions_message = "";

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {


            }
        }

        public bool Simple_Alert_EmbeddedObjects(string filepath)
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

        // Check for embedded objects and return alert
        public string Alert_EmbeddedObjects(string filepath)
        {
            string embedobj_message = "";

            using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                var list = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (var item in list)
                {
                    int count_ole = item.EmbeddedObjectParts.Count(); // Register the number of OLE
                    int count_image = item.ImageParts.Count(); // Register number of images
                    int count = count_ole + count_image; // Sum
                    if (count == 0) // If no embedded objects, inform user
                    {
                        embedobj_message = $"--> {count} embedded objects detected";
                        Console.WriteLine(embedobj_message);
                        return embedobj_message;
                    }
                    else
                    {
                        embedobj_message = $"--> {count} embedded objects detected";
                        Console.WriteLine(embedobj_message);
                        var embed_ole = item.EmbeddedObjectParts.ToList(); // Register each OLE to a list
                        var embed_image = item.ImageParts.ToList(); // Register each image to a list
                        int embedobj_no = 0;
                        foreach (var part in embed_ole) // Inform user of each object
                        {
                            embedobj_no++;
                            Console.WriteLine($"--> Embedded object #{embedobj_no}");
                            Console.WriteLine($"----> Content Type: {part.ContentType.ToString()}");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");

                        }
                        foreach (var part in embed_image) // Inform user of each object
                        {
                            embedobj_no++;
                            Console.WriteLine($"--> Embedded object #{embedobj_no}");
                            Console.WriteLine($"----> Content Type: {part.ContentType.ToString()}");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                        }
                        return embedobj_message;
                    }
                }
                return embedobj_message;
            }
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