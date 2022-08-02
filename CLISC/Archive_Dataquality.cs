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
        public string Check_DataQuality(string filepath)
        {
            string dataquality_message = "";

            try
            {
                // call the methods
                string extrels_message = Check_ExternalRelationships(filepath);
                bool rtdfunctions = Simple_Check_RTDFunctions(filepath);
                string embedobj_message = Alert_EmbeddedObjects(filepath);

                string messages_combined = "";

                return messages_combined;
            }

            catch (ArgumentNullException)
            {
                // BUG: Method cannot handle null filepaths. Must handle exception to it
                dataquality_message = "";

                return dataquality_message;
            }
        }

        public void Check_and_Remove_DataQuality(string filepath)
        {
            // Check for data to change
            bool extrels = Simple_Check_ExternalRelationships(filepath);
            bool rtdfunctions = Simple_Check_RTDFunctions(filepath);

            // If true, change data
            if (extrels == true)
            {
                Remove_ExternalRelationships(filepath);
                //Console.WriteLine($"--> External relationships removed"); <- UNCOMMENT THIS FOR MESSAGE OF REMOVAL
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
            // Open spreadsheet and find external relationships
            SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false);
            var external_relationships = spreadsheet.ExternalRelationships.ToList();
            spreadsheet.Close();
            Console.WriteLine($"DEBUG - {external_relationships}");
            bool check = false;
            if (external_relationships.Count == 0)
            {
                check = true;
            }
            return check;
        }

        public string Check_ExternalRelationships(string filepath)
        {
            // Open spreadsheet and find external relationships
            SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false);
            var external_relationships = spreadsheet.ExternalRelationships.ToList();
            spreadsheet.Close();

            // Data types
            int extrels_count = external_relationships.Count;
            int extrel_number = 0;
            string extrels_message;

            // If errors
            if (external_relationships.Any())
            {
                // Inform user
                Console.WriteLine(external_relationships); // To test if any errors are found and added to the list
                Console.WriteLine($"--> {extrels_count} relationships detected");
                foreach (var extrel in external_relationships)
                {
                    extrel_number++;
                    Console.WriteLine($"--> External relationship {extrel_number}");
                    Console.WriteLine("----> Relationship ID: " + extrel.Id);
                    Console.WriteLine("----> Relationship type: " + extrel.RelationshipType);
                    Console.WriteLine("----> Relationship target URI: " + extrel.Uri);
                    Console.WriteLine("----> Relationship external: " + extrel.IsExternal);
                    Console.WriteLine("----> Relationship container: " + extrel.Container);
                }
                // Add to number of spreadsheets with external relationships
                extrels_files++;
                // Turn list into string
                extrels_message = string.Join(Environment.NewLine, external_relationships);

                return extrels_message;
            }
            else
            {
                // If no errors, inform user
                Console.WriteLine("--> No external relationships detected");
                extrels_message = $"{extrels_count} external relationships";

                return extrels_message;
            }
        }

        // Remove external relationships
        public void Remove_ExternalRelationships(string filepath)
        {
            // Open spreadsheet and remove external relationships
            SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true);
            var external_relationships = spreadsheet.ExternalRelationships.ToList();

            //external_relationships.Remove(ExternalRelationship, extrel.Id);
            //spreadsheet.Save();
            //spreadsheet.Close();
            // Inform user
            //Console.WriteLine($"--> External relationship {extrel_number} removed");
            // Add to number of spreadsheets with external relationships

            spreadsheet.Close();
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
            using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(filepath, true))
            {
                WorksheetPart sourceSheetPart = GetWorksheetPartByName(spreadsheetDoc, "Test");

                var imagePart = sourceSheetPart.AddImagePart(ImagePartType.Emf, "rId1");
                imagePart.FeedData(File.Open(placeholderImagePath, FileMode.Open));

                var embeddedObject =
                    sourceSheetPart.AddEmbeddedObjectPart(@"application/vnd.openxmlformats-officedocument.oleObject");
                embeddedObject.FeedData(File.Open(embeddedFilePath, FileMode.Open));

            }
            return true;
        }

        // Check for embedded objects and return alert
        public string Alert_EmbeddedObjects(string filepath)
        {
            string embedobj_message = "";

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                var embedded_objects = spreadsheet.ExternalRelationships.ToList();
                int embedobj_count = embedded_objects.Count;
                int embedobj_number = 0;

                // If errors
                if (embedded_objects.Any())
                {
                    // Inform user
                    Console.WriteLine($"--> {embedobj_count} embedded objects detected");
                    foreach (var extrel in embedded_objects)
                    {
                        embedobj_number++;
                        Console.WriteLine($"--> External relationship {embedobj_number}");
                        Console.WriteLine("----> Relationship ID: " + extrel.Id);
                        Console.WriteLine("----> Relationship type: " + extrel.RelationshipType);
                        Console.WriteLine("----> Relationship target URI: " + extrel.Uri);
                        Console.WriteLine("----> Relationship external: " + extrel.IsExternal);
                        Console.WriteLine("----> Relationship container: " + extrel.Container);
                    }
                    // Add to number of spreadsheets with external relationships
                    embedobj_files++;
                    // Turn list into string
                    embedobj_message = string.Join(Environment.NewLine, embedded_objects);

                    return embedobj_message;
                }

                else
                {
                    // If no errors, inform user
                    Console.WriteLine("--> No embedded objects detected");
                    embedobj_message = $"{embedobj_count} embedded objects relationships";

                    return embedobj_message;
                }

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

                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetName).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                // Use its Worksheet property to get a reference to the cell 
                // whose address matches the address you supplied.
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

                                // For shared strings, look up the value in the
                                // shared strings table.
                                var stringTable =
                                    wbPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault();

                                // If the shared string table is missing, something 
                                // is wrong. Return the index that is in
                                // the cell. Otherwise, look up the correct text in 
                                // the table.
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
