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
    }
}
