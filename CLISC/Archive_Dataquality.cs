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
    public partial class Spreadsheet
    {
        public int extrels_files = 0;
        public int rtdfunctions_files = 0;
        public int embedobj_files = 0;

        // Perform data quality actions
        public string Transform_DataQuality(string filepath)
        {
            string dataquality_message = "";

            // Check for external relationships
            try
            {
                Console.WriteLine(filepath);

                // call the methods
                string extrels_message = Remove_ExternalRelationships(filepath);
                string rtdfunctions_message = Remove_RTDFunctions(filepath);
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

        // Remove external relationships
        public string Remove_ExternalRelationships(string filepath)
        {
            string extrels_message;

            SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false);
            var external_relationships = spreadsheet.ExternalRelationships.ToList();
            int extrels_count = external_relationships.Count;
            int extrel_number = 0;
            spreadsheet.Close();

            // If errors
            if (external_relationships.Any())
            {
                // Inform user
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
                // Open spreadsheet to remove external relationships
                //spreadsheet = SpreadsheetDocument.Open(filepath, true);
                //external_relationships.Remove(ExternalRelationship, extrel.Id);
                //spreadsheet.Save();
                //spreadsheet.Close();
                // Inform user
                //Console.WriteLine($"--> External relationship {extrel_number} removed");
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
                extrels_message = "0 external relationships";

                return extrels_message;
            }
        }

        // Check for embedded objects and return alert
        public string Remove_RTDFunctions(string filepath)
        {
            string rtdfunctions_message = "";


            return rtdfunctions_message;
        }


        // Check for embedded objects and return alert
        public string Alert_EmbeddedObjects(string filepath)
        {
            string embedobj_message = "";

            SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false);
            var embedded_objects = spreadsheet.ExternalRelationships.ToList();
            int embedobj_count = embedded_objects.Count;
            int embedobj_number = 0;
            spreadsheet.Close();


            // If errors
            if (embedded_objects.Any())
            {
                // Inform user
                Console.WriteLine($"--> {embedobj_count} relationships detected");
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
                // Open spreadsheet to remove external relationships
                spreadsheet = SpreadsheetDocument.Open(filepath, true);
                //external_relationships.Remove(ExternalRelationship, extrel.Id);
                spreadsheet.Save();
                spreadsheet.Close();
                // Inform user
                //Console.WriteLine($"--> External relationship {extrel_number} removed");
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
                embedobj_message = "0 embedded objects relationships";

                return embedobj_message;
            }
        }
    }
}
