using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CLISC
{
    public class Repair
    {
        public void Repair_OOXML(string filepath)
        {
            bool repair_1 = Repair_VBA(filepath);
            bool repair_2 = Repair_DefinedNames(filepath);

            // If any repair method has been performed
            if (repair_1 == true || repair_2 == true)
            {
                Console.WriteLine("--> Spreadsheet was repaired");
            }
        }

        // Repair spreadsheets that had VBA code (macros) in them
        public bool Repair_VBA(string filepath)
        {
            bool repaired = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                // Remove VBA project (if present) due to error in Open XML SDK
                VbaProjectPart vba = spreadsheet.WorkbookPart.VbaProjectPart;
                if (vba != null)
                {
                    spreadsheet.WorkbookPart.DeletePart(vba);
                    repaired = true;
                }

                // Correct the namespace for customUI14.xml, if wrong
                //WORK IN PROGRESS
                RibbonExtensibilityPart ribbon = spreadsheet.RibbonExtensibilityPart;
                if (ribbon != null)
                {
                    Uri uri = new Uri("/customUI/customUI14.xml", UriKind.Relative);
                    if (spreadsheet.Package.GetPart(uri) != null)
                    {
                        if (ribbon.RootElement.NamespaceUri != "http://schemas.microsoft.com/office/2009/07/customui")
                        {
                            var list = ribbon.RootElement.NamespaceDeclarations.ToList();
                            foreach (var name in list)
                            {
                                Console.WriteLine(name.Key + " " + name.Value);
                            }
                            Console.WriteLine(ribbon.RootElement.Prefix);
                            Console.WriteLine(ribbon.RootElement.NamespaceUri);
                        }
                    }
                }
            }
            return repaired;
        }

        // Repair invalid defined names
        public bool Repair_DefinedNames(string filepath)
        {
            bool repaired = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                DefinedNames definedNames = spreadsheet.WorkbookPart.Workbook.DefinedNames;

                // Remove legacy Excel 4.0 GET.CELL function (if present)
                if (definedNames != null)
                {
                    var definedNamesList = definedNames.ToList();
                    foreach (DefinedName definedName in definedNamesList)
                    {
                        if (definedName.InnerXml.Contains("GET.CELL"))
                        {
                            definedName.Remove();
                            repaired = true;
                        }
                    }
                }

                // Remove defined names with these " " (3 characters) in reference
                if (definedNames != null)
                {
                    var definedNamesList = definedNames.ToList();
                    foreach (DefinedName definedName in definedNamesList)
                    {
                        if (definedName.InnerXml.Contains("\" \""))
                        {
                            definedName.Remove();
                            repaired = true;
                        }
                    }
                }
            }
            return repaired;
        }

        // Delete query tables if query tables exists without relationships
        // WORK IN PROGRESS
        public bool Repair_QueryTables(string filepath)
        {
            bool repaired = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                for (int i = 0; i < 20; i++)
                {
                    Uri queryUri = new Uri($"/xl/queryTables/queryTable{i}.xml", UriKind.Relative);
                    if (spreadsheet.Package.PartExists(queryUri) == true)
                    {
                        // Delete query tables up until 20
                        spreadsheet.Package.DeletePart(queryUri);

                        // Delete query table relationships in tables
                        List<WorksheetPart> wsParts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                        foreach (WorksheetPart wsPart in wsParts)
                        {
                            var partsList = wsPart.Parts.ToList();
                            foreach (var part in partsList)
                            {
                                if (part.OpenXmlPart.Equals("DocumentFormat.OpenXml.Packaging.TableDefinitionPart"))
                                {
                                    string id = part.RelationshipId;
                                    wsPart.DeleteReferenceRelationship(id);
                                }

                                Console.WriteLine(part.OpenXmlPart);
                            }
                        }
                    }
                }
            }
            return repaired;
        }
    }
}