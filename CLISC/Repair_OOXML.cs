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
            Repair_VBA(filepath);
            Repair_DataConnections(filepath);
        }

        // Repair spreadsheets that had VBA code (macros) in them
        public void Repair_VBA(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                // Remove VBA project (if present) due to error in Open XML SDK
                VbaProjectPart vba = spreadsheet.WorkbookPart.VbaProjectPart;
                if (vba != null)
                {
                    spreadsheet.WorkbookPart.DeletePart(vba);
                }

                // Remove Excel 4.0 GET.CELL function (if present) due to error in Open XML SDK
                DefinedNames definedNames = spreadsheet.WorkbookPart.Workbook.DefinedNames;
                if (definedNames != null)
                {
                    var definedNamesList = definedNames.ToList();
                    foreach (DefinedName definedName in definedNamesList)
                    {
                        if (definedName.InnerXml.Contains("GET.CELL"))
                        {
                            definedName.Remove();
                        }
                    }
                }

                // Correct the namespace for customUI14.xml, if wrong
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
        }

        // Delete query tables if query tables exists without relationships
        public void Repair_QueryTables(string filepath)
        {
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
        }

        public void Repair_DataConnections(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                // If spreadsheet contains a custom XML Map, delete databinding
                if (spreadsheet.WorkbookPart.CustomXmlMappingsPart != null)
                {
                    CustomXmlMappingsPart xmlMap = spreadsheet.WorkbookPart.CustomXmlMappingsPart;
                    List<Map> maps = xmlMap.MapInfo.Elements<Map>().ToList();
                    foreach (Map map in maps)
                    {
                        if (map.DataBinding != null)
                        {
                            map.DataBinding.Remove();
                        }
                    }
                }
            }
        }
    }
}