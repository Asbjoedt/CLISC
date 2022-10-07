using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CLISC;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CLISC
{
    public class Repair
    {
        public void Repair_OOXML(string filepath)
        {
            Repair_VBA(filepath);
        }

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
    }
}