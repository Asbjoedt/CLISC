using CLISC;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                if (spreadsheet.WorkbookPart.Workbook.DefinedNames != null)
                {
                    var definednames = spreadsheet.WorkbookPart.Workbook.DefinedNames.ToList();
                    foreach (DocumentFormat.OpenXml.Spreadsheet.DefinedName definedname in definednames)
                    {
                        if (definedname.InnerXml.Contains("GET.CELL"))
                        {
                            definedname.Remove();
                        }
                    }
                }

                // Correct the namespace for customUI14.xml
                if (spreadsheet.RibbonExtensibilityPart != null)
                {
                    Uri uri = new Uri("/customUI/customUI14.xml", UriKind.Relative);
                    if (spreadsheet.Package.PartExists(uri) == true)
                    {
                        if (spreadsheet.RibbonExtensibilityPart.CustomUI.NamespaceUri != "http://schemas.microsoft.com/office/2009/07/customui:customUI")
                        {
                            spreadsheet.RibbonExtensibilityPart.CustomUI.RemoveNamespaceDeclaration("x");
                            spreadsheet.RibbonExtensibilityPart.CustomUI.AddNamespaceDeclaration("x", "http://schemas.microsoft.com/office/2009/07/customui:customUI");
                        }
                    }
                }
            }
        }
    }
}