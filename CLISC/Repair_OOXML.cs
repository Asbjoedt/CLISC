using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CLISC
{
    public class Repair
    {
        public void Repair_OOXML(string filepath)
        {
            bool repair_1 = Repair_VBA(filepath);
            bool repair_2 = Repair_CustomUI(filepath);
            bool repair_3 = Repair_DefinedNames(filepath);

            // If any repair method has been performed
            if (repair_1 || repair_2 || repair_3)
            {
                Console.WriteLine("--> Repair: Spreadsheet was repaired");
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
            }
            return repaired;
        }

        //WORK IN PROGRESS
        // Repair spreadsheets that has CustomUI in them
        public bool Repair_CustomUI(string filepath)
        {
            bool repaired = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                RibbonExtensibilityPart ribbon = spreadsheet.RibbonExtensibilityPart;
                if (ribbon != null)
                {
                    // Correct the namespace for CustomUI XML files, if wrong
                    if (ribbon.CustomUI.NamespaceUri == "http://schemas.microsoft.com/office/2006/01/customui")
                    {
                        //ribbon.CustomUI.RemoveNamespaceDeclaration("mso");
                        //ribbon.CustomUI.AddNamespaceDeclaration("mso", "http://schemas.microsoft.com/office/2009/07/customui");
                        //repaired = true;
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
    }
}