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
    public partial class Archive_Requirements
    {
        public bool Data { get; set; }

        public int Connections { get; set; }

        public int CellReferences { get; set; }

        public int RTDFunctions { get; set; }

        public int PrinterSettings { get; set; }

        public int ExternalObj { get; set; }

        public int EmbedObj { get; set; }

        public int Hyperlinks { get; set; }

        public bool ActiveSheet { get; set; }

        public bool AbsolutePath { get; set; }

        // Perform check of archival requirements
        public List<Archive_Requirements> Check_XLSX_Requirements(string filepath)
        {
            bool data = Check_Value(filepath);
            int connections = Check_DataConnections(filepath);
            int cellreferences = Check_CellReferences(filepath);
            int rtdfunctions = Check_RTDFunctions(filepath);
            int printersettings = Check_PrinterSettings(filepath);
            int extobjects = Check_ExternalObjects(filepath);
            int embedobj = Check_EmbeddedObjects(filepath);
            int hyperlinks = Check_Hyperlinks(filepath);
            bool activesheet = Check_ActiveSheet(filepath);
            //bool absolutepath = Check_AbsolutePath(filepath);

            // Add information to list and return it
            List<Archive_Requirements> Arc_Req = new List<Archive_Requirements>();
            Arc_Req.Add(new Archive_Requirements { Data = data, Connections = connections, CellReferences = cellreferences, RTDFunctions = rtdfunctions, PrinterSettings = printersettings, ExternalObj = extobjects, EmbedObj = embedobj, Hyperlinks = hyperlinks, ActiveSheet = activesheet });
            return Arc_Req;
        }

        // Check for any values by checking if sheets and cell values exist
        public bool Check_Value(string filepath)
        {
            bool hascellvalue = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                //Check if worksheets exist
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Sheets allSheets = wbPart.Workbook.Sheets;
                if (allSheets == null)
                {
                    Console.WriteLine("--> No cell values detected");
                    return hascellvalue;
                }
                // Check if any cells have any value
                foreach (Sheet aSheet in allSheets)
                {
                    WorksheetPart wsp = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(aSheet.Id);
                    Worksheet worksheet = wsp.Worksheet;
                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                    int row_count = rows.Count(); // Count number of rows
                    if (row_count > 0) // If any rows exist, this means cells exist
                    {
                        hascellvalue = true;
                        return hascellvalue;
                    }
                }
            }
            Console.WriteLine("--> No cell values detected");
            return hascellvalue;
        }

        // Check for data connections
        public int Check_DataConnections(string filepath)
        {
            int conn_count = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                ConnectionsPart conn = spreadsheet.WorkbookPart.ConnectionsPart;
                if (conn != null)
                {
                    conn_count = conn.Connections.Count();
                    Console.WriteLine($"--> {conn_count} data connections detected and removed");
                }
            }
            return conn_count;
        }

        // Check for external relationships
        public int Check_CellReferences(string filepath) // Find all external relationships
        {
            int cellreferences_count = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                List<ExternalWorkbookPart> extwbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts.ToList();
                if (extwbParts.Count > 0)
                {
                    foreach (ExternalWorkbookPart ext in extwbParts)
                    {
                        var elements = ext.ExternalLink.ChildElements.ToList();
                        foreach (var element in elements)
                        {
                            if (element.LocalName == "externalBook")
                            {
                                var externalLink = ext.ExternalLink.ToList();
                                foreach (ExternalBook externalBook in externalLink)
                                {
                                    var cellreferences = externalBook.SheetDataSet.ChildElements.ToList();
                                    foreach (var cellreference in cellreferences)
                                    {
                                        var cells = cellreference.InnerText.ToList();
                                        foreach (var cell in cells)
                                        {
                                            cellreferences_count++;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            if (cellreferences_count > 0)
            {
                Console.WriteLine($"--> {cellreferences_count} external cell references detected and removed");
            }
            return cellreferences_count;
        }

        // Check for external object references
        public int Check_ExternalObjects(string filepath)
        {
            int extobj_count = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                List<ExternalWorkbookPart> extwbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts.ToList();
                if (extwbParts.Count > 0)
                {
                    foreach (ExternalWorkbookPart ext in extwbParts)
                    {
                        var elements = ext.ExternalLink.ChildElements.ToList();
                        foreach (var element in elements)
                        {
                            if (element.LocalName == "oleLink")
                            {
                                var externalLink = ext.ExternalLink.ToList();
                                foreach (OleLink oleLink in externalLink)
                                {
                                    extobj_count++;
                                }
                            }
                        }
                    }
                }
            }
            if (extobj_count > 0)
            {
                Console.WriteLine($"--> {extobj_count} external objects detected and removed");
            }
            return extobj_count;
        }

        // Check for RTD functions
        public static int Check_RTDFunctions(string filepath) // Check for RTD functions
        {
            int rtd_functions_count = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                Sheets allSheets = wbPart.Workbook.Sheets;
                foreach (Sheet aSheet in allSheets)
                {
                    WorksheetPart wsp = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(aSheet.Id);
                    Worksheet worksheet = wsp.Worksheet;
                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                    foreach (var row in rows)
                    {
                        var cells = row.Elements<Cell>();
                        foreach (Cell cell in cells)
                        {
                            if (cell.CellFormula != null)
                            {
                                string formula = cell.CellFormula.InnerText;
                                if (formula.Length > 2)
                                {
                                    string hit = formula.Substring(0, 3); // Transfer first 3 characters to string
                                    if (hit == "RTD")
                                    {
                                        rtd_functions_count++;
                                        Console.WriteLine($"--> RTD function in sheet \"{aSheet.Name}\" cell {cell.CellReference} detected and removed");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return rtd_functions_count;
        }

        // Check for embedded objects
        public int Check_EmbeddedObjects(string filepath) // Check for embedded objects and return alert
        {
            int embedobj_count = 0;

            using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                var list = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (var item in list)
                {
                    int count_ole = item.EmbeddedObjectParts.Count(); // Register the number of OLE
                    int count_image = item.ImageParts.Count(); // Register number of images
                    int count_3d = item.Model3DReferenceRelationshipParts.Count(); // Register number of 3D models
                    embedobj_count = count_ole + count_image + count_3d; // Sum

                    if (embedobj_count > 0) // If embedded objects
                    {
                        Console.WriteLine($"--> {embedobj_count} embedded objects detected");
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
                    }
                }
            }
            return embedobj_count;
        }

        // Check for hyperlinks
        public int Check_Hyperlinks(string filepath) // Find all hyperlinks
        {
            int hyperlinks_count = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                List<HyperlinkRelationship> hyperlinks = spreadsheet
                    .GetAllParts()
                    .SelectMany(p => p.HyperlinkRelationships)
                    .ToList();
                hyperlinks_count = hyperlinks.Count;

                if (hyperlinks_count > 0) // If hyperlinks
                {
                    Console.WriteLine($"--> {hyperlinks_count} hyperlinks detected. Hyperlinks were not removed");
                    foreach (HyperlinkRelationship hyperlink in hyperlinks)
                    {
                        Console.WriteLine($"----> Hyperlink address: {hyperlink.Uri}");
                    }
                }
            }
            return hyperlinks_count;
        }

        // Check for printer settings
        public int Check_PrinterSettings(string filepath)
        {
            int printersettings_count = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                var worksheetpartslist = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                List<SpreadsheetPrinterSettingsPart> printerList = new List<SpreadsheetPrinterSettingsPart>();
                foreach (WorksheetPart worksheetpart in worksheetpartslist)
                {
                    printerList = worksheetpart.SpreadsheetPrinterSettingsParts.ToList();
                }
                foreach (SpreadsheetPrinterSettingsPart printer in printerList)
                {
                    printersettings_count++;
                }
                if (printerList.Count > 0)
                {
                    Console.WriteLine($"--> {printersettings_count} printersettings detected and removed");
                }
            }
            return printersettings_count;
        }

        // Check for active sheet
        public bool Check_ActiveSheet(string filepath)
        {
            bool activeSheet = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                BookViews bookViews = spreadsheet.WorkbookPart.Workbook.GetFirstChild<BookViews>();
                WorkbookView workbookView = bookViews.GetFirstChild<WorkbookView>();
                if (workbookView.ActiveTab != null)
                {
                    var activeSheetId = workbookView.ActiveTab.Value;
                    if (activeSheetId > 0)
                    {
                        Console.WriteLine("--> First sheet is not active sheet detected and changed");
                        activeSheet = true;
                    }
                }
            }
            return activeSheet;
        }

        // Check for absolute path
        public bool Check_AbsolutePath(string filepath)
        {
            bool absolutepath = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                if (spreadsheet.WorkbookPart.Workbook.AbsolutePath.Url != null)
                {
                    Console.WriteLine("--> Absolute path to local directory detected and removed");
                    absolutepath = true;
                }
            }
            return absolutepath;
        }
    }
}