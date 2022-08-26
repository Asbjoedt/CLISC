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
        public Tuple<bool, int, int, int, int, int, int> Check_XLSX_Requirements(string filepath) 
        {
            bool data = Check_Value(filepath);
            int connections = Check_DataConnections(filepath);
            int extrels = Check_ExternalRelationships(filepath);
            int rtdfunctions = Check_RTDFunctions(filepath);
            int printersettings = Check_PrinterSettings(filepath);
            int embedobj = Check_EmbeddedObjects(filepath);
            int hyperlinks = Check_Hyperlinks(filepath);

            (bool, int, int, int, int, int, int) pidgeon = (data, connections, extrels, rtdfunctions, printersettings, embedobj, hyperlinks);
            return pidgeon.ToTuple();
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
                    DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = wsp.Worksheet;
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
        public int Check_ExternalRelationships(string filepath) // Find all external relationships
        {
            int extrels_count = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                List<ExternalRelationship> extrels = spreadsheet 
                .GetAllParts()
                .SelectMany(p => p.ExternalRelationships)
                .ToList();
                extrels_count = extrels.Count;

                if (extrels.Count > 0) // If external relationships
                {
                    int extrel_number = 0;
                    Console.WriteLine($"--> {extrels.Count} external relationships detected");
                    foreach (ExternalRelationship rel in extrels)
                    {
                        extrel_number++;
                        Console.WriteLine($"--> External relationship {extrel_number}");
                        if (rel.RelationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" || rel.RelationshipType == "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject")
                        {
                            Console.WriteLine("--> External OLE object detected. Handle OLE object manually");
                        }
                        Console.WriteLine($"----> ID: {rel.Id}");
                        Console.WriteLine($"----> Target URI: {rel.Uri}");
                        Console.WriteLine($"----> Relationship type: {rel.RelationshipType}");
                        Console.WriteLine($"----> External: {rel.IsExternal}");
                    }
                    return extrels_count;
                }
                return extrels_count;
            }
        }

        public static int Check_RTDFunctions(string filepath) // Check for RTD functions
        {
            int rtd_functions = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                WorkbookPart wbPart = spreadsheet.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Sheets allSheets = wbPart.Workbook.Sheets;
                foreach (Sheet aSheet in allSheets)
                {
                    WorksheetPart wsp = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(aSheet.Id);
                    DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = wsp.Worksheet;
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
                                        rtd_functions++;
                                        Console.WriteLine($"--> RTD function in sheet \"{aSheet.Name}\" cell {cell.CellReference} detected and removed");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return rtd_functions;
        }

        public int Check_EmbeddedObjects(string filepath) // Check for embedded objects and return alert
        {
            int count_embedobj = 0;

            using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                var list = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (var item in list)
                {
                    int count_ole = item.EmbeddedObjectParts.Count(); // Register the number of OLE
                    int count_image = item.ImageParts.Count(); // Register number of images
                    int count_3d = item.Model3DReferenceRelationshipParts.Count(); // Register number of 3D models
                    count_embedobj = count_ole + count_image + count_3d; // Sum

                    if (count_embedobj > 0) // If embedded objects
                    {
                        Console.WriteLine($"--> {count_embedobj} embedded objects detected");
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
            return count_embedobj;
        }

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
                    Console.WriteLine($"--> {hyperlinks_count} hyperlinks detected");
                    int hyperlink_number = 0;
                    foreach (HyperlinkRelationship hyperlink in hyperlinks)
                    {
                        hyperlink_number++;
                        Console.WriteLine($"--> Hyperlink: {hyperlink_number}");
                        Console.WriteLine($"----> Address: {hyperlink.Uri}");
                    }
                }
            }
            return hyperlinks_count;
        }

        // Check for printer settings
        public int Check_PrinterSettings(string filepath)
        {
            int printersettings = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                var list = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (var item in list)
                {
                    var printerList = item.SpreadsheetPrinterSettingsParts.ToList();
                    if (printerList.Count > 0)
                    {
                        Console.WriteLine("--> Printersettings detected");
                    }
                    foreach (var part in printerList)
                    {
                        printersettings++;
                    }
                }
            }
            return printersettings;
        }
    }
}