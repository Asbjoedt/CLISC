using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2013.ExcelAc;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;

namespace CLISC
{
    public partial class Archive_Requirements
    {
        public bool Data { get; set; }

        public bool Conformance { get; set; }

        public int Connections { get; set; }

        public int CellReferences { get; set; }

        public int RTDFunctions { get; set; }

        public int PrinterSettings { get; set; }

        public int ExternalObj { get; set; }

        public int EmbedObj { get; set; }

        public int Hyperlinks { get; set; }

        public bool ActiveSheet { get; set; }

        public bool AbsolutePath { get; set; }

        public bool Metadata { get; set; }

        // Perform check of archival requirements
        public List<Archive_Requirements> Check_XLSX_Requirements(string filepath)
        {
            bool data = Check_Value(filepath);
            bool conformance = Check_Conformance(filepath);
            bool metadata = Check_Metadata(filepath);
            int connections = Check_DataConnections(filepath);
            int cellreferences = Check_CellReferences(filepath);
            int rtdfunctions = Check_RTDFunctions(filepath);
            int printersettings = Check_PrinterSettings(filepath);
            int extobjects = Check_ExternalObjects(filepath);
            int embedobj = Check_EmbeddedObjects(filepath);
            int hyperlinks = Check_Hyperlinks(filepath);
            bool activesheet = Check_ActiveSheet(filepath);
            bool absolutepath = Check_AbsolutePath(filepath);

            // Add information to list and return it
            List<Archive_Requirements> Arc_Req = new List<Archive_Requirements>();
            Arc_Req.Add(new Archive_Requirements { Data = data, Conformance = conformance, Connections = connections, CellReferences = cellreferences, RTDFunctions = rtdfunctions, PrinterSettings = printersettings, ExternalObj = extobjects, EmbedObj = embedobj, Hyperlinks = hyperlinks, ActiveSheet = activesheet, AbsolutePath = absolutepath, Metadata = metadata });
            return Arc_Req;
        }

        // Check for any values by checking if sheets and cell values exist
        public bool Check_Value(string filepath)
        {
            bool hascellvalue = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                //Check if worksheets exist
                Sheets allSheets = spreadsheet.WorkbookPart.Workbook.Sheets;
                if (allSheets == null)
                {
                    Console.WriteLine("--> No cell values detected");
                    return hascellvalue;
                }

                // Check if any cells have any value
                List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart part in worksheetparts)
                {
                    Worksheet worksheet = part.Worksheet;
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

        // Check for Strict conformance
        public bool Check_Conformance(string filepath)
        {
            bool conformance = false;

            // Perform check
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                Workbook workbook = spreadsheet.WorkbookPart.Workbook;
                if (workbook.Conformance == null || workbook.Conformance != "strict")
                {
                    Console.WriteLine("--> Transitional conformance identified - Spreadsheet was converted to Strict conformance");
                }
                else if (workbook.Conformance == "strict")
                {
                    Console.WriteLine("--> Strict conformance identified");
                    conformance = true;
                }
            }
            return conformance;
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

                    // Write information to metadata file
                    string folder = Path.GetDirectoryName(filepath);
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine("DATA CONNECTIONS");
                        w.WriteLine("---");
                    }
                    foreach (Connection connection in conn.Connections)
                    {
                        using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                        {
                            w.WriteLine($"NAME: {connection.Name}; DESCRIPTION: {connection.Description}; CREDENTIALS: {connection.Credentials}; CONNECTION FILE: {connection.ConnectionFile}");
                        }
                    }
                }
            }
            return conn_count;
        }

        // Check for external cell references
        public int Check_CellReferences(string filepath)
        {
            int cellreferences_count = 0;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                // Delete all cell references in worksheet
                List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart part in worksheetparts)
                {
                    Worksheet worksheet = part.Worksheet;
                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                    foreach (var row in rows)
                    {
                        var cells = row.Elements<Cell>();
                        foreach (Cell cell in cells)
                        {
                            if (cell.CellFormula != null)
                            {
                                string formula = cell.CellFormula.InnerText;
                                if (formula.Length > 0)
                                {
                                    string hit = formula.Substring(0, 1); // Transfer first 1 characters to string
                                    if (hit == "[")
                                    {
                                        cellreferences_count++;
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
                List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart part in worksheetparts)
                {
                    Worksheet worksheet = part.Worksheet;
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
                                        Console.WriteLine($"--> RTD function in sheet {part.Uri} cell {cell.CellReference} detected and removed");
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
            int embedobj_number = 0;
            int count_ole = 0;
            int count_package = 0;
            int count_image = 0;
            int count_drawing_image = 0;
            int count_3d = 0;

            using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                List<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart worksheetPart in worksheetParts)
                {
                    count_ole = worksheetPart.EmbeddedObjectParts.Count(); // Register the number of OLE
                    count_package = worksheetPart.EmbeddedPackageParts.Count(); // Register the number of packages
                    count_image = worksheetPart.ImageParts.Count(); // Register number of images
                    if (worksheetPart.DrawingsPart != null)
                    {
                        count_drawing_image = worksheetPart.DrawingsPart.ImageParts.Count();
                        count_image = count_image + count_drawing_image;
                    }
                    count_3d = worksheetPart.Model3DReferenceRelationshipParts.Count(); // Register number of 3D models
                    embedobj_count = count_ole + count_package + count_image + count_3d; // Sum

                    if (embedobj_count > 0) // If embedded objects
                    {
                        Console.WriteLine($"--> {embedobj_count} embedded objects detected");

                        List<EmbeddedObjectPart> embed_ole = worksheetPart.EmbeddedObjectParts.ToList();
                        List<EmbeddedPackagePart> embedobj_package = worksheetPart.EmbeddedPackageParts.ToList();
                        List<ImagePart> embed_image = worksheetPart.ImageParts.ToList();
                        List<ImagePart> embed_drawing_image = new List<ImagePart>();
                        if (count_drawing_image > 0)
                        {
                            embed_drawing_image = worksheetPart.DrawingsPart.ImageParts.ToList();
                        }
                        List<Model3DReferenceRelationshipPart> embed_3d = worksheetPart.Model3DReferenceRelationshipParts.ToList();

                        foreach (EmbeddedObjectPart part in embed_ole) // Inform user of each OLE object
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: OLE object");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");

                        }
                        foreach (EmbeddedPackagePart part in embedobj_package) // Inform user of each package object
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: Package object");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                        }
                        foreach (ImagePart part in embed_image) // Inform user of each .emf image object
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: Image object");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                        }
                        foreach (ImagePart part in embed_drawing_image) // Inform user of each image object
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: Image object");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                        }
                        foreach (Model3DReferenceRelationshipPart part in embed_3d) // Inform user of each 3D object
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: 3D model object");
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
                List<WorksheetPart> worksheetpartslist = spreadsheet.WorkbookPart.WorksheetParts.ToList();
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
                    Console.WriteLine($"--> {printersettings_count} printer settings detected and removed");
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
                if (spreadsheet.WorkbookPart.Workbook.BookViews != null)
                {
                    BookViews bookViews = spreadsheet.WorkbookPart.Workbook.BookViews;
                    if (bookViews.ChildElements.Where(p => p.OuterXml == "workbookView") != null)
                    {
                        WorkbookView workbookView = bookViews.GetFirstChild<WorkbookView>();
                        if (workbookView.ActiveTab != null)
                        {
                            var activeSheetId = workbookView.ActiveTab.Value;
                            if (activeSheetId > 0)
                            {
                                Console.WriteLine("--> First sheet is not active detected and changed");
                                activeSheet = true;
                            }
                        }
                    }
                }
            }
            return activeSheet;
        }

        // Check for absolute path
        public bool Check_AbsolutePath(string filepath)
        {
            bool absolutepath = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                if (spreadsheet.WorkbookPart.Workbook.AbsolutePath != null)
                {
                    Console.WriteLine("--> Absolute path to local directory detected and removed");
                    absolutepath = true;
                }
            }
            return absolutepath;
        }

        // Check for VBA projects
        public bool Check_VBA(string filepath)
        {
            bool vba = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                VbaProjectPart found = spreadsheet.WorkbookPart.VbaProjectPart;
                if (found != null)
                {
                    Console.WriteLine("--> VBA project found and removed");
                    vba = true;
                }
            }
            return vba;
        }

        // Check for metadata in file properties
        public bool Check_Metadata(string filepath)
        {
            bool metadata = false;
            string folder = Path.GetDirectoryName(filepath);

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                PackageProperties property = spreadsheet.Package.PackageProperties;

                // Write information to metadata file
                using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                {
                    w.WriteLine("FILE PROPERTIES METADATA");
                    w.WriteLine("---");
                }

                if (property.Creator != null)
                {
                    metadata = true;

                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"CREATOR: {property.Creator}");
                    }
                }
                if (property.Title != null)
                {
                    metadata = true;

                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"TITLE: {property.Title}");
                    }
                }
                if (property.Subject != null)
                {
                    metadata = true;

                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"SUBJECT: {property.Subject}");
                    }
                }
                if (property.Description != null)
                {
                    metadata = true;

                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"DESCRIPTION: {property.Description}");
                    }
                }
                if (property.Keywords != null)
                {
                    metadata = true;

                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"KEYWORDS: {property.Keywords}");
                    }
                }
                if (property.Category != null)
                {
                    metadata = true;

                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"CATEGORY: {property.Category}");
                    }
                }
                if (property.LastModifiedBy != null)
                {
                    metadata = true;

                    // Write information to metadata file
                    using (StreamWriter w = File.AppendText($"{folder}\\orgFile_metadata.txt"))
                    {
                        w.WriteLine($"LAST MODIFIED BY: {property.LastModifiedBy}");
                    }
                }
            }
            return metadata;
        }
    }
}