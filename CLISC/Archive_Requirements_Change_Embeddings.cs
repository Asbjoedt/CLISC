using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Vml;
using ImageMagick;

namespace CLISC
{
    public partial class Archive_Requirements
    {
        // Convert embedded objects
        public int Convert_EmbeddedObjects(string filepath)
        {
            int success = 0;
            int binary_fail = 0;
            int threeD_fail = 0;
            List<EmbeddedObjectPart> ole = new List<EmbeddedObjectPart>();
            List<EmbeddedPackagePart> packages = new List<EmbeddedPackagePart>();
            List<ImagePart> emf = new List<ImagePart>();
            List<ImagePart> images = new List<ImagePart>();
            List<Model3DReferenceRelationshipPart> threeD = new List<Model3DReferenceRelationshipPart>();

            // Open spreadsheet
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart worksheetPart in worksheetParts)
                {
                    // Perform check
                    ole = worksheetPart.EmbeddedObjectParts.Distinct().ToList();
                    packages = worksheetPart.EmbeddedPackageParts.Distinct().ToList();
                    emf = worksheetPart.ImageParts.Distinct().ToList();
                    if (worksheetPart.DrawingsPart != null) // DrawingsPart needs a null check
                    {
                        images = worksheetPart.DrawingsPart.ImageParts.Distinct().ToList();
                    }
                    threeD = worksheetPart.Model3DReferenceRelationshipParts.Distinct().ToList();

                    // Perform change

                    // Embedded binaries cannot be processed
                    foreach (EmbeddedObjectPart part in ole)
                    {
                        // Extract object
                        string output_filepath = Create_Output_Filepath(filepath, part.Uri.ToString());
                        Stream input_stream = part.GetStream();
                        Extract_EmbeddedObjects(input_stream, output_filepath);

                        // Register conversion fail
                        binary_fail++;
                    }
                    if (binary_fail > 0)
                    {
                        // Inform user
                        Console.WriteLine($"--> Change: {binary_fail} embedded binary files cannot be processed");
                    }

                    // Convert embedded packages to OpenDocument
                    foreach (EmbeddedPackagePart part in packages)
                    {
                        Convert_EmbedPackage(filepath, worksheetPart, part);
                        success++;
                    }

                    // 3D objects cannot be processed - Bug in Open XML SDK?
                    foreach (Model3DReferenceRelationshipPart part in threeD)
                    {
                        // Extract object
                        string output_filepath = Create_Output_Filepath(filepath, part.Uri.ToString());
                        Stream input_stream = part.GetStream();
                        Extract_EmbeddedObjects(input_stream, output_filepath);

                        // Register conversion fail
                        threeD_fail++;
                    }
                    if (threeD_fail > 0)
                    {
                        // Inform user
                        Console.WriteLine($"--> Change: {threeD_fail} embedded model 3D reference relationships cannot be processed");
                    }

                    // Convert Excel-generated .emf images to TIFF
                    foreach (ImagePart part in emf)
                    {
                        try
                        {
                            Convert_EmbedEmf(filepath, worksheetPart, part);
                            success++;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message); // DELETE THIS LINE LATER
                        }
                    }

                    // Convert embedded images to TIFF
                    foreach (ImagePart part in images)
                    {
                        Convert_EmbedImg(filepath, worksheetPart, part);
                        success++;
                    }
                }
            }
            return success;
        }

        // Convert embedded images to TIFF
        public void Convert_EmbedImg(string filepath, WorksheetPart worksheetPart, ImagePart part)
        {
            // Convert streamed image to new stream
            Stream stream = part.GetStream();
            Stream new_Stream = Convert_EmbedObj_ImageMagick(stream);
            stream.Dispose();

            // Extract converted image to folder
            string extract_filepath = Create_Output_Filepath(filepath, part.Uri.ToString());
            Extract_EmbeddedObjects(new_Stream, extract_filepath);

            // Add new ImagePart
            ImagePart new_ImagePart = worksheetPart.DrawingsPart.AddImagePart(ImagePartType.Tiff);

            // Save image from stream to new ImagePart
            new_Stream.Position = 0;
            new_ImagePart.FeedData(new_Stream);

            // Change relationships of image
            string id = Get_RelationshipId(part);
            Blip blip = worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>()
                            .Where(p => p.BlipFill.Blip.Embed == id)
                            .Select(p => p.BlipFill.Blip)
                            .Single();
            blip.Embed = Get_RelationshipId(new_ImagePart);

            // Delete original ImagePart
            worksheetPart.DrawingsPart.DeletePart(part);
        }

        // Convert Excel-generated .emf images to TIFF
        public void Convert_EmbedEmf(string filepath, WorksheetPart worksheetPart, ImagePart part)
        {
            // Convert streamed image to new stream
            Stream stream = part.GetStream();
            Stream new_Stream = Convert_EmbedObj_ImageMagick(stream);
            stream.Dispose();

            // Extract converted image to folder
            string extract_filepath = Create_Output_Filepath(filepath, part.Uri.ToString());
            Extract_EmbeddedObjects(new_Stream, extract_filepath);

            // Add new ImagePart
            ImagePart new_ImagePart = worksheetPart.VmlDrawingParts.First().AddImagePart(ImagePartType.Tiff);

            // Save image from stream to new ImagePart
            new_Stream.Position = 0;
            new_ImagePart.FeedData(new_Stream);

            // Change relationships of image
            string id = Get_RelationshipId(part);
            ImageData imageData = worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<ImageData>()
                            .Where(p => p.RelId == id)
                            .Select(p => p)
                            .Single();
            imageData.RelId = Get_RelationshipId(new_ImagePart);

            // Delete original ImagePart
            worksheetPart.VmlDrawingParts.First().DeletePart(part);
        }

        // Convert embedded packages to OpenDocument
        public void Convert_EmbedPackage(string filepath, WorksheetPart worksheetPart, EmbeddedPackagePart part)
        {
            // Extract EmbeddedPackage to folder
            string extract_filepath = Create_Output_Filepath(filepath, part.Uri.ToString());
            Stream stream = part.GetStream();
            Extract_EmbeddedObjects(stream, extract_filepath);
            stream.Dispose();

            // If not OpenDocument package, then convert to OpenDocument
            string extension = System.IO.Path.GetExtension(extract_filepath);
            if (extension != ".ods" || extension != ".fods" || extension != ".ots" || extension != ".fodt" || extension != ".odt" || extension != ".ott" || extension != ".fodp" || extension != ".odp" || extension != ".otp")
            {
                // Convert to OpenDocument
                string new_Extension = Create_OpenDocument_Extension(extract_filepath);
                Uri new_Uri = Create_Uri(part.Uri, new_Extension);
                string new_Filename = Create_Filename(new_Uri);
                string output_folder = System.IO.Path.GetDirectoryName(extract_filepath);
                string output_filepath = output_folder + "\\" + new_Filename;
                Convert_LibreOffice(extract_filepath, output_filepath);

                // Create new EmbeddedPackage
                string contentType = Create_ContentType(extract_filepath);
                EmbeddedPackagePart new_EmbeddedPackagePart = worksheetPart.AddEmbeddedPackagePart(contentType);

                // Feed converted data to the new package
                var new_Stream = new MemoryStream();
                using (FileStream file = new FileStream(output_filepath, FileMode.Open, FileAccess.Read))
                {
                    file.CopyTo(new_Stream);
                }
                new_EmbeddedPackagePart.FeedData(new_Stream);
                new_Stream.Dispose();

                // Change relationships
                Change_EmbedPackage_Relationships(worksheetPart, part.Uri);

                // Delete original EmbeddedPackage
                worksheetPart.DeletePart(part);
            }
        }

        // Extract embedded objects
        public void Extract_EmbeddedObjects(Stream input_stream, string output_filepath)
        {
            // Extract embedded object to folder
            using (FileStream fileStream = File.Create(output_filepath))
            {
                input_stream.Seek(0, SeekOrigin.Begin);
                input_stream.CopyTo(fileStream);
            }
        }

        // Convert embedded object to TIFF using ImageMagick
        public Stream Convert_EmbedObj_ImageMagick(Stream stream)
        {
            // Read the input stream in ImageMagick
            using (MagickImage image = new MagickImage(stream))
            {
                // Set input stream position to beginning
                stream.Position = 0;

                // Create a memorystream to write image to
                MemoryStream new_stream = new MemoryStream();

                // Adjust TIFF settings
                image.Format = MagickFormat.Tiff;
                image.Settings.ColorSpace = ColorSpace.RGB;
                image.Settings.Depth = 32;
                image.Settings.Compression = CompressionMethod.LZW;

                // Write image to stream
                image.Write(new_stream);

                // Return the memorystream
                return new_stream;
            }
        }

        // Convert embedded package to OpenDocument using LibreOffice
        public void Convert_LibreOffice(string input_filepath, string output_filepath)
        {
            string format = System.IO.Path.GetFullPath(output_filepath).Split(".").Last();
            string output_folder = System.IO.Path.GetDirectoryName(output_filepath);

            // Use LibreOffice command line for conversion
            Process app = new Process();
            string? dir = null;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // If app is run on Windows
            {
                dir = Environment.GetEnvironmentVariable("LibreOffice");
            }
            if (dir != null)
            {
                app.StartInfo.FileName = dir;
            }
            else
            {
                app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
            }
            app.StartInfo.Arguments = $"--headless --convert-to {format} \"{input_filepath}\" --outdir \"{output_folder}\"";
            app.Start();
            app.WaitForExit();
            app.Close();

            // Delete the original extracted file
            File.Delete(input_filepath);
        }

        // Get relationship id of an OpenXmlPart
        public string Get_RelationshipId(OpenXmlPart part)
        {
            string id = "";
            IEnumerable<OpenXmlPart> parentParts = part.GetParentParts();
            foreach (OpenXmlPart parentPart in parentParts)
            {
                if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.DrawingsPart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.VmlDrawingPart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.Model3DReferenceRelationshipPart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.OleObjectPart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
            }
            return id;
        }

        // Create new Uri with right extension for embedded object
        public Uri Create_Uri(Uri part_Uri, string new_Extension)
        {
            string input_path = part_Uri.ToString();
            int dot = input_path.LastIndexOf(".");
            string output_path = input_path.Substring(0, dot) + new_Extension;
            Uri new_uri = new Uri(output_path, UriKind.Relative);
            return new_uri;
        }

        // Create new filename with right extension for embedded object
        public string Create_Filename(Uri new_Uri)
        {
            string filename = new_Uri.ToString().Split("/").Last();
            return filename;
        }

        // Create a new output filepath for extracted embedded files
        public string Create_Output_Filepath(string filepath, string uri)
        {
            // Create new folder for embedded objects
            int backslash = filepath.LastIndexOf("\\");
            string file_folder = filepath.Substring(0, backslash);
            string new_folder = file_folder + "\\Embedded objects";
            Directory.CreateDirectory(new_folder);
         
            // Create and return output filepath
            string output_filepath = new_folder + "\\" + uri.Split("/").Last();
            return output_filepath;
        }

        // Transform extension to OpenDocument extension
        public string Create_OpenDocument_Extension(string filepath)
        {
            string extract_Extension = System.IO.Path.GetExtension(filepath).ToLower();
            string? new_Extension = null;
            switch (extract_Extension)
            {
                case ".xlsx":
                case ".xltm":
                case ".xlsm":
                case ".xltx":
                case ".xlam":
                case ".xls":
                case ".xlt":
                case ".xla":
                case ".numbers":
                    new_Extension = ".ods";
                    return new_Extension;

                case ".docx":
                case ".dotx":
                case ".dotm":
                case ".docm":
                case ".doc":
                case ".dot":
                    new_Extension = ".odt";
                    return new_Extension;

                case ".pptx":
                case ".pptm":
                case ".ppsm":
                case ".ppsx":
                case ".potm":
                case ".potx":
                case ".ppam":
                case ".ppt":
                case ".ppa":
                case ".pot":
                case ".pps":
                    new_Extension = ".odp";
                    return new_Extension;

                default:
                    throw new Exception();
            }
        }

        // Translate an OpenDocument content type from extension
        public string Create_ContentType(string filepath)
        {
            string extract_Extension = System.IO.Path.GetExtension(filepath).ToLower();
            string? contentType = null;
            switch (extract_Extension)
            {
                case ".xlsx":
                case ".xltm":
                case ".xlsm":
                case ".xltx":
                case ".xlam":
                case ".xls":
                case ".xlt":
                case ".xla":
                case ".numbers":
                    contentType = "application/vnd.oasis.opendocument.spreadsheet";
                    return contentType;

                case ".docx":
                case ".dotx":
                case ".dotm":
                case ".docm":
                case ".doc":
                case ".dot":
                    contentType = "application/vnd.oasis.opendocument.text";
                    return contentType;

                case ".pptx":
                case ".pptm":
                case ".ppsm":
                case ".ppsx":
                case ".potm":
                case ".potx":
                case ".ppam":
                case ".ppt":
                case ".ppa":
                case ".pot":
                case ".pps":
                    contentType = "application/vnd.oasis.opendocument.presentation";
                    return contentType;

                default:
                    throw new Exception();
            }
        }

        // Change relationships to new embedded package
        public void Change_EmbedPackage_Relationships(WorksheetPart worksheetPart, Uri original_Uri)
        {
            IEnumerable<IdPartPair> parts = worksheetPart.Parts;
            foreach (IdPartPair id in parts) 
            {

            }

            IEnumerable<EmbeddedPackagePart> embedpacks = worksheetPart.EmbeddedPackageParts;
            foreach (EmbeddedPackagePart part in embedpacks)
            {
                if (part.Uri == original_Uri)
                {

                }
            }
        }

        // Alternative approach
        public System.IO.Packaging.PackagePart Alternative_Create_ImagePart(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, Uri new_Uri, string output_filepath)
        {
            // https://stackoverflow.com/questions/18569113/open-xml-sdk-addimagepart-change-image-location-from-media-to-word-media

            // Create new part
            System.IO.Packaging.PackagePart packageImagePart = spreadsheet.Package.CreatePart(new_Uri, "image/tiff");

            // Feed data
            byte[] imageBytes = File.ReadAllBytes(output_filepath);
            packageImagePart.GetStream().Write(imageBytes, 0, imageBytes.Length);

            // Create relationships
            System.IO.Packaging.PackagePart worksheetPackagePart = spreadsheet.WorkbookPart.OpenXmlPackage.Package.GetPart(worksheetPart.Uri);
            Console.WriteLine(worksheetPackagePart.Uri);
            System.IO.Packaging.PackageRelationship imageReleationshipPart = worksheetPackagePart.CreateRelationship(new_Uri, System.IO.Packaging.TargetMode.Internal, "http://purl.oclc.org/ooxml/officeDocument/relationships/image");

            return packageImagePart;
        }
    }
}
