using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Linq;
using ImageMagick;

namespace CLISC
{
    public partial class Archive_Requirements
    {
        // Convert embedded objects to TIFF using ImageMagick
        public void Convert_EmbeddedObjects(string filepath)
        {
            // Create lists
            List<EmbeddedObjectPart> ole = new List<EmbeddedObjectPart>();
            List<EmbeddedPackagePart> packages = new List<EmbeddedPackagePart>();
            List<ImagePart> emf = new List<ImagePart>();
            List<ImagePart> images = new List<ImagePart>();
            List<Model3DReferenceRelationshipPart> threeD = new List<Model3DReferenceRelationshipPart>();

            // Open spreadsheet
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                WorkbookPart workbook = spreadsheet.WorkbookPart;
                IEnumerable<WorksheetPart> worksheetParts = workbook.WorksheetParts;
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

                    // Convert embedded objects
                    foreach (EmbeddedObjectPart part in ole)
                    {

                    }
                    foreach (EmbeddedPackagePart part in packages)
                    {

                    }
                    foreach (Model3DReferenceRelationshipPart part in threeD)
                    {
                        // Inform user that 3D objects hcannot be processed
                        Console.WriteLine("--> Change: Model 3D reference relationship could not be processed");
                    }
                    foreach (ImagePart part in emf)
                    {
                        // Convert to TIFF
                        Convert_EmbedObj(filepath, workbook, worksheetPart, part);
                    }
                    foreach (ImagePart part in images)
                    {
                        // Convert to TIFF
                        Convert_EmbedObj(filepath, workbook, worksheetPart, part);
                    }
                }
            }
        }

        // General method for converting embedded images to TIFF
        public void Convert_EmbedObj(string filepath, WorkbookPart workbook, WorksheetPart worksheet, ImagePart part) // maybe change ImagePart to OpenXmlPart
        {
            // Define data types
            Uri new_uri;
            string id;
            string new_filename;
            Stream stream = new MemoryStream();
            Stream new_stream = new MemoryStream();

            // Get data
            new_uri = Get_New_Uri(part.Uri);
            new_filename = Get_New_Filename(new_uri);
            stream = part.GetStream();

            // Convert image
            new_stream = Convert_EmbedObj_ImageMagick(stream);

            // Extract image
            Extract_EmbeddedObjects(new_stream, new_filename, filepath);

            // Change relationships of image
            

            // Process - Add new imagepart -> Change blip to new id -> delete old imagepart
            IEnumerable<OpenXmlPart> parentParts = part.GetParentParts();
            foreach (OpenXmlPart parentPart in parentParts)
            {
                if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.DrawingsPart")
                {
                    // Get id of old ImagePart
                    id = parentPart.GetIdOfPart(part);
                    Console.WriteLine(id);

                    // Add new ImagePart
                    ImagePart newImage = worksheet.DrawingsPart.AddImagePart(ImagePartType.Tiff);
                    newImage.FeedData(new_stream);
                    Console.WriteLine(newImage.Uri);

                    // Change blip relationship to new ImagePart
                    
                    Blip blip = ;
                    blip.Embed = worksheet.GetIdOfPart(newImage);

                    // Delete old ImagePart
                    worksheet.DeletePart(id);
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.VmlDrawingPart")
                {
                    // Get id of old ImagePart
                    id = parentPart.GetIdOfPart(part);
                    Console.WriteLine(id);

                    // Add new ImagePart
                    ImagePart newImage = worksheet.AddImagePart(ImagePartType.Tiff);
                    newImage.FeedData(new_stream);
                    Console.WriteLine(newImage.Uri);

                    // 
                    VmlDrawingPart vmlDrawingPart = (VmlDrawingPart)parentPart;
                }
            }
            // Get the blip
            //Blip blip = GetBlipForPicture(new_uri, spreadsheet);

            //part.Uri.MakeRelativeUri(new_uri);

            //XElement change_uri;
            //part.SetXElement();
        }

        // Create new Uri with right extension for embedded object
        public Uri Get_New_Uri(Uri part_uri)
        {
            string new_extension = ".tiff";
            string input_path = part_uri.ToString();
            int dot = input_path.LastIndexOf(".");
            string output_path = input_path.Substring(0, dot) + new_extension;
            Uri new_uri = new Uri(output_path, UriKind.Relative);
            return new_uri;
        }

        // Create new filename with right extension for embedded object
        public string Get_New_Filename(Uri new_uri)
        {
            string filename = new_uri.ToString().Split("/").Last();
            return filename;
        }

        // Convert embedded object to TIFF using ImageMagick
        public Stream Convert_EmbedObj_ImageMagick(Stream stream)
        {
            // Read the input stream in ImageMagick
            using (var image = new MagickImage(stream))
            {
                // Set input stream position to beginning
                stream.Position = 0;

                // Create a memorystream to write image to
                var memStream = new MemoryStream();

                // Write the image to memorystream
                image.SetCompression(CompressionMethod.LZW); // Not working
                image.Write(memStream, MagickFormat.Tiff);

                // Return the memorystream
                return memStream;
            }
        }

        // Change the relationships of the converted embedded object
        public void Change_EmbedObj_Relationships(Stream stream)
        {
            // https://learn.microsoft.com/en-us/dotnet/standard/linq/modify-office-open-xml-document
        }

        public void Delete_EmbedObj(WorksheetPart worksheet, List<OpenXmlPart> parts)
        {
            foreach (OpenXmlPart part in parts)
            {
                worksheet.DeletePart(part);
            }
        }

        // Extract embedded objects
        public void Extract_EmbeddedObjects(Stream input_stream, string new_filename, string filepath)
        {
            // Create new folder for embedded objects
            int backslash = filepath.LastIndexOf("\\");
            string file_folder = filepath.Substring(0, backslash);
            string new_folder = file_folder + "\\Embedded objects";
            Directory.CreateDirectory(new_folder);

            // Extract embedded object to folder
            string output_filepath = new_folder + "\\" + new_filename;
            using (var fileStream = File.Create(output_filepath))
            {
                input_stream.Seek(0, SeekOrigin.Begin);
                input_stream.CopyTo(fileStream);
            }
        }
    }
}
