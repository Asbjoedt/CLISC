using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Linq;
using DocumentFormat.OpenXml.Office2010.Excel;
using ImageMagick;

namespace CLISC
{
    public partial class Archive_Requirements
    {
        // Convert embedded objects to TIFF using ImageMagick
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
                        binary_fail++;
                    }
                    // Inform user
                    Console.WriteLine($"--> Change: {binary_fail} embedded binary files cannot be processed");

                    foreach (EmbeddedPackagePart part in packages)
                    {
                        // If another OOXML package then convert it to OpenDocument??

                    }

                    // 3D objects cannot be processed - Bug in Open XML SDK?
                    foreach (Model3DReferenceRelationshipPart part in threeD)
                    {
                        threeD_fail++;
                    }
                    // Inform user
                    Console.WriteLine($"--> Change: {threeD_fail} embedded model 3D reference relationships cannot be processed");

                    // Convert Excel-generated .emf images to TIFF
                    foreach (ImagePart part in emf)
                    {
                        // Create new image in TIFF file format and change relationships
                        //Console.WriteLine(part.Uri);
                        //Convert_EmbedObj(filepath, worksheetPart, part);

                        // Delete original embedded image
                        //worksheetPart.DeletePart(part);

                        // Add to success
                        //success++;
                    }

                    // Convert embedded images to TIFF
                    foreach (ImagePart part in images)
                    {
                        // Create new image in TIFF file format and change relationships
                        Convert_EmbedObj(filepath, worksheetPart, part);

                        // Delete original embedded image
                        worksheetPart.DrawingsPart.DeletePart(part);

                        // Add to success
                        success++;
                    }
                }
            }
            return success;
        }

        // General method for converting embedded images to TIFF
        public void Convert_EmbedObj(string filepath, WorksheetPart worksheetPart, ImagePart part) // maybe change ImagePart to OpenXmlPart
        {
            // Define data types
            Uri new_Uri;
            string id;
            string new_Id;
            string new_Filename;
            string parentPartType;
            string output_filepath;
            Stream stream = new MemoryStream();
            Stream new_Stream = new MemoryStream();
            ImagePart new_ImagePart;

            // Get data
            new_Uri = Get_New_Uri(part.Uri);
            new_Filename = Get_New_Filename(new_Uri);
            parentPartType = Get_ParentPart_String(part);
            id = Get_RelationshipId(part);
            stream = part.GetStream();

            // Convert streamed image to new stream
            new_Stream = Convert_EmbedObj_ImageMagick(stream);
            stream.Dispose();

            // Extract converted image to new folder
            output_filepath = Extract_EmbeddedObjects(new_Stream, new_Filename, filepath);

            // Save converted image to new ImagePart
            new_ImagePart = Create_ImagePart(output_filepath, worksheetPart, parentPartType, new_Stream, new_Uri);

            // Change relationships of image
            Change_Blip_Relationship(worksheetPart, new_ImagePart, id);
        }

        // Create new Uri with right extension for embedded object
        public Uri Get_New_Uri(Uri part_Uri)
        {
            string new_extension = ".tiff";
            string input_path = part_Uri.ToString();
            int dot = input_path.LastIndexOf(".");
            string output_path = input_path.Substring(0, dot) + new_extension;
            Uri new_uri = new Uri(output_path, UriKind.Relative);
            return new_uri;
        }

        // Create new filename with right extension for embedded object
        public string Get_New_Filename(Uri new_Uri)
        {
            string filename = new_Uri.ToString().Split("/").Last();
            return filename;
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

        // Create a new Part from the converted image
        public ImagePart Create_ImagePart(string output_filepath, WorksheetPart worksheetPart, string parentPartType, Stream stream, Uri new_Uri)
        {
            ImagePart new_ImagePart;
            string relationshipId;

            switch (parentPartType)
            {
                case "DocumentFormat.OpenXml.Packaging.DrawingsPart":
                    // Get new relationship id
                    relationshipId = Get_New_RelationshipId(worksheetPart.DrawingsPart.ImageParts);

                    // Add new ImagePart
                    new_ImagePart = worksheetPart.DrawingsPart.AddImagePart(ImagePartType.Tiff, relationshipId);
                    

                    break;

                case "DocumentFormat.OpenXml.Packaging.VmlDrawingPart":


                    // Add new ImagePart
                    new_ImagePart = worksheetPart.AddImagePart(ImagePartType.Tiff);
                    break;

                default:
                    // Add new ImagePart
                    new_ImagePart = worksheetPart.AddImagePart(ImagePartType.Tiff);
                    break;
            }
            // Save image from stream to new ImagePart
            stream.Position = 0;
            new_ImagePart.FeedData(stream);

            Console.WriteLine(new_ImagePart.Uri);

            // return the new ImagePart
            return new_ImagePart;
        }

        // Get relationship id from an ImagePart
        public string Get_RelationshipId(ImagePart part)
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
            }
            return id;
        }

        // Create a new available relationship id
        public string Get_New_RelationshipId(IEnumerable<ImagePart> imageParts)
        {
            int number = 1;
            string relationshipId;
            string new_RelationshipId = "rId" + number;
            List<char> ids = new List<char>();
            // Get all relationships to a list
            foreach (ImagePart imagePart in imageParts)
            {
                ids = Get_RelationshipId(imagePart).ToList();
                Console.WriteLine(Get_RelationshipId(imagePart));
            }
            // Enumerate the list until there is no match
            relationshipId = ids.Find(x => x.Equals(new_RelationshipId)).ToString();
            while (relationshipId == new_RelationshipId)
            {
                number++;
                new_RelationshipId = "rId" + number;
                relationshipId = ids.Find(x => x.Equals(new_RelationshipId)).ToString();
            }

            Console.WriteLine(new_RelationshipId);
            return new_RelationshipId;
        }

        // Get parent part as a string value
        public string Get_ParentPart_String(ImagePart part)
        {
            string parentPart_String = "";
            IEnumerable<OpenXmlPart> parentParts = part.GetParentParts();
            foreach (OpenXmlPart parentPart in parentParts)
            {
                if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.DrawingsPart")
                {
                    parentPart_String = parentPart.ToString();
                    return parentPart_String;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.VmlDrawingPart")
                {
                    parentPart_String = parentPart.ToString();
                    return parentPart_String;
                }
                else
                {
                    Console.WriteLine(parentPart.ToString());
                    Console.WriteLine(parentPart.Uri.ToString());
                    //throw new Exception("Spreadsheet contains errors related to embedded objects");
                }
            }
            return parentPart_String;
        }

        // Get the image controls associated with a relationship id
        public Blip Get_Blip(WorksheetPart worksheetPart, string id)
        {
            Blip blip = worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>()
                        .Where(p => p.BlipFill.Blip.Embed == id)
                        .Select(p => p.BlipFill.Blip)
                        .Single();
            return blip;
        }

        // Change the relationships of the converted embedded object
        public void Change_Blip_Relationship(WorksheetPart worksheetPart, ImagePart new_ImagePart, string id)
        {
            // Change blip relationship to new ImagePart
            Blip blip = Get_Blip(worksheetPart, id);
            blip.Embed = Get_RelationshipId(new_ImagePart);
        }

        // Extract embedded objects
        public string Extract_EmbeddedObjects(Stream input_stream, string new_filename, string filepath)
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
            // Return the output filepath
            return output_filepath;
        }

        // Alternative approach
        public void Alternative(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, Uri new_Uri, string output_filepath)
        {
            // https://stackoverflow.com/questions/18569113/open-xml-sdk-addimagepart-change-image-location-from-media-to-word-media

            // Create new part
            System.IO.Packaging.PackagePart packageImagePart = spreadsheet.Package.CreatePart(new_Uri, "Image/tiff");

            // Feed data
            byte[] imageBytes = File.ReadAllBytes(output_filepath);
            packageImagePart.GetStream().Write(imageBytes, 0, imageBytes.Length);

            // Create relationships
            System.IO.Packaging.PackagePart worksheetPackagePart = spreadsheet.WorkbookPart.OpenXmlPackage.Package.GetPart(worksheetPart.Uri);

            Console.Out.WriteLine(worksheetPackagePart.Uri);

            // URI to the image is relative to releationship document.
            System.IO.Packaging.PackageRelationship imageReleationshipPart = worksheetPackagePart.CreateRelationship(new_Uri, System.IO.Packaging.TargetMode.Internal, "http://purl.oclc.org/ooxml/officeDocument/relationships/image");
        }
    }
}