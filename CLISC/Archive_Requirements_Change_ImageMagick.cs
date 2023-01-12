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

                    // Convert embedded objects
                    foreach (EmbeddedObjectPart part in ole)
                    {

                    }
                    foreach (EmbeddedPackagePart part in packages)
                    {
                        // If another OOXML package then convert it to OpenDocument??
                    }
                    foreach (Model3DReferenceRelationshipPart part in threeD)
                    {
                        // Inform user that 3D objects hcannot be processed
                        Console.WriteLine("--> Change: Model 3D reference relationship could not be processed");
                    }
                    foreach (ImagePart part in emf)
                    {
                        // Create new image in TIFF file format and change relationships
                        //Convert_EmbedObj(filepath, worksheetPart, part);
                        // Delete original embedded object
                        //worksheetPart.DeletePart(part);
                    }
                    foreach (ImagePart part in images)
                    {
                        // Create new image in TIFF file format and change relationships
                        Convert_EmbedObj(filepath, worksheetPart, part);
                        // Delete original embedded object
                        worksheetPart.DrawingsPart.DeletePart(part);

                        worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<Relationship>()
                        .Where(p => p.Target == part.Uri)
                        .Select(p => p.BlipFill.Blip)
                        .Single();
                    }
                }
            }
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
            Stream stream = new MemoryStream();
            Stream new_Stream = new MemoryStream();
            ImagePart new_ImagePart;

            // Get data
            new_Uri = Get_New_Uri(part.Uri);
            new_Filename = Get_New_Filename(new_Uri);
            parentPartType = Get_ParentPartString(part);
            id = Get_RelationshipId(part);
            stream = part.GetStream();

            Console.WriteLine(parentPartType);

            // Convert streamed image to new stream
            new_Stream = Convert_EmbedObj_ImageMagick(stream);
            stream.Dispose();

            // Extract converted image to new folder
            Extract_EmbeddedObjects(new_Stream, new_Filename, filepath);

            // Save converted image to new ImagePart
            new_ImagePart = Create_ImagePart(worksheetPart, parentPartType, new_Stream);

            // Change relationships of image
            Change_EmbedObj_Relationships(worksheetPart, new_ImagePart, id);
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

                // Write the image to memorystream
                image.SetCompression(CompressionMethod.LZW); // Not working
                image.Write(new_stream, MagickFormat.Tiff);

                // Return the memorystream
                return new_stream;
            }
        }

        // Create a new Part from the converted image
        public ImagePart Create_ImagePart(WorksheetPart worksheetPart, string parentPartType, Stream stream)
        {
            ImagePart new_ImagePart;
            string relationshipId;

            switch (parentPartType)
            {
                case "DocumentFormat.OpenXml.Packaging.DrawingsPart":
                    // Add new ImagePart
                    new_ImagePart = worksheetPart.DrawingsPart.AddImagePart(ImagePartType.Tiff);
                    
                    // Create new relationship
                    relationshipId = Create_RelationshipId(new_ImagePart);
                    
                    worksheetPart.relatio
                    break;

                case "DocumentFormat.OpenXml.Packaging.VmlDrawingPart":
                    // Create new relationship


                    // Add new ImagePart
                    new_ImagePart = worksheetPart.AddImagePart(ImagePartType.Tiff);
                    break;

                default:
                    // Throw exception if parentPartType is not a valid type
                    throw new Exception("Spreadsheet contains errors related to embedded objects");
            }
            // Save image from stream to new ImagePart
            stream.Position = 0;
            new_ImagePart.FeedData(stream);

            Console.WriteLine(new_ImagePart.Uri);

            // return the new ImagePart
            return new_ImagePart;
        }

        // Change the relationships of the converted embedded object
        public void Change_EmbedObj_Relationships(WorksheetPart worksheetPart, ImagePart new_ImagePart, string id)
        {
            // Change blip relationship to new ImagePart
            Blip blip = Get_Blip(worksheetPart, id);
            blip.Embed = Get_RelationshipId(new_ImagePart);
            Console.WriteLine("new id:" + blip.Embed);
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
        public string Create_RelationshipId(ImagePart imagePart)
        {
            string new_RelationshipId = "";


            return new_RelationshipId;
        }

        // Get parent part as a string value
        public string Get_ParentPartString(ImagePart part)
        {
            string parentPartString = "";
            IEnumerable<OpenXmlPart> parentParts = part.GetParentParts();
            foreach (OpenXmlPart parentPart in parentParts)
            {
                if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.DrawingsPart")
                {
                    parentPartString = parentPart.ToString();
                    return parentPartString;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.VmlDrawingPart")
                {
                    parentPartString = parentPart.ToString();
                    return parentPartString;
                }
            }
            return parentPartString;
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
