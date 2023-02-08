using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Vml;
using ImageMagick;
using System.Xml.Linq;
using System.Xml;

namespace CLISC
{
    public partial class Archive_Requirements
    {
        // Convert embedded objects
        public Tuple<int, int, int> Convert_EmbeddedObjects(string filepath)
        {
            int success = 0;
            int fail = 0;
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

                    // Embedded binaries cannot be converted
                    foreach (EmbeddedObjectPart part in ole)
                    {
                        // Extract object
                        string output_filepath = Create_Output_Filepath(filepath, part.Uri.ToString());
                        Stream input_stream = part.GetStream();
                        Extract_EmbeddedObjects(input_stream, output_filepath);

                        // Register conversion fail
                        fail++;
                    }

                    // Embedded packages cannot be converted
                    foreach (EmbeddedPackagePart part in packages)
                    {
                        // Extract object
                        string output_filepath = Create_Output_Filepath(filepath, part.Uri.ToString());
                        Stream input_stream = part.GetStream();
                        Extract_EmbeddedObjects(input_stream, output_filepath);

                        // Register conversion fail
                        fail++;
                    }

                    // 3D objects cannot be processed - Bug in Open XML SDK?
                    foreach (Model3DReferenceRelationshipPart part in threeD)
                    {
                        // Extract object
                        string output_filepath = Create_Output_Filepath(filepath, part.Uri.ToString());
                        Stream input_stream = part.GetStream();
                        Extract_EmbeddedObjects(input_stream, output_filepath);

                        // Register conversion fail
                        fail++;
                    }

                    // Convert Excel-generated .emf images to TIFF
                    foreach (ImagePart imagePart in emf)
                    {
                        Convert_EmbedEmf(filepath, worksheetPart, imagePart);
                        success++;
                    }

                    // Convert embedded images to TIFF
                    foreach (ImagePart imagePart in images)
                    {
                        Convert_EmbedImg(filepath, worksheetPart, imagePart);
                        success++;
                    }
                }
            }
            // Calculate and return results
            int total = success + fail;
            return Tuple.Create(total, success, fail);
        }

        // Convert embedded images to TIFF
        public void Convert_EmbedImg(string filepath, WorksheetPart worksheetPart, ImagePart imagePart)
        {
            // Convert streamed image to new stream
            Stream stream = imagePart.GetStream();
            Stream new_Stream = Convert_ImageMagick(stream);
            stream.Dispose();

            // Extract converted image to folder
            string extract_filepath = Create_Output_Filepath(filepath, imagePart.Uri.ToString());
            Extract_EmbeddedObjects(new_Stream, extract_filepath);

            // Add new ImagePart
            ImagePart new_ImagePart = worksheetPart.DrawingsPart.AddImagePart(ImagePartType.Tiff);

            // Save image from stream to new ImagePart
            new_Stream.Position = 0;
            new_ImagePart.FeedData(new_Stream);

            // Change relationships of image
            string id = Get_RelationshipId(imagePart);
            Blip blip = worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>()
                            .Where(p => p.BlipFill.Blip.Embed == id)
                            .Select(p => p.BlipFill.Blip)
                            .Single();
            blip.Embed = Get_RelationshipId(new_ImagePart);

            // Delete original ImagePart
            worksheetPart.DrawingsPart.DeletePart(imagePart);
        }

        // Convert Excel-generated .emf images to TIFF
        public void Convert_EmbedEmf(string filepath, WorksheetPart worksheetPart, ImagePart imagePart)
        {
            // Convert streamed image to new stream
            Stream stream = imagePart.GetStream();
            Stream new_Stream = Convert_ImageMagick(stream);
            stream.Dispose();

            // Extract converted image to folder
            string extract_filepath = Create_Output_Filepath(filepath, imagePart.Uri.ToString());
            Extract_EmbeddedObjects(new_Stream, extract_filepath);

            // Add new ImagePart
            ImagePart new_ImagePart = worksheetPart.VmlDrawingParts.First().AddImagePart(ImagePartType.Tiff);

            // Save image from stream to new ImagePart
            new_Stream.Position = 0;
            new_ImagePart.FeedData(new_Stream);

            // Change relationships of image
            string id = Get_RelationshipId(imagePart);
            XDocument xElement = worksheetPart.VmlDrawingParts.First().GetXDocument();
            IEnumerable<XElement> descendants = xElement.FirstNode.Document.Descendants();
            foreach (XElement descendant in descendants)
            {
                if (descendant.Name == "{urn:schemas-microsoft-com:vml}imagedata")
                {
                    IEnumerable<XAttribute> attributes = descendant.Attributes();
                    foreach (XAttribute attribute in attributes)
                    {
                        if (attribute.Name == "{urn:schemas-microsoft-com:office:office}relid")
                        {
                            if (attribute.Value == id)
                            {
                                attribute.Value = Get_RelationshipId(new_ImagePart);
                                worksheetPart.VmlDrawingParts.First().SaveXDocument();
                            }
                        }
                    }
                }
            }
            // Delete original ImagePart
            worksheetPart.VmlDrawingParts.First().DeletePart(imagePart);
        }

        // Convert embedded object to TIFF using ImageMagick
        public Stream Convert_ImageMagick(Stream stream)
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
    }
}