using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Linq;
using ImageMagick;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.ExtendedProperties;

namespace CLISC
{
    public partial class Archive_Requirements
    {
        // Convert embedded objects to TIFF using ImageMagick
        public void Convert_EmbeddedObjects(string filepath)
        {
            string new_extension = ".tif";
            string input_embed_filepath;
            string output_embed_filepath;
            string id;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart worksheetPart in worksheetParts)
                {
                    IEnumerable<EmbeddedObjectPart> ole = worksheetPart.EmbeddedObjectParts;
                    IEnumerable<EmbeddedPackagePart> packages = worksheetPart.EmbeddedPackageParts;
                    IEnumerable<Model3DReferenceRelationshipPart> threeD = worksheetPart.Model3DReferenceRelationshipParts;
                    IEnumerable<ImagePart> emf_images = worksheetPart.ImageParts;
                    IEnumerable<ImagePart> drawing_images = new List<ImagePart>();
                    if (worksheetPart.DrawingsPart != null)
                    {
                        drawing_images = worksheetPart.DrawingsPart.ImageParts;
                    }

                    if (ole.Count() > 0)
                    {
                        foreach (EmbeddedObjectPart part in ole)
                        {

                        }
                    }

                    if (packages.Count() > 0)
                    {
                        foreach (EmbeddedPackagePart part in packages)
                        {

                        }
                    }
                    if (threeD.Count() > 0)
                    {
                        foreach (Model3DReferenceRelationshipPart part in threeD)
                        {

                        }
                    }
                    if (emf_images.Count() > 0)
                    {
                        foreach (ImagePart part in emf_images)
                        {

                        }
                    }
                    if (drawing_images.Count() > 0)
                    {
                        foreach (ImagePart part in drawing_images)
                        {
                            // Create new Uri
                            input_embed_filepath = part.Uri.ToString();
                            int idx = input_embed_filepath.LastIndexOf('.');
                            output_embed_filepath = input_embed_filepath.Substring(0, idx) + new_extension;
                            Uri new_uri = new Uri(output_embed_filepath, UriKind.Relative);
                            Console.WriteLine(new_uri);

                            // Convert
                            Stream stream = part.GetStream();
                            stream = Convert_to_TIF(stream);

                            // Change relationships
                            //Change_Embed_Relationships(filepath);

                            id = part.GetIdOfPart(part);
                            Console.WriteLine(id);
                            ReferenceRelationship reference = part.GetReferenceRelationship(id);
                            Console.WriteLine(reference);

                            //part.Uri.MakeRelativeUri(new_uri);

                            //XElement change_uri;
                            //part.SetXElement();
                        }
                    }
                }
            }
        }

        // Convert embedded object to TIF
        public Stream Convert_to_TIF(Stream stream)
        {
            stream.Position = 0;

            using (var memStream = new MemoryStream())
            {
                // Convert stream to memorystream
                stream.CopyTo(memStream);

                // Determine the image quality settings
                MagickReadSettings settings = new MagickReadSettings();
                settings.Density = new Density(300, 300);
                settings.Compression = CompressionMethod.LZW;

                using (MagickImage image = new MagickImage())
                {
                    // Read the file
                    image.Read(memStream, settings);

                    // Sets the output format
                    image.Format = MagickFormat.Tif;

                    // Write the image
                    image.Write(memStream);

                    // Convert memorystream to stream
                    memStream.CopyTo(stream);
                }
            }

            // Return the stream
            return stream;
        }

        // Change the relationships of the converted embedded object
        public void Change_Embed_Relationships(Stream stream)
        {
            // https://learn.microsoft.com/en-us/dotnet/standard/linq/modify-office-open-xml-document
        }
    }
}
