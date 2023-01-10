using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Linq;
using ImageMagick;



namespace CLISC
{
    public partial class Archive_Requirements
    {
        // Convert embedded objects to TIFF using ImageMagick
        public void Convert_EmbeddedObjects(string filepath)
        {
            Uri new_uri;
            string new_contentType = "image/tiff";
            string new_filename;
            Stream old_stream = new MemoryStream();
            Stream new_stream = new MemoryStream();

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart worksheetPart in worksheetParts)
                {
                    // Check for parts
                    IEnumerable<EmbeddedObjectPart> ole = worksheetPart.EmbeddedObjectParts;
                    IEnumerable<EmbeddedPackagePart> packages = worksheetPart.EmbeddedPackageParts;
                    IEnumerable<Model3DReferenceRelationshipPart> threeD = worksheetPart.Model3DReferenceRelationshipParts;
                    IEnumerable<ImagePart> emf_images = worksheetPart.ImageParts;
                    List<ImagePart> drawing_images = new List<ImagePart>();
                    if (worksheetPart.DrawingsPart != null)
                    {
                        drawing_images = worksheetPart.DrawingsPart.ImageParts.ToList();
                    }

                    // Convert each part
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
                            // Get data
                            new_uri = Get_New_Uri(part.Uri);
                            new_filename = Get_New_Filename(new_uri);
                            old_stream = part.GetStream();

                            // Convert embedded object
                            new_stream = Convert_to_TIFF(old_stream);

                            // Extract image
                            string path = @"C:\Users\Asbjo\Desktop\" + new_filename;
                            using (var fileStream = File.Create(path))
                            {
                                new_stream.Seek(0, SeekOrigin.Begin);
                                new_stream.CopyTo(fileStream);
                            }

                            // Add new Image
                            ImagePart newImage = worksheetPart.DrawingsPart.AddImagePart(ImagePartType.Tiff);
                            // Put image data into the ImagePart
                            newImage.FeedData(new_stream);
                            Console.WriteLine(newImage.Uri);

                            // Get the blip
                            //Blip blip = GetBlipForPicture(new_uri, spreadsheet);
                            // Point blip at new image
                            //blip.Embed = worksheetPart.GetIdOfPart(newImage);



                            // Get Id of embedded object
                            //string id = part.GetIdOfPart(part);
                            //Console.WriteLine(id);

                            //Change_Embed_Relationships(filepath);

                            //part.Uri.MakeRelativeUri(new_uri);

                            //XElement change_uri;
                            //part.SetXElement();
                        }
                    }
                }
            }
        }

        // Create new Uri with right extension for embedded object
        public Uri Get_New_Uri(Uri part_uri)
        {
            string new_extension = ".tiff";
            string input_path = part_uri.ToString();
            int idx = input_path.LastIndexOf('.');
            string output_path = input_path.Substring(0, idx) + new_extension;
            Uri new_uri = new Uri(output_path, UriKind.Relative);
            return new_uri;
        }

        // Create new filename with right extension for embedded object
        public string Get_New_Filename(Uri new_uri)
        {
            string filename = new_uri.ToString().Split('/').Last();
            return filename;
        }

        // Convert embedded object to TIFF
        public Stream Convert_to_TIFF(Stream stream)
        {
            using (var memStream = new MemoryStream())
            {
                // Change stream to memorystream
                stream.CopyTo(memStream);
                memStream.Position = 0;

                // Set image quality settings
                MagickReadSettings settings = new MagickReadSettings();
                settings.Compression = CompressionMethod.LZW;

                using (var image = new MagickImage())
                {
                    // Set output format
                    image.Format = MagickFormat.Tiff;

                    var info = new MagickImageInfo(memStream);
                    Console.WriteLine(info.Format);

                    // Write the image
                    image.Write(memStream);

                    Console.WriteLine(info.Format);

                    // Return the stream
                    return memStream;
                }
            }
        }

        // Change the relationships of the converted embedded object
        public void Change_Embed_Relationships(Stream stream)
        {
            // https://learn.microsoft.com/en-us/dotnet/standard/linq/modify-office-open-xml-document
        }
    }
}
