using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Linq;
using ImageMagick;
using FFMpegCore;
using FFMpegCore.Pipes;
using FFMpegCore.Enums;
using System.IO.Packaging;

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
                IEnumerable<WorksheetPart>? worksheetParts = spreadsheet.WorkbookPart?.WorksheetParts;

                if (worksheetParts != null)
                {
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
                        foreach (EmbeddedObjectPart embeddedObjectPart in ole)
                        {
                            // Extract object
                            string output_filepath = Create_Output_Filepath(filepath, embeddedObjectPart.Uri.ToString());
                            Stream input_stream = embeddedObjectPart.GetStream();
                            Extract_EmbeddedObjects(input_stream, output_filepath);
                            input_stream.Dispose();

                            // Identify object
                            string AV_type = Identify_Object(embeddedObjectPart);

                            // Convert if correct AV type
                            if (AV_type == "audio" || AV_type == "video")
                            {
                                Convert_EmbedAV(worksheetPart, embeddedObjectPart, AV_type);
                                success++;
                            }
                            else
                            {
                                // Register conversion fail
                                fail++;
                            }
                        }

                        // Embedded packages cannot be converted
                        foreach (EmbeddedPackagePart part in packages)
                        {
                            // Extract object
                            string output_filepath = Create_Output_Filepath(filepath, part.Uri.ToString());
                            Stream input_stream = part.GetStream();
                            Extract_EmbeddedObjects(input_stream, output_filepath);
                            input_stream.Dispose();

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
                            input_stream.Dispose();

                            // Register conversion fail
                            fail++;
                        }

                        // Convert Excel-generated .emf images to TIFF
                        foreach (ImagePart imagePart in emf)
                        {
                            // Extract object
                            string output_filepath = Create_Output_Filepath(filepath, imagePart.Uri.ToString());
                            Stream input_stream = imagePart.GetStream();
                            Extract_EmbeddedObjects(input_stream, output_filepath);
                            input_stream.Dispose();

                            // Convert object
                            Convert_EmbedEmf(worksheetPart, imagePart);
                            success++;
                        }

                        // Convert embedded images to TIFF
                        foreach (ImagePart imagePart in images)
                        {
                            // Extract object
                            string output_filepath = Create_Output_Filepath(filepath, imagePart.Uri.ToString());
                            Stream input_stream = imagePart.GetStream();
                            Extract_EmbeddedObjects(input_stream, output_filepath);
                            input_stream.Dispose();

                            // Convert object
                            Convert_EmbedImg(worksheetPart, imagePart);
                            success++;
                        }
                    }
                }
            }
            // Calculate and return results
            int total = success + fail;
            return Tuple.Create(total, success, fail);
        }

        // Convert embedded images to TIFF
        public void Convert_EmbedImg(WorksheetPart worksheetPart, ImagePart imagePart)
        {
            // Convert streamed image to new stream
            Stream stream = imagePart.GetStream();
            Stream new_Stream = Convert_ImageMagick(stream);
            stream.Dispose();

            if (worksheetPart.DrawingsPart != null)
            {
                // Add new ImagePart
                ImagePart? new_ImagePart = worksheetPart.DrawingsPart.AddImagePart(ImagePartType.Tiff);

                // Save image from stream to new ImagePart
                new_Stream.Position = 0;
                new_ImagePart.FeedData(new_Stream);

                // Change relationships of image
                string id = Get_RelationshipId(imagePart);
                Blip? blip = worksheetPart.DrawingsPart.WorksheetDrawing?.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>()
                                .Where(p => p.BlipFill?.Blip?.Embed == id)
                                .Select(p => p.BlipFill?.Blip)
                                .Single();

                if (blip != null)
                {
                    blip.Embed = Get_RelationshipId(new_ImagePart);
                }

                // Delete original ImagePart
                worksheetPart.DrawingsPart.DeletePart(imagePart);
            }
        }

        // Convert Excel-generated .emf images to TIFF
        public void Convert_EmbedEmf(WorksheetPart worksheetPart, ImagePart imagePart)
        {
            // Convert streamed image to new stream
            Stream stream = imagePart.GetStream();
            Stream new_Stream = Convert_ImageMagick(stream);
            stream.Dispose();

            // Add new ImagePart
            ImagePart new_ImagePart = worksheetPart.VmlDrawingParts.First().AddImagePart(ImagePartType.Tiff);

            // Save image from stream to new ImagePart
            new_Stream.Position = 0;
            new_ImagePart.FeedData(new_Stream);

            // Change relationships of image
            string id = Get_RelationshipId(imagePart);
            XDocument xElement = worksheetPart.VmlDrawingParts.First().GetXDocument();
            IEnumerable<XElement>? descendants = xElement.FirstNode?.Document?.Descendants();

            if (descendants != null)
            {
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
                image.Format = MagickFormat.Tif;
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
            string new_folder = file_folder + "\\Embedded original objects";
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

        public string Identify_Object(EmbeddedObjectPart embeddedObjectPart)
        {
            string AV_type = "";

            // Get extension of object
            int position = embeddedObjectPart.Uri.ToString().LastIndexOf(".");
			string extension = embeddedObjectPart.Uri.ToString().Substring(position);

            // Identifiable extensions
			string[] audio_extensions = { ".flac", ".mp3", ".ogg" };
			string[] video_extensions = { ".avi", ".mp2", ".mp4" };

			// Identify
            if (audio_extensions.Contains(extension))
			{
				AV_type = "audio";
			}
			else if (video_extensions.Contains(extension))
            {
                AV_type = "video";
            }

			return AV_type;
		}

		// Convert embedded audio and video
		public void Convert_EmbedAV(WorksheetPart worksheetPart, EmbeddedObjectPart embeddedObjectPart, string AV_type)
		{
			// Convert streamed AV to new stream
			Stream stream = embeddedObjectPart.GetStream();
            Stream new_Stream = new MemoryStream();
			if (AV_type == "video")
            {
				new_Stream = Convert_Video_FFmpeg(stream);
			}
            else if (AV_type == "audio")
            {
				new_Stream = Convert_Audio_FFmpeg(stream);
			}
			stream.Dispose();

			// Add new EmbeddedObjectPart
			EmbeddedObjectPart new_EmbeddedObjectPart = worksheetPart.AddEmbeddedObjectPart(contentType:"application/vnd.openxmlformats-officedocument.oleObject");

			// Save image from stream to new EmbeddedObjectPart
			new_Stream.Position = 0;
			new_EmbeddedObjectPart.FeedData(new_Stream);

			// Change relationships of EmbeddedObjectPart
			string id = Get_RelationshipId(embeddedObjectPart);


			// Delete original EmbeddedObjectPart
			worksheetPart.DeletePart(embeddedObjectPart);
		}

		public Stream Convert_Video_FFmpeg(Stream input_stream)
		{
			// Define path to FFmpeg
			string? binary_folder = Environment.GetEnvironmentVariable("FFmpeg");
			if (binary_folder == null)
			{
				binary_folder = "C:\\Program Files\\FFmpeg\\bin";
			}
			GlobalFFOptions.Configure(options => options.BinaryFolder = binary_folder);

			// Perform conversion
			Stream output_stream = new MemoryStream();
			FFMpegArguments
	            .FromPipeInput(new StreamPipeSource(input_stream))
	            .OutputToPipe(new StreamPipeSink(output_stream), options => options
		            .WithVideoCodec(VideoCodec.LibX264)
		            .WithAudioCodec(AudioCodec.Aac)
		            .WithVideoFilters(filterOptions => filterOptions
			            .Scale(VideoSize.Original))
		            .ForceFormat("mp4"));

            // Return the converted video as stream
			return output_stream;
		}

		public Stream Convert_Audio_FFmpeg(Stream input_stream)
		{
            // Define path to FFmpeg
            string? binary_folder = Environment.GetEnvironmentVariable("FFmpeg");
			if (binary_folder == null)
			{
                binary_folder = "C:\\Program Files\\FFmpeg\\bin";
			}
			GlobalFFOptions.Configure(options => options.BinaryFolder = binary_folder);

			// Perform conversion
			Stream output_stream = new MemoryStream();
			FFMpegArguments
				.FromPipeInput(new StreamPipeSource(input_stream))
				.OutputToPipe(new StreamPipeSink(output_stream), options => options
					.WithAudioCodec(AudioCodec.LibMp3Lame)
					.ForceFormat("mp3"));

			// Return the converted audio as stream
			return output_stream;
		}
	}
}