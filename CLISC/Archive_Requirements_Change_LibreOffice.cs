using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;

namespace CLISC
{
    public partial class Archive_Requirements
    {
        // Convert embedded package to OpenDocument
        public void Convert_EmbedPackage(string filepath, WorksheetPart worksheetPart, EmbeddedPackagePart part)
        {
            // Extract EmbeddedPackage
            string filename = Get_New_Filename(part.Uri);
            Stream stream = part.GetStream();
            stream.Position = 0;
            string extract_filepath = Extract_EmbeddedObjects(stream, filename, filepath);
            stream.Dispose();

            // Convert to OpenDocument
            string new_Extension = Get_OpenDocument_Extension(extract_filepath);
            Uri new_Uri = Get_New_Uri(part.Uri, new_Extension);
            string new_Filename = Get_New_Filename(new_Uri);
            string output_folder = Path.GetDirectoryName(extract_filepath);
            string output_filepath = output_folder + "\\" + new_Filename;
            Convert_LibreOffice(extract_filepath, output_filepath);

            // Create new EmbeddedPackage
            string contentType = "package";
            EmbeddedPackagePart new_EmbeddedPackagePart = worksheetPart.AddEmbeddedPackagePart(contentType);

            // Feed converted data to the new package
            var new_Stream = new MemoryStream();
            using (FileStream file = new FileStream(output_filepath, FileMode.Open, FileAccess.Read))
                file.CopyTo(new_Stream);
            new_EmbeddedPackagePart.FeedData(new_Stream);

            // Change relationships

        }

        // Transform extension to OpenDocument extension
        public string Get_OpenDocument_Extension(string extract_filepath)
        {
            string extract_Extension = Path.GetExtension(extract_filepath).ToLower();
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
                    throw new Exception("--> Change: Embedded package could not be processed");
            }
        }

        // Convert spreadsheets from OpenDocument file formats using LibreOffice
        public void Convert_LibreOffice(string input_filepath, string output_filepath)
        {
            string format = Path.GetFullPath(output_filepath).Split(".").Last();
            string output_folder = Path.GetDirectoryName(output_filepath);

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
            app.StartInfo.Arguments = "--headless --convert-to " + format + " " + input_filepath + " --outdir " + output_folder;
            app.Start();
            app.WaitForExit();
            app.Close();

            // Delete the original extracted file
            File.Delete(input_filepath);
        }
    }
}
