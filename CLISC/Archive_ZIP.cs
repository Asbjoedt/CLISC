using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Enumeration;
using System.IO.Compression;

namespace CLISC
{
    public partial class Spreadsheet
    {
        // Zip the archivable directory
        public void ZIP_Directory(string Results_Directory)
        {
            // Zip the folder
            string zip_path = results_directory + ".zip";
            ZipFile.CreateFromDirectory(results_directory, zip_path);

            // Create enumeration of unzipped folder and delete it
            DirectoryInfo unzipped_folder = new DirectoryInfo(results_directory);
            foreach (var file in unzipped_folder.EnumerateFiles("*", SearchOption.AllDirectories))
            {
                string item = file.ToString();
                File.Delete(item);
            }
            unzipped_folder = new DirectoryInfo(results_directory + "\\docCollection");
            foreach (var folder in unzipped_folder.EnumerateDirectories("*", SearchOption.AllDirectories))
            {
                string item = folder.ToString();
                Directory.Delete(item);
            }
            Directory.Delete(Results_Directory + "\\docCollection");
            Directory.Delete(Results_Directory);
        }
    }
}
