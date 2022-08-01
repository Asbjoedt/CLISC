using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;

namespace CLISC
{
    public partial class Archive
    {
        // Zip the archive directory
        public void ZIP_Directory(string Results_Directory)
        {
            // Perform zip
            string zip_path = Results_Directory + ".zip";
            ZipFile.CreateFromDirectory(Results_Directory, zip_path);

            // Delete original unzipped archive directory
            DirectoryInfo unzipped_folder = new DirectoryInfo(Results_Directory);
            unzipped_folder.Delete(true);
        }
    }
}
