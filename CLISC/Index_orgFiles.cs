using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public class orgIndex
    {
        // Define data types for this class
        public string Org_Filepath { get; set; }

        public string Org_Filename { get; set; }

        public string Org_Extension { get; set; }

        // Search input directory to index files with spreadsheet extensions
        public static List<orgIndex> Org_Files(string inputdir, bool recurse)
        {
            // Create new temporary list for enumeration of input directory
            var org_enumeration = new List<string>();
            // Recurse enumeration of original spreadsheets from input directory
            if (recurse == true)
            {
                org_enumeration = (List<string>)Directory.EnumerateFiles(inputdir, "*.*", SearchOption.AllDirectories)
                    .Where(file => FileFormats.Extension.Contains(Path.GetExtension(file)))
                    .ToList();
            }
            // No recurse enumeration
            else
            {
                org_enumeration = (List<string>)Directory.EnumerateFiles(inputdir, "*.*", SearchOption.TopDirectoryOnly)
                   .Where(file => FileFormats.Extension.Contains(Path.GetExtension(file)))
                   .ToList();
            }
            // Create new file index
            var Org_File_List = new List<orgIndex>();
            // Enrich metadata of each file and add to index of files
            foreach (var entry in org_enumeration)
            {
                // Find file information
                FileInfo file_info = new FileInfo(entry);
                string extension = file_info.Extension;
                string filename = file_info.Name;
                string filepath = file_info.FullName;
                // Add original spreadsheets file info to index of files
                Org_File_List.Add(new orgIndex() { Org_Filepath = filepath, Org_Filename = filename, Org_Extension = extension });
            }
            return Org_File_List;
        }
    }
}
