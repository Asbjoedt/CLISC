using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Spreadsheet
    {
        // Search input directory to index files with spreadsheet extensions
        public static List<fileIndex> Org_Files(string argument1, string argument3)
        {
            // Create new temporary list for enumeration of input directory
            var org_enumeration = new List<string>();
            // Recurse enumeration of original spreadsheets from input directory
            if (argument3 == "Recurse=Yes")
            {
                org_enumeration = (List<string>)Directory.EnumerateFiles(argument1, "*.*", SearchOption.AllDirectories)
                    .Where(file => FileFormats.Extension.Contains(Path.GetExtension(file)))
                    .ToList();
            }
            // No recurse enumeration
            else
            {
                org_enumeration = (List<string>)Directory.EnumerateFiles(argument1, "*.*", SearchOption.TopDirectoryOnly)
                   .Where(file => FileFormats.Extension.Contains(Path.GetExtension(file)))
                   .ToList();
            }
            // Create new file index
            var Org_File_List = new List<fileIndex>();
            // Enrich metadata of each file and add to index of files
            foreach (var entry in org_enumeration)
            {
                // Find file information
                FileInfo file_info = new FileInfo(entry);
                string extension = file_info.Extension;
                string filename = file_info.Name;
                string filepath = file_info.FullName;
                // Add original spreadsheets file info to index of files
                Org_File_List.Add(new fileIndex() { File_Folder = "", Org_Filepath = filepath, Org_Filename = filename, Org_Extension = extension, Copy_Filepath = "", Copy_Filename = "", Copy_Extension = "", Conv_Filepath = "", Conv_Filename = "", Conv_Extension = "", Convert_Success = false });
            }
            return Org_File_List;
        }
    }
}
