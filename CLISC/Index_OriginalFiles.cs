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

        // Search input directory to index all files
        public static List<orgIndex> Org_Files(string inputdir, bool recurse)
        {
            // Search recursively or not
            SearchOption searchoption = SearchOption.TopDirectoryOnly;
            if (recurse == true)
            {
                searchoption = SearchOption.AllDirectories;
            }

            // Enumerate input directory
            IEnumerable<string> org_enumeration = Directory.EnumerateFiles(inputdir, "*", searchoption).ToList();

            // Create new fileIndex for spreadsheets
            List<orgIndex> Org_File_List = new List<orgIndex>();

            // Enrich metadata of each file and add to index of files if spreadsheet
            foreach (var entry in org_enumeration)
            {
                FileInfo file_info = new FileInfo(entry);
                if (fileFormatIndex.Extension_Array.Contains(file_info.Extension) || fileFormatIndex.Extension_Upper_Array.Contains(file_info.Extension))
                {
                    string extension = file_info.Extension.ToLower();
                    string filename = file_info.Name;
                    string filepath = file_info.FullName;
                    Org_File_List.Add(new orgIndex() { Org_Filepath = filepath, Org_Filename = filename, Org_Extension = extension });
                }
            }
            return Org_File_List;
        }
    }
}
