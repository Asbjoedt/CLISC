using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Spreadsheet
    {
        // Enumerate docCollection in two different ways
        public List<string> Enumerate_docCollection(string argument0, string docCollection)
        {
            // For archiving, return enumeration of folders
            if (argument0 == "Count&Convert&Compare&Archive")
            {
                var folder_enumeration = new List<string>();

                // Enumerate of spreadsheets in docCollection
                folder_enumeration = (List<string>)Directory.EnumerateDirectories(docCollection, "*", SearchOption.TopDirectoryOnly)
                    .ToList();

                return folder_enumeration;
            }

            // For ordinary use, return enumeration of files
            else
            {
                var file_enumeration = new List<string>();

                // Enumerate of spreadsheets in docCollection
                file_enumeration = (List<string>)Directory.EnumerateFiles(docCollection, "*.*", SearchOption.TopDirectoryOnly)
                    .ToList();

                return file_enumeration;
            }
        }

    }

}
