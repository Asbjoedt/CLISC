using System;
using System.IO;
using System.IO.Enumeration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Spreadsheet
    {
        public List<string> Enumerate_Original2(string argument1, string argument3)
        {
            
            var org_enumeration = new List<string>();

            // Recurse enumeration of original spreadsheets from input directory
            if (argument3 == "Recurse=Yes")
            {
                org_enumeration = (List<string>)Directory.EnumerateFiles(argument1, "*.*", SearchOption.AllDirectories)
                    .Where(file => File_Format.Contains(Path.GetExtension(file)))
                    .ToList();

                return org_enumeration;
            }

            // No recurse enumeration
            else
            {
                org_enumeration = (List<string>)Directory.EnumerateFiles(argument1, "*.*", SearchOption.TopDirectoryOnly)
                   .Where(file => File_Format.Contains(Path.GetExtension(file)))
                   .ToList();

                return org_enumeration;
            }

        }

    }

}
