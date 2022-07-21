using System;
using System.IO;
using System.IO.Enumeration;
using System.Collections.Generic;
using Enumerate_org = System.Collections.Generic.IEnumerable<T>;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Spreadsheet
    {

        public IEnumerable<T> Enumerate_Original<T>(string argument1, string argument3)
        {
            Enumerate_org<T> org_enumeration = new Enumerate_org<T>();

            // Recurse enumeration of original spreadsheets from input directory
            if (argument3 == "Recurse=Yes")
            {
                org_enumeration = Directory.EnumerateFiles(argument1, "*.*", SearchOption.AllDirectories)
                .Where(file => file.EndsWith(".fods") || file.EndsWith(".ods") || file.EndsWith(".ots") || file.EndsWith("xla") || file.EndsWith(".xls") || file.EndsWith(".xls") || file.EndsWith(".xlt") || file.EndsWith(".xlam") || file.EndsWith(".xlsb") || file.EndsWith(".xlsm") || file.EndsWith(".xlsx") || file.EndsWith(".xltm") || file.EndsWith(".xltx"))
                .ToList();

                return org_enumeration;
            }

            // No recurse enumeration
            else
            {
                org_enumeration = Directory.EnumerateFiles(argument1, "*.*", SearchOption.TopDirectoryOnly)
                .Where(file => file.EndsWith(".fods") || file.EndsWith(".ods") || file.EndsWith(".ots") || file.EndsWith("xla") || file.EndsWith(".xls") || file.EndsWith(".xls") || file.EndsWith(".xlt") || file.EndsWith(".xlam") || file.EndsWith(".xlsb") || file.EndsWith(".xlsm") || file.EndsWith(".xlsx") || file.EndsWith(".xltm") || file.EndsWith(".xltx"))
                .ToList();

                return org_enumeration;
            }

        }

    }

}
