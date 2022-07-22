using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{

    public partial class Spreadsheet
    {

        public List<string> Enumerate_docCollection()
        {

            var doc_enumeration = new List<string>();

            // Enumerate of spreadsheets in docCollection
            doc_enumeration = (List<string>)Directory.EnumerateFiles(results_directory, "*.*", SearchOption.TopDirectoryOnly)
                .Where(file => file_format.Contains(Path.GetExtension(file)))
                .ToList();

                return doc_enumeration;

        }

    }

}
