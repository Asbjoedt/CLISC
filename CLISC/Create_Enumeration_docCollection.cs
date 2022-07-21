using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{

    public partial class Spreadsheet
    {
        // Create data types for original spreadsheets
        public string conv_extension = "";
        public string conv_filename = "";
        public string conv_filepath = "";

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
