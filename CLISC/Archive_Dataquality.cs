using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace CLISC
{
    public partial class Spreadsheet
    {
        public string Manipulate_DataQuality(string conv_filepath)
        {
            string dataquality_message = "";

            SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(conv_filepath, false);

            // Check for external relationships
            //external_relationships = spreadsheet.ExternalRelationships;

            //data_parts = spreadsheet.DataParts.ToList;
            //spreadsheet.Close();


            return dataquality_message;
        }
    }
}
