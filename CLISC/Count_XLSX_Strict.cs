using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace CLISC
{

    public partial class Spreadsheet
    {

        public bool ooxml_strict_conformance = false;

        // Identify OOXML Strict conformance
        SpreadsheetDocument OpenSpreadsheetDocument(string filePath)
        {
            return SpreadsheetDocument.Open(filePath, false);
        }

        Tuple<bool, IEnumerable<ValidationErrorInfo>> Validate(OpenXmlPackage doc, FileFormatVersions version)
        {
            OpenXmlValidator openXmlValidator = new OpenXmlValidator();
            bool isStrict = doc.StrictRelationshipFound;
            IEnumerable<ValidationErrorInfo> errors = openXmlValidator.Validate(doc);
            return new Tuple<bool, IEnumerable<ValidationErrorInfo>>(isStrict, errors);
        }

    }

}
