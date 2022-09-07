using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public class namespaceIndex
    {
        public string Prefix { get; set; }

        public string Transitional { get; set; }

        public string Strict { get; set; }

        public static List<namespaceIndex> Create_Namespaces_Index() 
        { 
            List<namespaceIndex> list = new List<namespaceIndex>();

            // xmlns
            list.Add(new namespaceIndex() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", Strict = "http://purl.oclc.org/ooxml/spreadsheetml/main" });
            // docProps/
            list.Add(new namespaceIndex() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", Strict = "http://purl.oclc.org/ooxml/officeDocument/extendedProperties" });
            // docProps/vt
            list.Add(new namespaceIndex() { Prefix = "vt", Transitional = "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes", Strict = "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes" });
            // relationships/r
            list.Add(new namespaceIndex() { Prefix = "r", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships" });
            // relationship/styles
            list.Add(new namespaceIndex() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/styles" });
            // relationship/theme
            list.Add(new namespaceIndex() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/theme" });
            // relationship/worksheet
            list.Add(new namespaceIndex() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" });
            // relationship/sharedstrings
            list.Add(new namespaceIndex() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings" });
            // relationship/externalLink
            list.Add(new namespaceIndex() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLink" });
            // relationship/officeDocument
            list.Add(new namespaceIndex() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" });
            // relationship/externallink/externalLinkPath
            list.Add(new namespaceIndex() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLinkPath" });
            // relationship/oleObject
            list.Add(new namespaceIndex() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject" });
            // relationship/image
            list.Add(new namespaceIndex() { Prefix = "", Transitional = "", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/image" });
            // drawingml/a
            list.Add(new namespaceIndex() { Prefix = "a", Transitional = "http://schemas.openxmlformats.org/drawingml/2006/main", Strict = "http://purl.oclc.org/ooxml/drawingml/main" });
            // drawingml/xdr
            list.Add(new namespaceIndex() { Prefix = "xdr", Transitional = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", Strict = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" });
            // customXml/ds
            list.Add(new namespaceIndex() { Prefix = "ds", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/customXml", Strict = "" });
            // urn for Strict - NO NAMESPACE FOR TRANSITIONAL
            list.Add(new namespaceIndex() { Prefix = "v", Transitional = "", Strict = "urn:schemas-microsoft-com:vml" });
            // docProps/core.xml - NO NAMESPACE FOR TRANSITIONAL
            list.Add(new namespaceIndex() { Prefix = "dc", Transitional = "", Strict = "http://purl.org/dc/elements/1.1/" });
            // docProps/core.xml - NO NAMESPACE FOR TRANSITIONAL
            list.Add(new namespaceIndex() { Prefix = "dcterms", Transitional = "", Strict = "http://purl.org/dc/terms/" });
            // docProps/core.xml - NO NAMESPACE FOR TRANSITIONAL
            list.Add(new namespaceIndex() { Prefix = "dcmitype", Transitional = "", Strict = "http://purl.org/dc/dcmitype/" });
            // 
            list.Add(new namespaceIndex() { Prefix = "", Transitional = "", Strict = "" });

            return list;
        }
    }
}
