using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Results
    {
        // Methods for results reporting
        public void Count_Results()
        {
            Console.WriteLine("---");
            Console.WriteLine("CLISC SUMMARY");
            Console.WriteLine("---");
            Console.WriteLine($"COUNT: {Count.numTOTAL} spreadsheet files in total");
            Console.WriteLine($"Results saved to CSV log in filepath: {Results.CSV_filepath}");
        }

        public void Convert_Results()
        {
            // Calculate fails
            int fail_conversion = Count.numTOTAL - Conversion.numCOMPLETE;

            Console.WriteLine("---");
            Console.WriteLine("CLISC SUMMARY");
            Console.WriteLine("---");
            Console.WriteLine($"COUNT: {Count.numTOTAL} spreadsheet files in total");
            Console.WriteLine($"CONVERT: {fail_conversion} of {Count.numTOTAL} spreadsheets failed conversion");
            Console.WriteLine($"Results saved to CSV log in filepath: {Results.CSV_filepath}");
        }
        public void Compare_Results()
        {
            // Calculate fails
            int fail_conversion = Count.numTOTAL - Conversion.numCOMPLETE;
            int fail_comparison = Conversion.numCOMPLETE - Compare.numTOTAL_compare;

            Console.WriteLine("---");
            Console.WriteLine("CLISC SUMMARY");
            Console.WriteLine("---");
            Console.WriteLine($"COUNT: {Count.numTOTAL} spreadsheet files in total");
            Console.WriteLine($"CONVERT: {fail_conversion} of {Count.numTOTAL} spreadsheets failed conversion");
            Console.WriteLine($"COMPARE: {fail_comparison} of {Conversion.numCOMPLETE} converted spreadsheets failed comparison");
            Console.WriteLine($"COMPARE: {Compare.numTOTAL_diff} of {Compare.numTOTAL_compare} compared spreadsheets have cell value differences");

        }
        public void Archive_Results()
        {
            // Calculate fails
            int fail_conversion = Count.numTOTAL - Conversion.numCOMPLETE;
            int fail_comparison = Conversion.numCOMPLETE - Compare.numTOTAL_compare;

            Console.WriteLine("---");
            Console.WriteLine("CLISC SUMMARY");
            Console.WriteLine("---");
            Console.WriteLine($"COUNT: {Count.numTOTAL} spreadsheets");
            Console.WriteLine($"CONVERT: {fail_conversion} of {Count.numTOTAL} spreadsheets failed conversion");
            Console.WriteLine($"COMPARE: {fail_comparison} of {Conversion.numCOMPLETE} converted spreadsheets failed comparison");
            Console.WriteLine($"COMPARE: {Compare.numTOTAL_diff} of {Compare.numTOTAL_compare} compared spreadsheets have cell value differences");
            Console.WriteLine($"ARCHIVE: {Archive.invalid_files} of {Conversion.numCOMPLETE} converted spreadsheets have invalid file formats");
            Console.WriteLine($"ARCHIVE: {Archive.cellvalue_files} of {Conversion.numCOMPLETE} converted spreadsheets had no cell values");
            Console.WriteLine($"ARCHIVE: {Archive.connections_files} of {Conversion.numCOMPLETE} converted spreadsheets had data connections - Data connections were removed");
            Console.WriteLine($"ARCHIVE: {Archive.cellreferences_files} of {Conversion.numCOMPLETE} converted spreadsheets had external cell references - External cell references were removed");
            Console.WriteLine($"ARCHIVE: {Archive.rtdfunctions_files} of {Conversion.numCOMPLETE} converted spreadsheets had RTD functions - RTD functions were removed");
            Console.WriteLine($"ARCHIVE: {Archive.extobj_files} of {Conversion.numCOMPLETE} converted spreadsheets had external object references - External object references were removed");
            Console.WriteLine($"ARCHIVE: {Archive.embedobj_files} of {Conversion.numCOMPLETE} converted spreadsheets have embedded objects  - Embedded IMAGE objects were converted to .tiff");
            Console.WriteLine($"ARCHIVE: {Archive.printersettings_files} of {Conversion.numCOMPLETE} converted spreadsheets had printer settings - Printer settings were removed");
            Console.WriteLine($"ARCHIVE: {Archive.activesheet_files} of {Conversion.numCOMPLETE} converted spreadsheets did not have active first sheet - Active sheet was changed");
            Console.WriteLine($"ARCHIVE: {Archive.metadata_files} of {Conversion.numCOMPLETE} converted spreadsheets have metadata  - Metadata were NOT removed");
            Console.WriteLine($"ARCHIVE: {Archive.hyperlinks_files} of {Conversion.numCOMPLETE} converted spreadsheets have hyperlinks - Hyperlinks were NOT removed");
        }
    }
}
