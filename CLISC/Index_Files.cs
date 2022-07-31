using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public class fileIndex
    {
        // Create public data types for use in fileIndex
        public string? File_Folder { get; set; }

        public string Org_Filepath { get; set; }

        public string Org_Filename { get; set; }

        public string Org_Extension { get; set; }

        public string? Copy_Filepath { get; set; }

        public string? Copy_Filename { get; set; }

        public string? Copy_Extension { get; set; }

        public string? Conv_Filepath { get; set; }

        public string? Conv_Filename { get; set; }

        public string? Conv_Extension { get; set; }

        public string? XLSX_Conv_Filepath { get; set; }

        public string? XLSX_Conv_Filename { get; set; }

        public string? XLSX_Conv_Extension { get; set; }

        public string? ODS_Conv_Filepath { get; set; }

        public string? ODS_Conv_Filename { get; set; }

        public string? ODS_Conv_Extension { get; set; }

        public bool? Convert_Success { get; set; }
    }
}
