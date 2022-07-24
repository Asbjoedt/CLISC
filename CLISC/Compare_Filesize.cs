using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Spreadsheet
    {
        // Calculate filesize
        public int Calculate_Filesize(string Filepath)
        {
            {
                FileInfo file = new FileInfo(Filepath);
                int filesize = (int)file.Length;
                int filesize_kb = filesize / 1024;
                return filesize_kb;
            }
        }
    }
}
