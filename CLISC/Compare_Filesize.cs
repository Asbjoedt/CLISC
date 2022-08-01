using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Compare
    {
        // Calculate filesize
        public int Calculate_Filesize(string filepath)
        {
            {
                FileInfo file = new FileInfo(filepath);
                int filesize = (int)file.Length;
                int filesize_kb = filesize / 1024;
                return filesize_kb;
            }
        }
    }
}
