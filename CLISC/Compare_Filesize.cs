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
        public int? Calculate_Filesize(string filepath)
        {
            
            int? filesize = null;
            int? filesize_kb = null;

            try
            {
                FileInfo fi = new FileInfo(filepath);
                filesize = (int)fi.Length;
                filesize_kb = filesize / 1024;
                return filesize_kb;
            }

            // If conversion does not exist, do nothing
            catch (SystemException)
            {
                filesize_kb = null;
                return filesize_kb;
            }

        }

    }

}
