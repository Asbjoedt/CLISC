using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;

namespace CLISC
{
    public partial class Spreadsheet
    {
        public string CalculateMD5(string filepath)
        {
            using (var md5 = MD5.Create())
            {
                using (var stream = File.OpenRead(filepath))
                {
                    var checksum = md5.ComputeHash(stream);
                    return BitConverter.ToString(checksum).Replace("-", "").ToLowerInvariant();
                }
            }
        }
    }
}
