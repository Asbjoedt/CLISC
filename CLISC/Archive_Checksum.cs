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
        // Calculate MD5 checksum to fingerprint the spreadsheet
        public string Calculate_MD5(string filepath)
        {
            try
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
            // If no converted spreadsheet exist
            catch (System.ArgumentException)
            {
                return "";
            }

            catch (System.IO.FileNotFoundException)
            {
                return "";
            }
        }
    }
}
