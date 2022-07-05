using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Spreadsheet
    {
        bool password_exist = false;
        bool PasswordProtection(string filepath)
        {
            char[] chBuffer = new char[4096];
            TextReader trReader = new StreamReader(filepath, Encoding.UTF8, true);
            // Read the buffer
            trReader.ReadBlock(chBuffer, 0, chBuffer.Length);
            trReader.Close();
            // Remove non-printable and unicode characters, we're only interested in ASCII character set
            for (int i = 0; i < chBuffer.Length; i++)
            {
                if ((chBuffer[i] < ' ') || (chBuffer[i] > '~')) chBuffer[i] = ' ';
            }
            string strBuffer = new string(chBuffer);
            // .xls format files contains this text when password protected
            if (strBuffer.Contains("M i c r o s o f t   E n h a n c e d   C r y p t o g r a p h i c   P r o v i d e r"))
            {
                password_exist = true;
            }
            // .xlsx format files contain this text when password protected
            if (strBuffer.Contains("E n c r y p t e d P a c k a g e"))
            {
                password_exist = true;
            }
            return (password_exist);
        }
    }
}
