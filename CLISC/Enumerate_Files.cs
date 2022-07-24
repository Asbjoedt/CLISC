using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Spreadsheet
    {
        // Create private data types for the enumeration
        private string _File_Folder = "", _Org_Filepath = "", _Org_Filename = "", _Org_Extension = "", _Copy_Filepath = "", _Copy_Filename = "", _Copy_Extension = "", _Conv_Filepath = "", _Conv_Filename = "", _Conv_Extension = "";
        private bool _Convert_Success;
        // Create public data types for the enumeration
        public string File_Folder
        {
            get { return _File_Folder; }
            set { _File_Folder = value; }
        }
        public string Org_Filepath
        {
            get { return _Org_Filepath; }
            set { _Org_Filepath = value; }
        }
        public string Org_Filename
        {
            get { return _Org_Filename; }
            set { _Org_Filename = value; }
        }
        public string Org_Extension
        {
            get { return _Org_Extension; }
            set { _Org_Extension = value; }
        }
        public string Copy_Filepath
        {
            get { return _Copy_Filepath; }
            set { _Copy_Filepath = value; }
        }
        public string Copy_Filename
        {
            get { return _Copy_Filename; }
            set { _Copy_Filename = value; }
        }
        public string Copy_Extension
        {
            get { return _Copy_Extension; }
            set { _Copy_Extension = value; }
        }
        public string Conv_Filepath
        {
            get { return _Conv_Filepath; }
            set { _Conv_Filepath = value; }
        }
        public string Conv_Filename
        {
            get { return _Conv_Filename; }
            set { _Conv_Filename = value; }
        }
        public string Conv_Extension
        {
            get { return _Conv_Extension; }
            set { _Conv_Extension = value; }
        }
        public bool Convert_Success
        {
            get { return _Convert_Success; }
            set { _Convert_Success = value; }
        }

        // Enumerate input folder to index files with spreadsheet extensions
        public IEnumerable<string> Enumerate_Original(string argument1, string argument3)
        {
            // Define new enumerable as a list
            IEnumerable<string> Org_Enumeration = new List<string>();
            // Recurse enumeration of original spreadsheets from input directory
            if (argument3 == "Recurse=Yes")
            {
                Org_Enumeration = Directory.EnumerateFiles(argument1, "*.*", SearchOption.AllDirectories)
                    .Where(file => File_Format.Contains(Path.GetExtension(file)));
            }
            // No recurse enumeration
            else
            {
                Org_Enumeration = Directory.EnumerateFiles(argument1, "*.*", SearchOption.TopDirectoryOnly)
                   .Where(file => File_Format.Contains(Path.GetExtension(file)));
            }

            foreach (var file_entry in Org_Enumeration)
            {
                // Create instance for finding file information
                FileInfo file_info = new FileInfo(file_entry);

                // Merge data types
                Org_Extension = file_info.Extension;
                Org_Filename = file_info.Name;
                Org_Filepath = file_info.FullName;
            }

            return Org_Enumeration;
        }


        Spreadsheet[] file_entry = new Spreadsheet[]
        {
            new Spreadsheet { File_Folder = "", Org_Filepath = file_info.FullName, Org_Filename = "", Org_Extension = "", Copy_Filepath = "", Copy_Filename = "", Copy_Extension = "", Conv_Filepath = "", Conv_Filename = "", Conv_Extension = "", Convert_Success = null }
        };

    }
}
