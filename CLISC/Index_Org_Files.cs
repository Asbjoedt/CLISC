using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public class fileIndex
    {
        // Create private data types
        private string _File_Folder = "", _Org_Filepath = "", _Org_Filename = "", _Org_Extension = "", _Copy_Filepath = "", _Copy_Filename = "", _Copy_Extension = "", _Conv_Filepath = "", _Conv_Filename = "", _Conv_Extension = "";
        private bool _Convert_Success;

        // Create public list
        public fileIndex(string _File_Folder, string _Org_Filepath, string _Org_Filename, string _Org_Extension, string _Copy_Filepath, string _Copy_Filename, string _Copy_Extension, string _Conv_Filepath, string _Conv_Filename, string _Conv_Extension, bool _Convert_Success)
        {
            this._File_Folder = _File_Folder;
            this._Org_Filepath = _Org_Filepath;
            this._Org_Filename = _Org_Filename;
            this._Org_Extension = _Org_Extension;
            this._Copy_Filepath = _Copy_Filepath;
            this._Copy_Filename = _Copy_Filename;
            this._Copy_Extension = _Copy_Extension;
            this._Conv_Filepath = _Conv_Filepath;
            this._Conv_Filename = _Conv_Filename;
            this._Conv_Extension = _Conv_Extension;
            this._Convert_Success = _Convert_Success;
        }

        // Create public strings
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

        // Create public bools
        public bool Convert_Success
        {
            get { return _Convert_Success; }
            set { _Convert_Success = value; }
        }

        // Search input directory to index files with spreadsheet extensions
        public static List<fileIndex> Org_Files(string argument1, string argument3)
        {
            // Create new file index
            List<fileIndex> Org_File_List = new List<fileIndex>();
            // Create new temporary list for enumeration of input directory
            var org_enumeration = new List<string>();
            // Recurse enumeration of original spreadsheets from input directory
            if (argument3 == "Recurse=Yes")
            {
                org_enumeration = (List<string>)Directory.EnumerateFiles(argument1, "*.*", SearchOption.AllDirectories)
                    .Where(file => FileFormats.Extension.Contains(Path.GetExtension(file)))
                    .ToList();
            }
            // No recurse enumeration
            else
            {
                org_enumeration = (List<string>)Directory.EnumerateFiles(argument1, "*.*", SearchOption.TopDirectoryOnly)
                   .Where(file => FileFormats.Extension.Contains(Path.GetExtension(file)))
                   .ToList();
            }
            // Enrich metadata of each file and add to index of files
            foreach (var file_entry in org_enumeration)
            {
                // Find file information
                FileInfo file_info = new FileInfo(file_entry);
                string extension = file_info.Extension;
                string filename = file_info.Name;
                string filepath = file_info.FullName;
                // Add original spreadsheets file info to index of files
                Org_File_List.Add(new fileIndex("", filepath, filename, extension, "", "", "", "", "", "", false));
            }
            return Org_File_List;
        }
    }
}
