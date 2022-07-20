using System;
using System.IO;
using System.IO.Enumeration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLISC
{
    public partial class Spreadsheet
    {

        public void Enumerate_Original(string argument1, string argument3)
        {
            // Prepare for enumeration of files with spreadsheet file extensions
            var spreadsheets_enumeration = new FileSystemEnumerable<FileSystemInfo>(argument1, (ref FileSystemEntry entry) => entry.ToFileSystemInfo(), new EnumerationOptions() { RecurseSubdirectories = true });

            // Enumerate spreadsheets recursively
            if (argument3 == "Recurse=Yes")
            {
                // Create enumeration of files with spreadsheet file extensions
                spreadsheets_enumeration = new FileSystemEnumerable<FileSystemInfo>(argument1, (ref FileSystemEntry entry) => entry.ToFileSystemInfo(), new EnumerationOptions() { RecurseSubdirectories = true })
                {
                    ShouldIncludePredicate = (ref FileSystemEntry entry) =>
                    {
                        if (entry.IsDirectory)
                        {
                            return false;
                        }
                        foreach (string extension in file_format)
                        {
                            var fileExtension = Path.GetExtension(entry.FileName);
                            if (fileExtension.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                        }
                        return false;
                    }
                };
            }

            // Enumerate spreadsheets NOT recursively
            else
            {
                // Create enumeration of files with spreadsheet file extensions
                spreadsheets_enumeration = new FileSystemEnumerable<FileSystemInfo>(argument1, (ref FileSystemEntry entry) => entry.ToFileSystemInfo(), new EnumerationOptions() { RecurseSubdirectories = false })
                {
                    ShouldIncludePredicate = (ref FileSystemEntry entry) =>
                    {
                        if (entry.IsDirectory)
                        {
                            return false;
                        }
                        foreach (string extension in file_format)
                        {
                            var fileExtension = Path.GetExtension(entry.FileName);
                            if (fileExtension.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                        }
                        return false;
                    }
                };
            }
        }
    }
}
