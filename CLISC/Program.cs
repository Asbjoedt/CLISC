using System;
using System.IO;
using System.IO.Enumeration;
using System.IO.Compression;

namespace CLISC
{

    public class Program
    {

        public static void Main(string[] args)
        {

            Console.WriteLine("CLISC - Command Line Interface Spreadsheet Count Convert & Compare");
            Console.WriteLine("@Asbjørn Skødt, web: https://github.com/Asbjoedt/CLISC");
            Console.WriteLine("---");

            string argument1 = Convert.ToString(args[1]);
            string argument2 = Convert.ToString(args[2]);
            string argument3 = Convert.ToString(args[3]);
            string argument4 = Convert.ToString(args[4]);

            try
            {

            // Object reference
            Spreadsheet process = new Spreadsheet();

                // Validate recurse and archive arguments
                if (argument3 == "Recurse=Yes" || argument3 == "Recurse=No")
                {
                    if (argument4 == "Archive=Yes" || argument4 == "Archive=No")
                    {

                        // Method references
                        switch (args[0])
                        {
                            case "Count":
                                process.Count(argument1, argument2, argument3);
                                break;
                            case "Count&Convert":
                                process.Count(argument1, argument2, argument3);
                                process.Convert(argument1, argument2, argument3, argument4);
                                break;
                            case "Count&Convert&Compare":
                                process.Count(argument1, argument2, argument3);
                                process.Convert(argument1, argument2, argument3, argument4);
                                process.Compare(argument1, argument2, argument3, argument4);
                                break;
                            default:
                                Console.WriteLine("Invalid first argument. First argument must be one these: Count, Count&Convert, Count&Convert&Compare");
                                break;
                        }

                    }

                    // Inform user of invalid archive argument
                    else
                    {
                        Console.WriteLine("Invalid archive argument. It must be one of these Archive=Yes or Archive=No");
                    }

                }

                // Inform user of invalid recurse argument
                else
                {
                    Console.WriteLine("Invalid recurse argument. It must be one of these Resurse=Yes or Recurse=No");
                }

            }

            // Inform user of argument errors
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("The number of arguments used are invalid. Consult GitHub documentation");
            }

            finally
            {

                // Zip the output directory
                if (argument4 == "Archive=Yes")
                {
                    // Identify CLISC subdirectory
                    string dateStamp = Spreadsheet.GetTimestamp(DateTime.Now);
                    int results_directory_number = 1;
                    string results_directory = argument2 + "\\CLISC_" + dateStamp + "_v" + results_directory_number;
                    while (Directory.Exists(@results_directory))
                    {
                        results_directory_number++;
                        results_directory = argument2 + "\\CLISC_" + dateStamp + "_v" + results_directory_number;
                    }
                    results_directory_number = results_directory_number - 1;
                    results_directory = argument2 + "\\CLISC_" + dateStamp + "_v" + results_directory_number;

                    // Zip the folder
                    string startPath = results_directory;
                    string zipPath = results_directory + ".zip";

                    ZipFile.CreateFromDirectory(startPath, zipPath);

                    // Create enumeration of unzipped folder and delete it
                    DirectoryInfo unzipped_folder = new DirectoryInfo(results_directory);
                    foreach (var file in unzipped_folder.EnumerateFiles("*", SearchOption.AllDirectories))
                    {
                        string item = file.ToString();
                        File.Delete(item);
                    }
                    unzipped_folder = new DirectoryInfo(results_directory + "\\docCollection");
                    foreach (var folder in unzipped_folder.EnumerateDirectories("*", SearchOption.AllDirectories))
                    {
                        string item = folder.ToString();
                        Directory.Delete(item);
                    }
                    Directory.Delete(results_directory +"\\docCollection");
                    Directory.Delete(results_directory);

                }

                // Inform user of end of CLISC
                Console.WriteLine("CLISC has finished");
                Console.WriteLine("---");
            }

        }

    }

}