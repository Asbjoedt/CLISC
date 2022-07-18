using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.ComponentModel;

namespace CLISC
{
    
    public partial class Spreadsheet
    {

        public void Compare_Workbook(string results_directory, string folder, string compare_org_filepath, string compare_conv_filepath)
        {
            try
            {
                //Create "Beyond Compare" script file
                string bcscript_filepath = results_directory + "\\bcscript.txt";
                string bcscript_results_filepath = folder + "\\comparisonResults.html";
                using (StreamWriter bcscript = File.CreateText(bcscript_filepath))
                {
                    bcscript.WriteLine($"data-report layout:interleaved options:display-mismatches title:CLISC_Comparison_Results output-to:\"{bcscript_results_filepath}\" output-options:wrap-word,html-color \"{compare_org_filepath}\" \"{compare_conv_filepath}\"");
                }

                // Use Beyond Compare 4 command line for comparison
                Process app = new Process();
                app.StartInfo.FileName = "C:\\Program Files\\Beyond Compare 4\\BCompare.exe";
                app.StartInfo.Arguments = $"\"@{bcscript_filepath}\" /silent";
                app.Start();
                app.WaitForExit();
                app.Close();

                // Delete BC script
                File.Delete(bcscript_filepath);

                // If there is workbook differences
                //if (fail)
                //{
                //    numTOTAL_diff++;
                //
                //    // Inform user
                //    Console.WriteLine(compare_conv_filepath);
                //    Console.WriteLine($"--> Comparison {success} - Workbook differences identified");
                //}

                // No workbook differences
                //else
                //{
                //    // Inform user
                //    Console.WriteLine($"--> Comparison {success}");
                //}

            }

            // Error message if BC is not detected
            catch (Win32Exception)
            {
                Console.WriteLine($"--> {compare_error_message[1]}");
            }

        }

    }

}
