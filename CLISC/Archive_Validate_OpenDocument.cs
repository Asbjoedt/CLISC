using System.Diagnostics;
using System.ComponentModel;

namespace CLISC
{
    public partial class Validation
    {
        public bool? Validate_OpenDocument(string filepath)
        {
            bool? valid = null;

            try
            {
                // Use ODF Validator for validation of OpenDocument spreadsheets
                Process app = new Process();
                app.StartInfo.UseShellExecute = false;
                app.StartInfo.FileName = "javaw";

                string optionone = "-Djavax.xml.validation.SchemaFactory:<http://relaxng.org/ns/structure/1.0>=org.iso_relax.verifier.jaxp.validation.RELAXNGSchemaFactoryImpl";
                string optiontwo = "-Dorg.iso_relax.verifier.VerifierFactoryLoader=com.sun.msv.verifier.jarv.FactoryLoaderImpl";

                // Use environment variable or direct path
                string? dir = Environment.GetEnvironmentVariable("ODFValidator");
                if (dir != null)
                    app.StartInfo.Arguments = "-jar " + optionone + " " + optiontwo + " " + dir;
                else
                    app.StartInfo.Arguments = "-jar " + optionone + " " + optiontwo + " \"C:\\Program Files\\ODF Validator\\odfvalidator-0.12.0-jar-with-dependencies.jar\" " + filepath;

				app.Start();
                app.WaitForExit();
                int return_code = app.ExitCode;
                app.Close();

                // Inform user of validation results
                if (return_code == 0)
                {
					valid = true;
					Console.WriteLine("--> Validate: File format is valid");
				}
                if (return_code == 1)
                    Console.WriteLine("--> Validate: File format validation could not be completed");
                if (return_code == 2)
					Console.WriteLine("--> Validate: File format is invalid");
                return valid;
            }
            catch (Win32Exception)
            {
                Console.WriteLine("--> Validate: File format validation requires ODF Validator and Java Development Kit");
                return valid;
            }
        }
    }
}
