using System;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Win32;

namespace SidexisIntegrationSetup
{
    internal class Program
    {
        public static Task Main(string[] args)
        {
            try
            {
                if (!IsAdministrator())
                {
                    throw new Exception("Please run this program as an administrator.");
                }
                
                Console.WriteLine("SidexisIntegrationSetup has started.");

                var programPath = typeof(Program).Assembly.Location;
                var mainDirectory = Directory.GetParent(programPath)?.FullName ?? "";

                // Launch the Sidexis integration
                // This will register the application's URI onto the machine, which will allow it to be launched from the browser

                var sidexisExeFile = Path.Combine(mainDirectory, "SidexisConnector", "SidexisConnector.exe");
                if (!DoesFileExist(sidexisExeFile))
                {
                    throw new Exception("Could not find SidexisConnector.exe.");
                }

                LaunchProgram(sidexisExeFile, "SidexisConnector.exe");


                // Check that the URI has been registered correctly
                var regKey = Registry.ClassesRoot.OpenSubKey("SidexisConnector", false);
                if (regKey == null)
                {
                    throw new Exception("Could not register the URI 'SidexisConnector://'.");
                }

                Console.WriteLine("The URI 'SidexisConnector://' has been registered.");
                regKey.Close();


                // Launch the SLIDA interface to link the integration to Sidexis
                var slidaExeFile = Path.Combine(mainDirectory, "SidexisSlidaConfiguration",
                    "SidexisSlidaConfiguration.exe");
                if (!DoesFileExist(slidaExeFile))
                {
                    throw new Exception("Could not find SidexisSlidaConfiguration.exe.");
                }

                LaunchProgram(slidaExeFile, "SidexisSlidaConfiguration.exe");


                // Confirm that the integration is linked to (i.e. can send messages to) Sidexis,
                // This is done by checking if the mailslot file exists

                // Get the XML file for Sidexis configuration
                var slidaXmlFile = Path.Combine(mainDirectory, "SidexisSlidaConfiguration",
                    "SidexisSlidaConfiguration.xml");
                if (!DoesFileExist(slidaXmlFile))
                {
                    throw new Exception("Could not find SidexisSlidaConfiguration.xml.");
                }

                var slidaXmlDoc = new XmlDocument();
                slidaXmlDoc.Load(slidaXmlFile);

                // Get the file path of the Sidexis mailslot file
                var mailboxNodePath = "/SlidaConfiguration/CommunicationPartners/CommunicationPartner/MailboxFilename";
                var mailboxNode = slidaXmlDoc.SelectSingleNode(mailboxNodePath);
                if (mailboxNode == null)
                {
                    throw new Exception("Could not get mailslot file.");
                }

                // Check if the mailslot file exists
                string mailboxFilename = mailboxNode.InnerText.Trim();
                if (DoesFileExist(mailboxFilename))
                {
                    Console.WriteLine("TidyClinic integration successfully linked to Sidexis.");
                }
                else
                {
                    File.Create(mailboxFilename);
                    if (DoesFileExist(mailboxFilename))
                    {
                        Console.WriteLine("TidyClinic integration successfully linked to Sidexis.");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.Message}");
            }
            finally
            {
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }

            return Task.CompletedTask;
        }
        
        private static bool IsAdministrator()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static bool DoesFileExist(string filePath)
        {
            return File.Exists(filePath);
        }

        private static void LaunchProgram(string programPath, string programName)
        {
            // Create start info
            var startInfo = new ProcessStartInfo
            {
                FileName = programPath,
                UseShellExecute = true
            };

            // Start the process
            var process = Process.Start(startInfo);
            if (process != null)
            {
                Console.WriteLine($"Running {programName}");
                process.WaitForExit();
            }
            else
            {
                throw new Exception($"Could not start {programName}");
            }
        }
    }
}