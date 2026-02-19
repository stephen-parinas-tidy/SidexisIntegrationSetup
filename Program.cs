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
        /// <summary>
        /// The entry point for the Sidexis integration setup process.
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static Task Main(string[] args)
        {
            try
            {
                // Ensure the program is running with Administrator privileges.
                // Required because URI protocol registration writes to HKCR.
                if (!IsAdministrator())
                {
                    throw new Exception("Please run this program as an administrator.");
                }
                
                Console.WriteLine("SidexisIntegrationSetup has started.");

                // Get the full path of this executable.
                // Used to resolve relative paths to the connector and configuration executables.
                var programPath = typeof(Program).Assembly.Location;
                var mainDirectory = Directory.GetParent(programPath)?.FullName ?? "";
                
                
                // STEP 1: Launch SidexisConnector.exe
                // This executable is responsible for registering the custom URI: SidexisConnector://
                // Once registered, the application can be launched from a browser.
                var sidexisExeFile = Path.Combine(mainDirectory, "SidexisConnector", "SidexisConnector.exe");
                if (!DoesFileExist(sidexisExeFile))
                {
                    throw new Exception("Could not find SidexisConnector.exe.");
                }

                LaunchProgram(sidexisExeFile, "SidexisConnector.exe");
                
                
                // STEP 2: Verify URI registration
                // After running SidexisConnector.exe, the registry should contain a key under HKCR named "SidexisConnector".
                var regKey = Registry.ClassesRoot.OpenSubKey("SidexisConnector", false);
                if (regKey == null)
                {
                    throw new Exception("Could not register the URI 'SidexisConnector://'.");
                }

                Console.WriteLine("The URI 'SidexisConnector://' has been registered.");
                regKey.Close();


                // STEP 3: Launch Sidexis SLIDA Configuration
                // This tool links the integration to Sidexis by configuring communication settings between the systems.
                var slidaExeFile = Path.Combine(mainDirectory, "SidexisSlidaConfiguration", "SidexisSlidaConfiguration.exe");
                if (!DoesFileExist(slidaExeFile))
                {
                    throw new Exception("Could not find SidexisSlidaConfiguration.exe.");
                }

                LaunchProgram(slidaExeFile, "SidexisSlidaConfiguration.exe");

                
                // STEP 4: Confirm integration linkage
                // The integration is considered successfully linked if the mailslot file defined in the SLIDA configuration XML exists.
                // This file is used for message communication with Sidexis.
                
                // Get the XML file for Sidexis configuration
                var slidaXmlFile = Path.Combine(mainDirectory, "SidexisSlidaConfiguration", "SidexisSlidaConfiguration.xml");
                if (!DoesFileExist(slidaXmlFile))
                {
                    throw new Exception("Could not find SidexisSlidaConfiguration.xml.");
                }

                var slidaXmlDoc = new XmlDocument();
                slidaXmlDoc.Load(slidaXmlFile);

                // XPath used to locate the mailslot file path in the configuration XML.
                var mailboxNodePath = "/SlidaConfiguration/CommunicationPartners/CommunicationPartner/MailboxFilename";
                var mailboxNode = slidaXmlDoc.SelectSingleNode(mailboxNodePath);
                if (mailboxNode == null)
                {
                    throw new Exception("Could not get mailslot file.");
                }
                
                // Retrieve the mailbox file path defined in the XML.
                string mailboxFilename = mailboxNode.InnerText.Trim();
                
                // Check if the mailslot file exists. If not, attempt to create it.
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
        
        /// <summary>
        /// Determines whether the current process is running with administrator privileges.
        /// </summary>
        private static bool IsAdministrator()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        /// <summary>
        /// Checks whether a file exists at the specified path.
        /// </summary>
        private static bool DoesFileExist(string filePath)
        {
            return File.Exists(filePath);
        }

        /// <summary>
        /// Launches an external executable and waits for it to exit.
        /// UseShellExecute is set to true so the process launches with standard Windows behaviour.
        /// </summary>
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