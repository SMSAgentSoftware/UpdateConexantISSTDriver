using System;
using WUApiLib;
using Microsoft.Win32;
using System.ServiceProcess;
using System.IO;

// Thanks to: http://www.nullskull.com/a/1592/install-windows-updates-using-c--wuapi.aspx

namespace Install_Conexant_Driver
{
    class Program
    {
        static IUpdate ConexantUpdate = null;
        static string DriverTitle = "Conexant - MEDIA - ";
        static bool RestoreDisableWUA = false;
        static bool RestoreUseWUServer = false;

        static void Main(string[] args)
        {
            ISearchResult uResult = null;
            UpdateSession uSession = null;
            

            // Check if we have access to Windows Update online, if not let's get access temporarily
            Log2File("Program starting", "Information");
            var DisableWUA = GetLMRegKeyValue("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate", "DisableWindowsUpdateAccess");
            var UseWUServer = GetLMRegKeyValue("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU", "UseWUServer");
            bool ServiceRestart = false;
            if (DisableWUA != null && DisableWUA.ToString() == "1")
            {
                Log2File("Opening 'DisableWindowsUpdateAccess' registry key", "Information");
                SetLMRegKeyValue("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate", "DisableWindowsUpdateAccess", 0);
                ServiceRestart = true;
                RestoreDisableWUA = true;
            }
            if (UseWUServer != null && UseWUServer.ToString() == "1")
            {
                Log2File("Opening 'UseWUServer' registry key", "Information");
                SetLMRegKeyValue("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU", "UseWUServer", 0);
                ServiceRestart = true;
                RestoreUseWUServer = true;
            }
            if (ServiceRestart)
            {
                RestartService("wuauserv", 30);               
            }


            // Search for a driver updates
            try
            {
                Log2File("Creating update session", "Information");
                uSession = new UpdateSession();
                IUpdateSearcher uSearcher = uSession.CreateUpdateSearcher();
                uSearcher.Online = true;
                Log2File("Searching for available driver updates", "Information");
                uResult = uSearcher.Search("IsInstalled=0 and Type = 'Driver'");
                
            }
            catch (Exception ex)
            {
                Log2File(ex.Message, "Error");
                RestoreWURegKeys();
                Environment.Exit(1);
            }

            // Exit if none found
            if (uResult.Updates.Count == 0)
            {
                // No updates found
                Log2File("No driver updates found!", "Warning");
                RestoreWURegKeys();
                Environment.Exit(0);
            }

            // Check if the desired update is found
            foreach (IUpdate update in uResult.Updates)
            {
                if (update.Title.StartsWith(DriverTitle))
                {
                    Log2File($"Found update: {update.Title}", "Information");
                    ConexantUpdate = update;
                }
            }

            // If the desired update is found
            if (ConexantUpdate != null && ConexantUpdate.Title.StartsWith(DriverTitle))
            {
                UpdateCollection uCollection = new UpdateCollection();
                // Check if need to download it
                if (ConexantUpdate.IsDownloaded == false)
                {
                    uCollection.Add(ConexantUpdate);
                    try
                    {
                        Log2File("Downloading update", "Information");
                        UpdateDownloader downloader = uSession.CreateUpdateDownloader();
                        downloader.Updates = uCollection;
                        downloader.Download();
                    }
                    catch (Exception ex)
                    {
                        Log2File(ex.Message, "Error");
                        RestoreWURegKeys();
                        Environment.Exit(1);
                    }

                    // Check it was downloaded
                    if (uCollection[0].IsDownloaded == false)
                    {
                        Log2File("The update was not downloaded successfully!", "Error");
                        RestoreWURegKeys();
                        Environment.Exit(1);
                    }
                }
                else
                {
                    Log2File("Update already downloaded", "Information");
                }

                // Install the update
                IInstallationResult installationRes = null;
                try
                {
                    Log2File("Installing update", "Information");
                    IUpdateInstaller installer = uSession.CreateUpdateInstaller();
                    installer.Updates = uCollection;
                    installationRes = installer.Install();
                }
                catch (Exception ex)
                {
                    Log2File(ex.Message, "Error");
                    RestoreWURegKeys();
                    Environment.Exit(1);
                }

                // Check the installation
                if (installationRes.GetUpdateResult(0).HResult == 0)
                {
                    if (installationRes.RebootRequired == true)
                    {
                        Log2File("Update was successfully installed. A reboot is required.", "Information");
                    }
                    else
                    {
                        Log2File("Update was successfully installed.", "Information");
                    }
                }
            }
            else
            {
                Log2File($"No driver update matching '{DriverTitle}' found!", "Warning");
                RestoreWURegKeys();
                Environment.Exit(0);
            }

            // Restore previous Windows Update settings if necessary
            RestoreWURegKeys();

        }

        // Method to get a value from the LM registry
        private static object GetLMRegKeyValue(string Branch, string KeyName)
        {
            Log2File($"Reading registry key '{KeyName}' at '{Branch}'", "Information");
            try
            {
                RegistryKey Key = Registry.LocalMachine.OpenSubKey(Branch);
                if (Key != null)
                {
                    Object v = Key.GetValue(KeyName);
                    if (v != null)
                    {
                        Log2File($"Current value: {v}", "Information");
                        return v;
                    }
                    else
                    {
                        Log2File("Value is null", "Warning");
                        return null;
                    }
                }
                else
                {
                    Log2File("Key is null", "Warning");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Log2File(ex.Message, "Error");
                return null;
            }
        }

        // Method to set a value in the LM registry
        private static string SetLMRegKeyValue(string Branch, string KeyName, object Value)
        {
            RegistryKey Key;
            Log2File($"Setting registry key '{KeyName}' at '{Branch}' with value '{Value}'", "Information");
            try
            {
                Key = Registry.LocalMachine.OpenSubKey(Branch, true);
                if (Key is null)
                {
                    Log2File("Key is null. Cannot set.", "Warning");
                    return "Key not found";
                }

                try
                {
                    Key.SetValue(KeyName, Value);
                    Key.Close();
                    Log2File("Successfully set key", "Information");
                    return "Success";
                }
                catch (Exception ex)
                {
                    Log2File(ex.Message, "Error");
                    return ex.Message;
                }
            }
            catch (Exception ex)
            {
                Log2File(ex.Message, "Error");
                return ex.Message;
            }
        }

        // Method to restart a service
        private static void RestartService(string ServiceName, int Timeout)
        {
            Log2File($"Resarting '{ServiceName}' service with {Timeout} second timeout.", "Information");
            TimeSpan timeout = TimeSpan.FromSeconds(Timeout);
            try
            {
                ServiceController service = new ServiceController(ServiceName);
                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped, timeout);
                Log2File("Service stopped", "Information");
                service.Start();
                service.WaitForStatus(ServiceControllerStatus.Running, timeout);
                Log2File("Service started", "Information");
            }
            catch (Exception ex)
            {
                Log2File(ex.Message, "Error");
                Environment.Exit(1);
            }
        }

        // Method to restore Windows Update registry keys
        private static void RestoreWURegKeys()
        {
            if (RestoreDisableWUA)
            {
                Log2File("Restoring 'DisableWindowsUpdateAccess' reg key value", "Information");
                SetLMRegKeyValue("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate", "DisableWindowsUpdateAccess", 1);
            }
            if (RestoreUseWUServer)
            {
                Log2File("Restoring 'UseWUServer' reg key value", "Information");
                SetLMRegKeyValue("SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU", "UseWUServer", 1);
            }
            if (RestoreUseWUServer || RestoreDisableWUA)
            {
                RestartService("wuauserv", 30);
            }
        }

        // Method to write to a log file in the Temp directory
        private static void Log2File(string Message, string LogLevel)
        {
            var t = Environment.GetEnvironmentVariable("Temp");
            string filePath = $"{t}\\ConexantDriverUpdate.log";
            string LogEntry = $"{DateTime.Now.ToString("HH:mm:ss.fff MM.dd.yyyy")}  |  [{LogLevel}]  |  {Message}";

            try
            {
                using (StreamWriter streamWriter = new StreamWriter(filePath, true))
                {
                    streamWriter.WriteLine(LogEntry);
                    streamWriter.Close();
                }
            }
            catch (Exception)
            {
            }          
        }
    }
}
