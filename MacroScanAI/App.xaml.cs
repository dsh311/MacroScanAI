/*
 * Copyright (C) 2025 David S. Shelley <davidsmithshelley@gmail.com>
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License 
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */

using MacroScanAI.Controls.TreeAndEditor;
using MacroScanAI.Utils;
using OpenMcdf;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows;
using static MacroScanAI.Utils.VbaDirReader;

namespace MacroScanAI
{
    public class AnalysisReport
    {
        public string SourceFilePath { get; set; } = String.Empty;
        public DateTime ReportGeneratedAt { get; set; } = DateTime.UtcNow;
        public string AllCodeVerdict { get; set; } = "Benign";
        public List<ModuleAnalysisResult> Modules { get; set; } = new List<ModuleAnalysisResult>();
    }

    public partial class App : Application
    {
        // Needed for command line
        public static VbaModuleReader aVBAModuleReader = new VbaModuleReader();

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            bool quietMode = false;

            // Ensure legacy encodings are available for VBA files
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Check command-line arguments
            if (e.Args.Length >= 2)
            {
                ConsoleHelperUtil.ShowConsole();
                Console.WriteLine("");

                string inputPath = e.Args[0];
                string outputDir = e.Args[1];

                // Default: look for environment variable
                string apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY") ?? string.Empty;

                // Parse additional arguments
                for (int i = 2; i < e.Args.Length; i++)
                {
                    if (e.Args[i].Equals("--apikey", StringComparison.OrdinalIgnoreCase) && i + 1 < e.Args.Length)
                    {
                        apiKey = e.Args[i + 1];
                        i++;
                    }
                    else if (e.Args[i].Equals("--quiet", StringComparison.OrdinalIgnoreCase) ||
                             e.Args[i].Equals("-q", StringComparison.OrdinalIgnoreCase))
                    {
                        quietMode = true;
                    }
                }

                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    Console.Error.WriteLine("Error: API key is required. Pass it via --apikey <key> or set OPENAI_API_KEY environment variable.");
                    Console.Out.Flush();
                    Console.Error.Flush();
                    System.Threading.Thread.Sleep(15);
                    Shutdown();
                    return;
                }

                try
                {
                    RunCommandLineMode(inputPath, outputDir, apiKey, quietMode).GetAwaiter().GetResult();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"? Error: {ex.Message}");
                }

                // Exit application after CLI operation
                Console.Out.Flush();
                Console.Error.Flush();
                System.Threading.Thread.Sleep(15);

                Shutdown();
                return;
            }

            // --- GUI mode ---
            var mainWindow = new MainWindow();
            mainWindow.Show();
        }

        void Print(string msg, bool quietMode)
        {
            if (!quietMode)
            {
                Console.WriteLine(msg);
            }
        }

        void PrintAlways(string msg)
        {
            Console.WriteLine(msg);
        }

        private async Task RunCommandLineMode(string inputPath, string outputDir, string apiKey, bool quietMode)
        {
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException($"Input file not found: {inputPath}");
            }

            Directory.CreateDirectory(outputDir);

            try
            {
                RootStorage rootStorage;
                try
                {
                    PrintAlways("-------------------------");
                    PrintAlways("Scanning file: " + inputPath);

                    // Is it an older .xls, .doc, etc
                    rootStorage = RootStorage.OpenRead(inputPath);
                }
                catch (Exception ex)
                {
                    // Check if its a newer .zip office doc
                    RootStorage? maybeRootStorage = StreamInspector.LoadVbaProjectFromZip(inputPath);
                    if (maybeRootStorage == null)
                    {
                        PrintAlways($"Error, could not open root storage.");
                        return;
                    }
                    else
                    {
                        rootStorage = maybeRootStorage;
                    }
                }

                var rootNode = OleTreeBuilder.BuildTree(rootStorage);

                OleNode? oleVBAStorageNode = StreamInspector.GetVBAStorageNodeFromRoot(rootNode);

                Dictionary<string, VbaModuleInfo> vbaModules = new Dictionary<string, VbaModuleInfo>();


                if (oleVBAStorageNode != null)
                {
                    // Find the VBA dir string and use it to find modules
                    var dirStream = StreamInspector.GetStreamByName(oleVBAStorageNode, "dir");
                    if (dirStream != null)
                    {
                        vbaModules = VbaDirReader.GetModuleTextOffsetsFromDirStream(dirStream.Stream);


                        aVBAModuleReader.UpdateModules(vbaModules, oleVBAStorageNode);
                    }
                }

                var report = new AnalysisReport
                {
                    SourceFilePath = inputPath // full path to original file
                };

                bool errorOccured = false;

                foreach (var vbaStorageItem in oleVBAStorageNode.Children)
                {
                    string vbaStorageItemName = vbaStorageItem.Name;

                    if (aVBAModuleReader.IsModule(vbaStorageItemName))
                    {
                        string vbaCode = aVBAModuleReader.GetVbaCodeFromModuleStream(vbaStorageItem.Stream);

                        if (!string.IsNullOrEmpty(vbaCode))
                        {
                            VbaModuleInfo foundModule = vbaModules[vbaStorageItemName];
                            string ext = foundModule.SaveFileExtension;
                            string moduleNameWithExtension = vbaStorageItemName + "." + ext;

                            var result = await AIHelper.AnalyzeVbaModuleAsync(apiKey, vbaCode, moduleNameWithExtension)
                                                       .ConfigureAwait(false);

                            if (result != null)
                            {
                                if (result.Verdict == "Error")
                                {
                                    errorOccured = true;
                                }

                                // Add to report
                                report.Modules.Add(result);

                                // Print to console what is going on
                                Print("   Scanned module: " + result.ModuleName + ", result: " + result.Verdict, quietMode);

                                // Update overall verdict
                                int Rank(string verdict) =>
                                    verdict switch
                                    {
                                        "Malicious" => 3,
                                        "Suspicious" => 2,
                                        "Benign" => 1,
                                        _ => 0
                                    };

                                if (Rank(result.Verdict) > Rank(report.AllCodeVerdict))
                                {
                                    report.AllCodeVerdict = result.Verdict;
                                }
                            }
                        }
                    }
                } // End looping throught vba code modules


                // If one error occured then final result is error, trust nothing
                if (errorOccured)
                {
                    report.AllCodeVerdict = "Error";
                }

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                string jsonString = JsonSerializer.Serialize(report, options);
                string fileName = Path.GetFileName(inputPath);
                string reportFileName = fileName + ".report.json";
                string jsonFilePath = Path.Combine(outputDir, reportFileName);

                // Save JSON
                File.WriteAllText(jsonFilePath, jsonString);

                PrintAlways("Final json report: " + jsonFilePath);
                PrintAlways("Final result: " + report.AllCodeVerdict);
                PrintAlways("-------------------------");


            }
            catch (Exception ex)
            {
                PrintAlways($"Error opening file: {ex.Message}");
            }
        }

    }

}
