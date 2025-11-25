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
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using static MacroScanAI.Utils.VbaDirReader;

namespace MacroScanAI.Windows.ScanWithAI
{
    /// <summary>
    /// Interaction logic for ScanWithAIWindow.xaml
    /// </summary>
    public partial class ScanWithAIWindow : Window
    {
        private VbaModuleReader aVBAModuleReader = new VbaModuleReader();

        private OleNode _rootNode;

        public ScanWithAIWindow()
        {
            InitializeComponent();

            // Get the API key from environment variable
            string? apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");

            // If it's null, default to empty string
            ApiKeyBox.Password = apiKey ?? string.Empty;

            SetOverallRichTextBox("Not Started");
        }

        public ScanWithAIWindow(OleNode rootNode) : this()
        {
            _rootNode = rootNode;
        }

        private async void BeginScanButton_Click(object sender, RoutedEventArgs e)
        {
            string apiKey = ApiKeyBox.Password;

            try
            {
                // Clear the results, incase they run scan twice
                ResultsBox.Document.Blocks.Clear();
                OverallResultsBox.Document.Blocks.Clear();

                GetAllModuleCode(apiKey);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error calling OpenAI: " + ex.Message);
            }
        }

        private async void GetAllModuleCode(string apiKey)
        {
            if (_rootNode == null || apiKey == string.Empty)
            {
                return;
            }

            OleNode? oleVBAStorageNode = StreamInspector.GetVBAStorageNodeFromRoot(_rootNode);
            Dictionary<string, VbaModuleInfo> vbaModules = new Dictionary<string, VbaModuleInfo>();

            int totalModules = 0;
            if (oleVBAStorageNode != null)
            {
                // Find the VBA dir string and use it to find modules
                var dirStream = StreamInspector.GetStreamByName(oleVBAStorageNode, "dir");
                if (dirStream != null)
                {
                    vbaModules = VbaDirReader.GetModuleTextOffsetsFromDirStream(dirStream.Stream);


                    aVBAModuleReader.UpdateModules(vbaModules, oleVBAStorageNode);
                }

                totalModules = oleVBAStorageNode.Children.Count;
            }

            // Now dump all the VBA code found in the modules
            SetOverallRichTextBox("Pending");
            int processedModules = 0;
            string allCodeVerdict = "Benign";
            foreach (var vbaStorageItem in oleVBAStorageNode.Children)
            {
                string vbaStorageItemName = vbaStorageItem.Name;
                processedModules++;

                if (aVBAModuleReader.IsModule(vbaStorageItemName))
                {
                    string vbaCode = aVBAModuleReader.GetVbaCodeFromModuleStream(vbaStorageItem.Stream);
                    if (!string.IsNullOrEmpty(vbaCode))
                    {
                        VbaModuleInfo foundModule = vbaModules[vbaStorageItemName];
                        string ext = foundModule.SaveFileExtension;
                        string moduleNameWithExtension = vbaStorageItemName + "." + ext;

                        var result = await AIHelper.AnalyzeVbaModuleAsync(apiKey, vbaCode, moduleNameWithExtension);

                        if (result != null)
                        {
                            int Rank(string verdict) =>
                            verdict switch
                            {
                                "Malicious" => 3,
                                "Suspicious" => 2,
                                "Benign" => 1,
                                _ => 0   // unknown = best (harmless)
                            };

                            if (Rank(result.Verdict) > Rank(allCodeVerdict))
                            {
                                allCodeVerdict = result.Verdict;
                            }
                        }

                        AppendResultToRichTextBox(result);
                    }
                }

                // Update ProgressBar (make sure to scale 0-100)
                ScanProgressBar.Value = (double)processedModules / totalModules * 100;
            }

            if (processedModules == 0)
            {
                //ResultsTextBox.AppendText("No VBA code found to scan");
            }

            SetOverallRichTextBox(allCodeVerdict);

        }
        private void AppendResultToRichTextBox(ModuleAnalysisResult result)
        {
            if (ResultsBox.Document == null)
            {
                ResultsBox.Document = new FlowDocument();
            }

            var paragraph = new Paragraph();

            // Module Name header
            paragraph.Inlines.Add(new Bold(new Run(result.ModuleName + "\n")) { Foreground = Brushes.Cyan, FontSize = 18 });

            // Verdict with color based on value
            Brush verdictColor = result.Verdict switch
            {
                "Malicious" => Brushes.Red,
                "Suspicious" => Brushes.Orange,
                "Benign" => Brushes.LightGreen,
                _ => Brushes.White
            };
            paragraph.Inlines.Add(new Run("Verdict: ") { Foreground = Brushes.White, FontWeight = FontWeights.Bold });
            paragraph.Inlines.Add(new Run(result.Verdict + "  ") { Foreground = verdictColor, FontWeight = FontWeights.Bold });

            // Confidence
            paragraph.Inlines.Add(new Run($"Confidence: {result.Confidence}\n") { Foreground = Brushes.LightBlue });

            // Summary
            paragraph.Inlines.Add(new Run(result.Summary + "\n") { Foreground = Brushes.White });

            // Indicators
            if (result.Indicators != null && result.Indicators.Count > 0)
            {
                foreach (var ind in result.Indicators)
                {
                    paragraph.Inlines.Add(new Run("• " + ind + "\n") { Foreground = Brushes.Yellow, FontWeight = FontWeights.Bold });
                }
            }

            ResultsBox.Document.Blocks.Add(paragraph);

            // Add a horizontal line
            var line = new Border
            {
                BorderBrush = Brushes.Gray,
                BorderThickness = new Thickness(0, 0.5, 0, 0),
                Margin = new Thickness(0, 4, 0, 4)
            };
            ResultsBox.Document.Blocks.Add(new BlockUIContainer(line));

            // Scroll to end
            ResultsBox.ScrollToEnd();
        }

        private void SetOverallRichTextBox(string theVerdict)
        {
            OverallResultsBox.Document.Blocks.Clear();

            var paragraph = new Paragraph();

            // Module Name header
            paragraph.Inlines.Add(new Bold(new Run("Overall Result: ")) { Foreground = Brushes.Cyan });

            // Verdict with color based on value
            Brush verdictColor = theVerdict switch
            {
                "Malicious" => Brushes.Red,
                "Suspicious" => Brushes.Orange,
                "Benign" => Brushes.LightGreen,
                _ => Brushes.White
            };
            paragraph.Inlines.Add(new Run(theVerdict + "  ") { Foreground = verdictColor, FontWeight = FontWeights.Bold });

            OverallResultsBox.Document.Blocks.Add(paragraph);

            // Scroll to end
            OverallResultsBox.ScrollToEnd();
        }

    }
}
