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

using ICSharpCode.AvalonEdit.Search;
using MacroScanAI.Utils;
using MacroScanAI.Windows.ScanWithAI;
using OpenMcdf;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
using static MacroScanAI.Utils.VbaDirReader;

namespace MacroScanAI.Controls.TreeAndEditor
{
    public partial class TreeAndEditor : System.Windows.Controls.UserControl
    {
        private OleNode _rootNode;
        private VbaModuleReader theModuleReader = new VbaModuleReader();

        private SearchPanel _searchPanel;

        public event Action<string>? FileOpened;

        public enum EditorMode
        {
            Vba,
            Hex
        }

        public TreeAndEditor()
        {
            InitializeComponent();

            // Registers the legacy code pages
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            _searchPanel = SearchPanel.Install(VbaEditor.TextArea);

            VbaEditor.SyntaxHighlighting = new CustomVbaHighlighting(true);
        }

        public void SetEditorMode(EditorMode mode)
        {
            switch (mode)
            {
                case EditorMode.Vba:
                    VbaEditor.SyntaxHighlighting = new CustomVbaHighlighting(true);
                    break;

                case EditorMode.Hex:
                    // No hilighting
                    VbaEditor.SyntaxHighlighting = null;
                    break;
            }

            // Clear undo to avoid weird "no undo group" errors
            VbaEditor.Document.UndoStack.ClearAll();
        }

        public async Task<bool> OpenFileAsync(string fullFilePath)
        {
            try
            {
                RootStorage rootStorage;

                // Attempt to open legacy or modern office formats
                try
                {
                    rootStorage = RootStorage.OpenRead(fullFilePath);
                }
                catch
                {
                    var maybeRootStorage = StreamInspector.LoadVbaProjectFromZip(fullFilePath);
                    if (maybeRootStorage == null)
                    {
                        System.Windows.MessageBox.Show("Could not open root storage.");
                        return false;
                    }
                    rootStorage = maybeRootStorage;
                }

                // Build the tree model
                var rootNode = OleTreeBuilder.BuildTree(rootStorage);

                // Find the VBA storage node
                OleNode? oleVBAStorageNode = StreamInspector.GetVBAStorageNodeFromRoot(rootNode);
                Dictionary<string, VbaModuleInfo> vbaModules = new Dictionary<string, VbaModuleInfo>();

                if (oleVBAStorageNode != null)
                {
                    var dirStream = StreamInspector.GetStreamByName(oleVBAStorageNode, "dir");
                    if (dirStream != null)
                    {
                        vbaModules = VbaDirReader.GetModuleTextOffsetsFromDirStream(dirStream.Stream);
                        theModuleReader.UpdateModules(vbaModules, oleVBAStorageNode);
                    }

                    // Mark modules in tree and update icons
                    MarkVBAStorageModules(oleVBAStorageNode);
                    UpdateVBAStorageTreeIcons(oleVBAStorageNode);
                }

                // Save complete tree without filtering
                _rootNode = rootNode.CloneSelf(rootNode);

                // Find the first VBA module, if it exists
                bool showOnlyModules = ModulesOnlyRadioBtn.IsChecked == true;
                string selectThisModuleName = GetNameOfFirstVBAModuleWithCode(oleVBAStorageNode);
                if (selectThisModuleName != null && selectThisModuleName != String.Empty)
                {
                    // Show the Scan with AI button since there is VBA module
                    scanWithAI_Btn.Visibility = Visibility.Visible;

                    // Show the save button
                    saveFileBtn.Visibility = Visibility.Visible;
                }

                // Apply any tree filtering
                ApplyTreeFilter(ref rootNode, selectThisModuleName, showOnlyModules);

                // Show options for filtering
                showOptionsGrid.Visibility = Visibility.Visible;

                txBxPressF.Visibility = Visibility.Visible;

                // Update TreeView items to the filtered version
                OleTreeView.ItemsSource = new List<OleNode> { rootNode };

                // Yield until layout updates
                await OleTreeView.Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Background);

                // Notify parent that a file has been opened
                FileOpened?.Invoke(fullFilePath);

                return true;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error opening file: {ex.Message}");
                return false;
            }
        }

        private string GetNameOfFirstVBAModuleWithCode(OleNode? oleVBAStorageNode)
        {
            if (oleVBAStorageNode == null)
            {
                return String.Empty;
            }

            foreach (var vbaStorageItem in oleVBAStorageNode.Children)
            {
                string vbaStorageItemName = vbaStorageItem.Name;

                // If the stream is a module then mark it as having module linage
                if (theModuleReader.IsModule(vbaStorageItemName))
                {
                    // Verify it has code and is not empty before choosing
                    if (vbaStorageItem.Stream != null)
                    {
                        string actualVBA = theModuleReader.GetVbaCodeFromModuleStream(vbaStorageItem.Stream);
                        if (actualVBA != String.Empty)
                        {
                            return vbaStorageItemName;
                        }
                    }
                }
            }

            return String.Empty;
        }

        private void MarkVBAStorageModules(OleNode? oleVBAStorageNode)
        {
            if (oleVBAStorageNode == null)
            {
                return;
            }

            foreach (var vbaStorageItem in oleVBAStorageNode.Children)
            {
                string vbaStorageItemName = vbaStorageItem.Name;

                // If the stream is a module then mark it as having module linage
                if (theModuleReader.IsModule(vbaStorageItemName))
                {
                    vbaStorageItem.HasModuleLinage = true;
                }
            }
        }

        private void UpdateVBAStorageTreeIcons(OleNode? oleVBAStorageNode)
        {
            if (oleVBAStorageNode == null)
            {
                return;
            }

            foreach (var vbaStorageItem in oleVBAStorageNode.Children)
            {
                string vbaStorageItemName = vbaStorageItem.Name;

                // If the stream is a module then change icon
                if (theModuleReader.IsModule(vbaStorageItemName))
                {
                    string iconPath = "Assets/file_vbascript.png";

                    VbaModuleInfo? theModule = theModuleReader.GetModuleFromName(vbaStorageItemName);
                    if (theModule != null)
                    {
                        switch (theModule.SaveFileExtension)
                        {
                            case "bas":
                                iconPath = "Assets/file_vbascript_module.png";
                                break;

                            case "cls":
                                iconPath = "Assets/file_vbascript_class.png";
                                break;

                            case "frm":
                                iconPath = "Assets/file_vbascript_form.png";
                                break;

                            default:
                                iconPath = "Assets/file_vbascript.png";
                                break;
                        }
                    }

                    vbaStorageItem.Icon = new BitmapImage(new Uri(iconPath, UriKind.Relative));
                }
                else
                {
                    // Check for dir stream
                    if (vbaStorageItemName.Equals("dir"))
                    {
                        string path = "Assets/file_dir.png";
                        vbaStorageItem.Icon = new BitmapImage(new Uri(path, UriKind.Relative));
                    }

                    // Check for SRP files
                    if (vbaStorageItemName.StartsWith("__SRP"))
                    {
                        string path = "Assets/file_srp.png";
                        vbaStorageItem.Icon = new BitmapImage(new Uri(path, UriKind.Relative));
                    }

                    // Check for the _VBA_PROJECT stream
                    if (vbaStorageItemName.Equals("_VBA_PROJECT"))
                    {
                        string path = "Assets/file_vbaproject.png";
                        vbaStorageItem.Icon = new BitmapImage(new Uri(path, UriKind.Relative));
                    }
                }

            }
        }

        private void ApplyTreeFilter(ref OleNode rootNode, string selectThisModuleName, bool showOnlyModules)
        {
            if (rootNode == null)
            {
                return;
            }

            // Filter the root node recursively
            OleNode? filteredRoot = FilterNode(rootNode, selectThisModuleName, showOnlyModules);

            if (filteredRoot != null)
            {
                OleTreeView.ItemsSource = new List<OleNode> { filteredRoot };
                rootNode = filteredRoot; // update rootNode to the filtered version
            }
            else
            {
                OleTreeView.ItemsSource = null;
                rootNode = null;
            }
        }

        // Recursive filter method: returns null if the node should be removed
        private OleNode? FilterNode(OleNode node, string selectThisModuleName, bool showOnlyModules)
        {
            if (node == null)
            {
                return null;
            }

            // Filter children recursively
            var filteredChildren = node.Children?
                .Select(child => FilterNode(child, selectThisModuleName, showOnlyModules))
                .Where(child => child != null)
                .Cast<OleNode>()
                .ToList() ?? new List<OleNode>();

            // Decide if this node should be included
            bool include = !showOnlyModules || node.HasModuleLinage || filteredChildren.Count > 0;

            if (!include)
            {
                return null;
            }

            // Clone the node for the filtered tree
            var clone = new OleNode
            {
                Name = node.Name,
                Stream = node.Stream,
                Parent = node.Parent,
                HasModuleLinage = node.HasModuleLinage,
                Children = filteredChildren,
                IsExpanded = filteredChildren.Any(c => c.HasModuleLinage || c.Children.Any()),
                IsSelected = ((selectThisModuleName != String.Empty) && (node.Name == selectThisModuleName)),
                Icon = node.Icon,
                IsStream = node.IsStream
            };

            // Update parent references
            foreach (var child in filteredChildren)
            {
                child.Parent = clone;
            }

            return clone;
        }

        public void SetVbaCode(string actualVBA)
        {
            var doc = VbaEditor.Document;
            if (doc == null) return;

            // Run after UI work so we don't race with AvalonEdit's undo handling
            VbaEditor.Dispatcher.InvokeAsync(() =>
            {
                try
                {
                    // Clear undo history so there are no open/partially-closed groups left
                    doc.UndoStack.ClearAll();

                    // Replace document text directly (avoids some of the higher-level TextEditor.Text plumbing)
                    doc.Text = actualVBA;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("SetVbaCode failed: " + ex);
                }
            }, System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        }

        private void OleTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (OleTreeView.SelectedItem is OleNode node && node.IsStream && node.Stream != null)
            {
                // Show the name of the stream on the right side
                streamNameTxtBox.Text = node.Name;

                // VBA module streams
                if (theModuleReader.IsVbaModule(node))
                {
                    // If VBA/dir stream
                    if (VbaDirReader.IsVBADirStream(node))
                    {
                        SetEditorMode(EditorMode.Hex);
                        string hexView = GetFormattedHexViewFromNode(node);
                        SetVbaCode(hexView);
                    }
                    else
                    {
                        // Figure out the type of stream
                        try
                        {
                            if (node.Name.Equals("_VBA_PROJECT"))
                            {
                                SetEditorMode(EditorMode.Hex);
                                string vbaProjectInfo = VbaProjectReader.GetVBAProject(node.Stream);
                                SetVbaCode(vbaProjectInfo);
                            }
                            // If the stream clicked is a 'module'
                            else if (theModuleReader.IsModule(node.Name))
                            {
                                SetEditorMode(EditorMode.Vba);
                                string actualVBA = theModuleReader.GetVbaCodeFromModuleStream(node.Stream);
                                SetVbaCode(actualVBA);
                            }
                            else
                            {
                                SetEditorMode(EditorMode.Hex);
                                string hexView = GetFormattedHexViewFromNode(node);
                                SetVbaCode(hexView);
                            }
                        }
                        catch (Exception ex)
                        {
                            SetEditorMode(EditorMode.Hex);
                            string errMsg = $"\n\n--- VBA Code (decompress failed) ---\n{ex.Message}";
                            SetVbaCode(errMsg);
                        }
                    }
                }
                else
                {
                    SetEditorMode(EditorMode.Hex);
                    // Non-VBA streams, show as hex
                    string hexView = GetFormattedHexViewFromNode(node);
                    SetVbaCode(hexView);
                }
            }
            else
            {
                SetEditorMode(EditorMode.Hex);
                SetVbaCode("");
            }
        }

        private string GetFormattedHexViewFromNode(OleNode node)
        {
            if (node == null || !node.IsStream) { return string.Empty; }

            var sb = new StringBuilder();

            // 1. Hex view
            var bytes = ReadStreamData(node.Stream);
            sb.AppendLine(ToHexString(bytes));

            return sb.ToString();
        }

        private string ToHexString(byte[] data)
        {
            const int bytesPerLine = 16;
            var sb = new System.Text.StringBuilder();

            for (int i = 0; i < data.Length; i += bytesPerLine)
            {
                sb.Append($"{i:X8}: ");
                for (int j = 0; j < bytesPerLine; j++)
                {
                    if (i + j < data.Length)
                        sb.Append($"{data[i + j]:X2} ");
                    else
                        sb.Append("   ");
                }
                sb.Append("  ");
                for (int j = 0; j < bytesPerLine && i + j < data.Length; j++)
                {
                    var b = data[i + j];
                    sb.Append(b >= 32 && b <= 126 ? (char)b : '.');
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private byte[] ReadStreamData(CfbStream stream)
        {
            byte[] buffer = new byte[stream.Length];

            // Reset the stream position to the start
            stream.Seek(0, System.IO.SeekOrigin.Begin);

            // Read the data
            stream.Read(buffer, 0, buffer.Length);

            return buffer;
        }

        private void SaveStream_Click(object sender, RoutedEventArgs e)
        {
            if (OleTreeView.SelectedItem is OleNode node && node.IsStream && node.Stream != null)
            {
                // Get suggested extension (based on name and/or stream content)
                string ext = StreamInspector.SuggestStreamFileExtension(node);

                // Clean the original node name for a valid file system file name
                string safeName = StreamInspector.GetValidFileName(node.Name ?? string.Empty);

                // Combine clean name and suggested extension
                string suggestedFileName = safeName + ext;

                var dlg = new Microsoft.Win32.SaveFileDialog
                {
                    FileName = suggestedFileName,
                    Filter = "All Files|*.*"
                };

                if (dlg.ShowDialog() == true)
                {
                    node.Stream.Seek(0, SeekOrigin.Begin);
                    using var fs = File.Create(dlg.FileName);
                    node.Stream.CopyTo(fs);
                }
            }
        }

        private void OleTreeView_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var clickedItem = VisualUpwardSearch<TreeViewItem>(e.OriginalSource as DependencyObject);
            if (clickedItem != null)
            {
                clickedItem.Focus();
                clickedItem.IsSelected = true;
                e.Handled = false; // Let the context menu still open
            }
        }

        private static T? VisualUpwardSearch<T>(DependencyObject? source) where T : DependencyObject
        {
            while (source != null && source is not T)
            {
                source = VisualTreeHelper.GetParent(source);
            }
            return source as T;
        }

        private void ShowAll_Checked(object sender, RoutedEventArgs e)
        {
            if (_rootNode == null) { return; }

            OleNode copiedRootNode = _rootNode.CloneSelf(_rootNode);

            OleNode? oleVBAStorageNode = StreamInspector.GetVBAStorageNodeFromRoot(copiedRootNode);
            string selectThisModuleName = GetNameOfFirstVBAModuleWithCode(oleVBAStorageNode);
            ApplyTreeFilter(ref copiedRootNode, selectThisModuleName, false);

            // Update TreeView items to the filtered version
            OleTreeView.ItemsSource = new List<OleNode> { copiedRootNode };
        }

        private void ShowModulesOnly_Checked(object sender, RoutedEventArgs e)
        {
            if (_rootNode == null) {  return; }

            OleNode copiedRootNode = _rootNode.CloneSelf(_rootNode);

            OleNode? oleVBAStorageNode = StreamInspector.GetVBAStorageNodeFromRoot(copiedRootNode);
            string selectThisModuleName = GetNameOfFirstVBAModuleWithCode(oleVBAStorageNode);
            ApplyTreeFilter(ref copiedRootNode, selectThisModuleName, true);
            // Update TreeView items to the filtered version
            OleTreeView.ItemsSource = new List<OleNode> { copiedRootNode };
        }

        private async void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "All Files|*.*"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    string fullPath = dlg.FileName;
                    bool fileOpened = await OpenFileAsync(fullPath);
                    if (fileOpened)
                    {
                        docFileNameTxtBox.Text = fullPath;
                        docFileNameTxtBox.ScrollToEnd();
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Error opening file: {ex.Message}");
                }
            }
        }

        private void UserControl_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.F && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                OpenAvalonSearch();
                e.Handled = false; // optional: prevents further processing
            }
        }

        private void OpenAvalonSearch()
        {
            // Make sure the editor has focus
            VbaEditor.Focus();
            ICSharpCode.AvalonEdit.Search.SearchCommands.FindNext.Execute(null, VbaEditor.TextArea);
        }

        private void ScanWithAI_Click(object sender, RoutedEventArgs e)
        {
            var scanWindow = new ScanWithAIWindow(_rootNode);

            // Get the parent window
            var parentWindow = Window.GetWindow(this);
            if (parentWindow != null)
            {
                scanWindow.Owner = parentWindow;
                scanWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                // Set width and height to 90% of parent
                scanWindow.Width = parentWindow.ActualWidth * 0.9;
                scanWindow.Height = parentWindow.ActualHeight * 0.9;
            }

            scanWindow.ShowDialog();
        }

        private void SaveFileButton_Click(object sender, RoutedEventArgs e)
        {
            // Create the dialog
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save File",
                Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                FileName = "code.txt"
            };

            // Show dialog and verify selection
            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                try
                {
                    // Get the text from AvalonEdit
                    string textToSave = VbaEditor.Text;

                    // Save to file
                    File.WriteAllText(dlg.FileName, textToSave);

                    MessageBox.Show("File saved successfully!", "Saved",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error saving file:\n" + ex.Message,
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void Settings_Click(object sender, RoutedEventArgs e)
        {

        }

        private void About_Click(object sender, RoutedEventArgs e)
        {
            string msg =
                "MacroScanAI\n" +
                "Version " + Assembly.GetExecutingAssembly().GetName().Version + "\n" +
                "By David S. Shelley (2025)\n\n" +
                "This software uses several open-source libraries.\n" +
                "See 'NOTICE' and 'THIRD_PARTY_NOTICES.md' for details.";

            MessageBox.Show(msg, "About MacroScanAI", MessageBoxButton.OK, MessageBoxImage.Information);
        }

    }
}
