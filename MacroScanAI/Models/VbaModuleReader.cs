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

using OpenMcdf;
using System.Diagnostics;
using System.IO;
using System.Text;
using Kavod.Vba.Compression;
using static MacroScanAI.Utils.VbaDirReader;
using MacroScanAI.Controls.TreeAndEditor;

namespace MacroScanAI.Utils
{
    public class VbaModuleReader
    {

        private Dictionary<string, VbaModuleInfo> _vbaModules = new Dictionary<string, VbaModuleInfo>();

        public bool IsVbaModule(Controls.TreeAndEditor.OleNode node)
        {
            if (node == null)
            {
                return false;
            }

            // Traverse up the tree to check parent nodes
            var parent = node;
            bool inVbaFolder = false;
            bool inVbaProjectCur = false;

            while (parent != null)
            {
                if (parent.Name.Equals("VBA", StringComparison.OrdinalIgnoreCase))
                {
                    inVbaFolder = true;
                }

                parent = parent.Parent;
            }

            return inVbaFolder;
        }

        private string DetermineModuleType( OleNode vbaStorage, VbaModuleInfo moduleInfo)
        {
            // 1. Check the ModuleType value (from dir stream)
            if (moduleInfo.Type == ModuleType.StandardModule)
            {
                return "bas";
            }

            if (moduleInfo.Type == ModuleType.ClassDocOrFormModule)
            {
                // 2. Check if a sibling storage exists with the same name
                //    (e.g. _VBA_PROJECT_CUR/UserForm1/)
                bool hasDesignerStorage = vbaStorage.Parent.Children.Any(c =>
                    c.Name.Equals(moduleInfo.ModuleName, StringComparison.OrdinalIgnoreCase)
                    && !c.IsStream);

                if (hasDesignerStorage)
                {
                    return "frm";
                }


                return "cls";
            }

            return "txt";
        }

        public void UpdateModules(Dictionary<string, VbaModuleInfo> modules, OleNode vbaStorage)
        {
            if (modules == null)
            {
                throw new ArgumentNullException(nameof(modules));
            }

            // Clear the existing dictionary and add the new entries
            _vbaModules.Clear();

            foreach (var kvp in modules)
            {
                VbaModuleInfo moduleInfo = kvp.Value;
                string fileExtension = DetermineModuleType(vbaStorage, moduleInfo);
                moduleInfo.SaveFileExtension = fileExtension;

                _vbaModules[kvp.Key] = moduleInfo;
            }
        }

        public bool IsModule(string moduleName)
        {
            if (string.IsNullOrEmpty(moduleName))
            {
                return false;
            }
        

            return _vbaModules.ContainsKey(moduleName);
        }

        public VbaModuleInfo? GetModuleFromName(string moduleName)
        {
            if (_vbaModules.ContainsKey(moduleName))
            {
                return _vbaModules[moduleName];
            }

            return null;
        }

        private string RemoveVbaAttributeLines(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return input;
            }

            var lines = input.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var filteredLines = lines
                .Where(line => !line.TrimStart().StartsWith("Attribute", StringComparison.OrdinalIgnoreCase))
                .ToArray();

            return string.Join(Environment.NewLine, filteredLines);
        }

        public string GetVbaCodeFromModuleStream(CfbStream vbaModuleStream)
        {
            string result = string.Empty;

            try
            {
                string streamName = vbaModuleStream.EntryInfo.Name;


                if (_vbaModules.TryGetValue(streamName, out VbaModuleInfo moduleInfo))
                {

                    uint textOffset = moduleInfo.TextOffset;

                    // Move to the start of the text code in the module stream
                    vbaModuleStream.Position = textOffset;

                    using var ms = new MemoryStream();
                    vbaModuleStream.CopyTo(ms);
                    byte[] compressedBuffer = ms.ToArray();
                    byte[] allBytes = VbaCompression.Decompress(compressedBuffer);

                    // Decode to string using the project code page
                    string vbaText = Encoding.GetEncoding((int)moduleInfo.CodePage).GetString(allBytes);

                    // Optionally trim null terminators
                    vbaText = vbaText.TrimEnd('\0');
                    result = RemoveVbaAttributeLines(vbaText);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return result;
        }

    }
}
