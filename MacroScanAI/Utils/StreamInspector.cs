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
using System.IO;
using System.IO.Compression;
using System.Text;
using MacroScanAI.Controls.TreeAndEditor;

namespace MacroScanAI.Utils
{
    public static class StreamInspector
    {
        public static OleNode? GetStreamByName(OleNode? parentNode, string streamName)
        {
            if (parentNode == null)
            {
                return null;
            }

            return parentNode.Children
                .FirstOrDefault(c =>
                    c.Name.Equals(streamName, StringComparison.OrdinalIgnoreCase) &&
                    c.IsStream);
        }

        public static OleNode? GetVBAStorageNodeFromRoot(OleNode rootNode)
        {
            if (rootNode == null)
            {
                return null;
            }

            OleNode? foundVBAStorageNode = null;

            // check under _VBA_PROJECT_CUR
            var vbaProjectStorage = rootNode.Children
                .FirstOrDefault(n => n.Name.Equals("_VBA_PROJECT_CUR", StringComparison.OrdinalIgnoreCase) && !n.IsStream);
            if (vbaProjectStorage != null)
            {
                foundVBAStorageNode = GetChildVbaStorage(vbaProjectStorage);
            }

            // if not found, check under Macros (Word)
            if (foundVBAStorageNode == null)
            {
                var macrosStorage = rootNode.Children
                    .FirstOrDefault(n => n.Name.Equals("Macros", StringComparison.OrdinalIgnoreCase) && !n.IsStream);
                if (macrosStorage != null)
                {
                    foundVBAStorageNode = GetChildVbaStorage(macrosStorage);
                }
            }

            // If its comming from a .zip file, then the VBA storage is a child of root
            if (foundVBAStorageNode == null)
            {
                foundVBAStorageNode = GetChildVbaStorage(rootNode);
            }

            return foundVBAStorageNode;

        }

        private static OleNode? GetChildVbaStorage(OleNode parent)
        {
            OleNode? foundVBAStorageNode = null;

            foreach (var child in parent.Children)
            {
                if (foundVBAStorageNode != null) { break; }

                if (child.Name.Equals("VBA", StringComparison.OrdinalIgnoreCase) && !child.IsStream)
                {
                    // We'll assume its the actual VBA storage when it has a dir stream
                    foreach (var vbaItem in child.Children)
                    {
                        if (vbaItem.Name.Equals("dir", StringComparison.OrdinalIgnoreCase) && vbaItem.IsStream)
                        {
                            foundVBAStorageNode = child;
                            break;
                        }
                    }
                }
            }

            return foundVBAStorageNode;
        }

        public static RootStorage? LoadVbaProjectFromZip(string filePath)
        {
            using (var archive = ZipFile.OpenRead(filePath))
            {
                var vbaEntry = archive.Entries.FirstOrDefault(e =>
                    e.FullName.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase));

                if (vbaEntry == null)
                    return null;

                var ms = new MemoryStream();
                using (var entryStream = vbaEntry.Open())
                    entryStream.CopyTo(ms);

                ms.Position = 0;

                try
                {
                    return RootStorage.Open(ms);
                }
                catch
                {
                    // Could not open as CFBF
                    return null;
                }
            }
        }

        public static string GetValidFileName(string? rawName)
        {
            string name = rawName ?? string.Empty;

            // --- Step 1: Remove control characters (like '\u0005')
            string cleanName = new string(name.Where(c => !char.IsControl(c)).ToArray());

            // --- Step 2: Replace invalid file name characters
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                cleanName = cleanName.Replace(c, '_');
            }

            // --- Step 3: If empty, use fallback name with random number
            if (string.IsNullOrWhiteSpace(cleanName))
            {
                var random = new Random();
                int randomNumber = random.Next(10000000, 99999999); // 8 digits
                cleanName = $"EMPTYNAME_{randomNumber}";
            }

            return cleanName;
        }


        // Returns the best-guess file extension for a given stream
        public static string SuggestStreamFileExtension(OleNode node)
        {
            if (!node.IsStream || node.Stream == null)
                return string.Empty;

            string name = node.Name?.ToLowerInvariant() ?? string.Empty;

            // 1. Name-based hints for classic Office internals
            if (name.Contains("worddocument")) return ".doc";
            if (name.Contains("workbook")) return ".xls";
            if (name.Contains("powerpoint document")) return ".ppt";
            if (name.Contains("pictures")) return ".wmf";
            if (name.Contains("summaryinformation")) return ".property";
            if (name.Contains("documentsummaryinformation")) return ".property";
            if (name.Contains("compobj")) return ".compobj";
            if (name.Contains("olepres")) return ".presentation";

            // Read first bytes to check for magic numbers
            byte[] buffer = ReadStreamPrefix(node.Stream, 16);

            if (buffer.Length >= 8)
            {
                // Compound Binary File Format
                if (StartsWith(buffer, new byte[] { 0xD0, 0xCF, 0x11, 0xE0 })) return ".cfb";
                // ZIP (modern Office, OOXML)
                if (StartsWith(buffer, new byte[] { 0x50, 0x4B, 0x03, 0x04 })) return ".zip";
                // PNG
                if (StartsWith(buffer, new byte[] { 0x89, 0x50, 0x4E, 0x47 })) return ".png";
                // JPEG
                if (StartsWith(buffer, new byte[] { 0xFF, 0xD8 })) return ".jpg";
                // PDF
                if (StartsWith(buffer, Encoding.ASCII.GetBytes("%PDF-"))) return ".pdf";
                // RTF
                if (StartsWith(buffer, Encoding.ASCII.GetBytes("{\\rtf"))) return ".rtf";
                // HTML
                if (StartsWith(buffer, Encoding.ASCII.GetBytes("<ht"))) return ".html";
                // Executable or DLL
                if (StartsWith(buffer, new byte[] { 0x4D, 0x5A })) return ".exe";
            }

            // Embedded OLE object check
            if (name.Contains("ole10native") || name.Contains("package"))
            {
                try
                {
                    byte[] data = ReadStreamAll(node.Stream);
                    string? embeddedName = ExtractOleNativeEmbeddedName(data);
                    if (!string.IsNullOrEmpty(embeddedName))
                    {
                        string ext = Path.GetExtension(embeddedName);
                        if (!string.IsNullOrEmpty(ext))
                            return ext;
                    }
                }
                catch { }
            }

            // Fallback
            return ".bin";
        }

        // --- Helper: Reads up to n bytes ---
        private static byte[] ReadStreamPrefix(CfbStream stream, int count)
        {
            byte[] buffer = new byte[Math.Min(count, (int)stream.Length)];
            stream.Seek(0, SeekOrigin.Begin);
            stream.Read(buffer, 0, buffer.Length);
            return buffer;
        }

        // --- Helper: Reads entire stream ---
        private static byte[] ReadStreamAll(CfbStream stream)
        {
            byte[] buffer = new byte[stream.Length];
            stream.Seek(0, SeekOrigin.Begin);
            stream.Read(buffer, 0, buffer.Length);
            return buffer;
        }

        // --- Helper: Compare bytes ---
        private static bool StartsWith(byte[] data, byte[] prefix)
        {
            if (data.Length < prefix.Length) return false;
            for (int i = 0; i < prefix.Length; i++)
            {
                if (data[i] != prefix[i]) return false;
            }
            return true;
        }

        // --- Helper: Extract embedded filename from Ole10Native ---
        private static string? ExtractOleNativeEmbeddedName(byte[] data)
        {
            // OLE10Native format: [header][file name][null][source path][null][temp path][null][data...]
            try
            {
                int offset = 6; // skip header
                int end = Array.IndexOf<byte>(data, 0, offset);
                if (end > offset)
                {
                    string name = Encoding.ASCII.GetString(data, offset, end - offset);
                    return Path.GetFileName(name);
                }
            }
            catch { }
            return null;
        }


    }
}
