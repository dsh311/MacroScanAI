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
using System.Text;

namespace MacroScanAI.Utils
{
    internal static class VbaProjectReader
    {

        public static string GetVBAProject(CfbStream vbaModuleStream)
        {
            var result = new StringBuilder();

            try
            {
                vbaModuleStream.Position = 0;
                using var br = new BinaryReader(vbaModuleStream, Encoding.Default, leaveOpen: true);

                // Spec: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/ef7087ac-3974-4452-aab2-7dba2214d239
                ushort reserved1 = br.ReadUInt16();  // MUST be 0x61CC. MUST be ignored.
                ushort version = br.ReadUInt16();    // MUST be ignored on read.
                byte reserved2 = br.ReadByte();      // MUST be 0x00. MUST be ignored.
                ushort reserved3 = br.ReadUInt16();  // Undefined. MUST be ignored.

                // --- Remaining bytes (performance cache) ---
                int performanceCacheSize = (int)(vbaModuleStream.Length - vbaModuleStream.Position);
                byte[] performanceCacheBytes = br.ReadBytes(performanceCacheSize);

                // --- Build output string ---
                result.AppendLine("VBA Project Stream Info:");
                result.AppendLine($"Reserved1 = 0x{reserved1:X4}");
                result.AppendLine($"Version    = 0x{version:X4}");
                result.AppendLine($"Reserved2  = 0x{reserved2:X2}");
                result.AppendLine($"Reserved3  = 0x{reserved3:X4}");
                result.AppendLine();
                result.AppendLine($"PerformanceCache ({performanceCacheSize} bytes):");
                result.AppendLine(DumpHex(performanceCacheBytes));
            }
            catch (Exception ex)
            {
                result.AppendLine($"Error reading _VBA_PROJECT stream: {ex.Message}");
            }

            return result.ToString();
        }

        private static string DumpHex(byte[] data)
        {
            const int bytesPerLine = 16;
            var sb = new StringBuilder();

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
                    byte b = data[i + j];
                    sb.Append(b >= 32 && b <= 126 ? (char)b : '.');
                }

                sb.AppendLine();
            }

            return sb.ToString();
        }

    }
}
