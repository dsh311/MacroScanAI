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
using Kavod.Vba.Compression;
using MacroScanAI.Controls.TreeAndEditor;
using System.Diagnostics; // Debug writeline

namespace MacroScanAI.Utils
{
    public static class VbaDirReader
    {
        public static bool IsVBADirStream(OleNode node)
        {
            if (node == null)
            {
                return false;
            }
            string name = node.Name;
            // VBA streams are case-insensitive
            return string.Equals(name, "dir", StringComparison.OrdinalIgnoreCase);
        }

        // PROJECT INFORMATION Record
        // See diagram https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/5abef063-3661-46dd-ba80-8cb507afdb1d
        const ushort PROJECTSYSKIND = 0x0001; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/338ed66e-45f0-4550-819d-a25f8450439a
        const ushort PROJECTLCID = 0x0002; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/1136037b-5e9e-4e2d-81f8-615ace60be9d
        const ushort PROJECTCODEPAGE = 0x0003; // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/bd134afb-1cbd-4ceb-8b11-bfa822455655
        const ushort PROJECTNAME = 0x0004; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/28ea157e-1ae0-43e7-b7b2-8fc885f6e5fa
        const ushort PROJECTDOCSTRING = 0x0005; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/dc196b9e-3299-49d0-9651-1d690bb89b8b
        const ushort PROJECTHELPFILEPATH = 0x0006; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/b1e1f51f-6bef-49fe-b6a9-76e174d51b0d
        const ushort PROJECTHELPCONTEXT = 0x0007; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/ce2aae43-1f7a-41e5-b9ba-005bbe445214
        const ushort PROJECTLIBFLAGS = 0x0008; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/d6eb54ae-d765-47f4-9feb-1fe94ee9f47e
        const ushort PROJECTVERSION = 0x0009; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/29fbfea3-498d-4dac-8db1-f765213aced3
        const ushort PROJECTCONSTANTS = 0x000C; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/042a3b56-56bc-4897-bcb1-4138e05b996e
        const ushort PROJECTLCIDINVOKE = 0x0014; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/40f6865f-559e-411a-b93c-4037f600776e


        const ushort REFERENCENAME = 0x0016; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/135dd749-c217-4d73-8549-b54e52e89945
        const ushort REFERENCECONTROL = 0x002F; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/d64485fa-8562-4726-9c5e-11e8f01a81c0
        const ushort REFERENCEORIGINAL = 0x0033; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/3ba66994-8c7a-4634-b2da-f9331ace6686
        const ushort REFERENCEREGISTERED = 0x000D; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/6c39388e-96f5-4b93-b90a-ae625a063fcf
        const ushort REFERENCEPROJECT = 0x000E; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/08280eb0-d628-495c-867f-5985ed020142

        const ushort PROJECTMODULES = 0x000F; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/93ec5c79-b87f-4f5b-95d8-c6ac12e09ec5
        const ushort PROJECTCOOKIE = 0x0013; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/5fc0e8fc-58f2-4fe1-ac9a-60268ad8740d
        const ushort MODULENAME = 0x0019; //https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/4918bdd5-df77-43c2-8ef3-2d13fda9dae6

        // Variables to store the parsed values
        static uint codePage = 1252; // default

        
        static string projectName = "";
        static uint projectSysKind = 0;
        static string projectDocString = "";
        static string projectDocStringUnicode = "";

        static string projectHelpFilePath = "";
        static string projectHelpFilePathTwo = "";

        static uint projectHelpContext = 0;

        static uint projectLibFlags = 0;

        static string projectVersion = "";
        static string projectConstants = "";
        static string projectConstantsUnicode = "";
        static byte[] projectReferences = null;
        static string projectModules = null;
        
        public enum ModuleType : ushort
        {
            StandardModule = 0x21, // Standard code module
            ClassDocOrFormModule = 0x22, // Everything else module
            Unknown = 0xFFFF
        }
        public class VbaModuleInfo
        {
            public string ModuleName { get; }
            public string StreamName { get; }
            public uint TextOffset { get; }
            public uint CodePage { get; }

            public ModuleType Type { get; }

            public string SaveFileExtension { get; set; }

            public VbaModuleInfo(string moduleName,
                string streamName,
                uint codePage,
                uint textOffset,
                ModuleType type = ModuleType.Unknown)
            {
                ModuleName = moduleName;
                StreamName = streamName;
                CodePage = codePage;
                TextOffset = textOffset;
                Type = type;
            }
        }

        public static Dictionary<string, VbaModuleInfo> GetModuleTextOffsetsFromDirStream(CfbStream dirStream)
        {
            if (dirStream == null || dirStream.Length < 2)
            {
                throw new ArgumentException("Invalid dir stream");
            }

            // --- Step 1: Read compressed bytes ---
            byte[] buffer = new byte[dirStream.Length];
            dirStream.Seek(0, SeekOrigin.Begin);
            dirStream.Read(buffer, 0, buffer.Length);

            string? projectName = null;

            Dictionary<string, VbaModuleInfo> moduleDict = new Dictionary<string, VbaModuleInfo>();

            try
            {
                byte[] allBytes = VbaCompression.Decompress(buffer);
                MemoryStream decompressedStream = new MemoryStream(allBytes);
                // Ensure position is at start
                decompressedStream.Position = 0;

                using var br = new BinaryReader(decompressedStream);

                // Read the PROJECTINFORMATION
                while (decompressedStream.Position < decompressedStream.Length)
                {
                    // Work way through the records
                    ushort recordTypeId = br.ReadUInt16(); // 2 bytes

                    if (recordTypeId == PROJECTSYSKIND)
                    {
                        uint recordSize = br.ReadUInt32(); // 4 bytes, MUST be 0x00000004.
                        // SysKind is a 4-byte unsigned integer
                        projectSysKind = br.ReadUInt32(); // 4 bytes, 0=16-bitWin, 1=32-bitWin, 3=Mac, 4=64-bitWin
                    }


                    else if (recordTypeId == PROJECTLCID)
                    {
                        uint recordSize = br.ReadUInt32(); // 4 bytes, MUST be 0x00000004.
                        //The LCID value for the VBA project.
                        uint theLcid = br.ReadUInt32(); // 4 bytes, MUST be 0x00000409 or 1033 decimal.
                    }


                    else if (recordTypeId == PROJECTLCIDINVOKE)
                    {
                        uint recordSize = br.ReadUInt32(); // 4 bytes, MUST be 0x00000004.
                        // LCID value used for Invoke calls.
                        uint LcidInvoke = br.ReadUInt32(); //4 bytes, MUST be 0x00000409 or 1033 decimal.
                    }

                    // All modules in project share the same code page
                    else if (recordTypeId == PROJECTCODEPAGE)
                    {
                        uint recordSize = br.ReadUInt32(); // 4 bytes, MUST be 0x00000002.
                        codePage = br.ReadUInt16(); // 2 bytes, Code page
                    }


                    else if (recordTypeId == PROJECTNAME)
                    {
                        uint SizeOfProjectName = br.ReadUInt32(); // 4 bytes, MUST be >= 1 and <= 128
                        byte[] projectNameBytes = br.ReadBytes((int)SizeOfProjectName); // Variable
                        projectName = Encoding.GetEncoding((int)codePage).GetString(projectNameBytes).TrimEnd('\0');
                    }


                    else if (recordTypeId == PROJECTDOCSTRING)
                    {
                        uint SizeOfDocString = br.ReadUInt32(); // 4 bytes, MUST be <= 200
                        byte[] docStringBytes = br.ReadBytes((int)SizeOfDocString); // Variable
                        projectDocString = Encoding.GetEncoding((int)codePage).GetString(docStringBytes).TrimEnd('\0');


                        uint reserved = br.ReadUInt16(); // 2 bytes MUST be 0x0040 or 64 decimal. MUST be ignored.
                        uint SizeOfDocStringUnicode = br.ReadUInt32(); // 4 bytes, MUST BE EVEN

                        // Read the DocStringUnicode variable
                        byte[] docStringUnicodeBytes = br.ReadBytes((int)SizeOfDocStringUnicode); // Variable
                        projectDocStringUnicode = Encoding.Unicode.GetString(docStringUnicodeBytes).TrimEnd('\0');
                    }


                    else if (recordTypeId == PROJECTHELPFILEPATH)
                    {
                        uint SizeOfHelpFile1 = br.ReadUInt32(); // 4 bytes, MUST be <= 260

                        byte[] helpFile1Bytes = br.ReadBytes((int)SizeOfHelpFile1); // Variable
                        projectHelpFilePath = Encoding.GetEncoding((int)codePage).GetString(helpFile1Bytes).TrimEnd('\0');


                        uint reserved = br.ReadUInt16(); // 2 bytes Reserved MUST be 0x003D or 61 decimal. MUST be ignored.
                        uint SizeOfHelpFile2  = br.ReadUInt32(); // 4 bytes SizeOfHelpFile2, must = SizeOfHelpFile1

                        //HelpFile2 (variable)
                        //br.ReadBytes((int)SizeOfHelpFile2);
                        byte[] helpFile2Bytes = br.ReadBytes((int)SizeOfHelpFile2); // Variable
                        projectHelpFilePathTwo = Encoding.GetEncoding((int)codePage).GetString(helpFile2Bytes).TrimEnd('\0');

                    }


                    else if (recordTypeId == PROJECTHELPCONTEXT)
                    {
                        uint recordSize = br.ReadUInt32(); // 4 bytes, MUST be 0x00000004.
                        projectHelpContext = br.ReadUInt32(); // 4 bytes
                    }


                    else if (recordTypeId == PROJECTLIBFLAGS)
                    {
                        uint recordSize = br.ReadUInt32(); // 4 bytes, MUST be 0x00000004.
                        projectLibFlags = br.ReadUInt32(); // 4 bytes, MUST be 0x00000000.
                    }


                    else if (recordTypeId == PROJECTVERSION)
                    {
                        uint reserved = br.ReadUInt32(); // 4 bytes, MUST be 0x00000004. MUST be ignored.
                        uint versionMajor = br.ReadUInt32(); // 4 bytes
                        ushort versionMinor = br.ReadUInt16(); // 2 bytes

                        projectVersion = $"{versionMajor}.{versionMinor}";
                    }

                    else if (recordTypeId == PROJECTCONSTANTS)
                    {
                        uint SizeOfConstants = br.ReadUInt32(); // 4 bytes, must be <= 1015

                        byte[] constantsBytes = br.ReadBytes((int)SizeOfConstants); // Variable
                        projectConstants = Encoding.GetEncoding((int)codePage).GetString(constantsBytes).TrimEnd('\0');


                        uint reserved = br.ReadUInt16(); // 2 bytes Reserved MUST be 0x003D or 61 decimal. MUST be ignored.
                        uint SizeOfConstantsUnicode = br.ReadUInt32(); // 4 bytes SizeOfHelpFile2, must = SizeOfHelpFile1

                        //HelpFile2 (variable)
                        //br.ReadBytes((int)SizeOfHelpFile2);
                        byte[] ConstantsUnicodeBytes = br.ReadBytes((int)SizeOfConstantsUnicode); // Variable
                        projectConstantsUnicode = Encoding.Unicode.GetString(ConstantsUnicodeBytes).TrimEnd('\0');

                        // This is the last record in the PROJECTINFORMATION
                        break; // Break so we can move on to the PROJECTREFERENCES
                    }

                    else
                    {
                        // We already read the ID, so read the size???
                        uint recordSize = br.ReadUInt32();
                        // Skip payload of records we're not interested in
                        decompressedStream.Seek(recordSize, SeekOrigin.Current);
                    }
                }


                // Read the ReferencesRecord
                while (decompressedStream.Position < decompressedStream.Length)
                {
                    // Work way through the records
                    ushort recordTypeId = br.ReadUInt16(); // 2 bytes


                    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/1cf3c0b7-71ca-41cb-83f8-6360181512e2
                    if (recordTypeId == 0x000F) // PROJECTMODULES
                    {
                        // Rewind 2 bytes so the next stage can read the record ID again
                        decompressedStream.Position -= 2;
                        break; // Exit references loop; now ready to read modules
                    }


                    else if (recordTypeId == REFERENCENAME)
                    {
                        uint SizeOfName = br.ReadUInt32(); // 4 bytes

                        byte[] nameBytes = br.ReadBytes((int)SizeOfName); // Variable
                        string refName = Encoding.GetEncoding((int)codePage).GetString(nameBytes).TrimEnd('\0');


                        uint reserved = br.ReadUInt16(); // 2 bytes, MUST be 0x003E or 62 decimal. MUST be ignored.
                        uint SizeOfNameUnicode = br.ReadUInt32(); // 4 bytes SizeOfNameUnicode

                        byte[] NameUnicodeBytes = br.ReadBytes((int)SizeOfNameUnicode); // Variable
                        string refNameUnicode = Encoding.Unicode.GetString(NameUnicodeBytes).TrimEnd('\0');
                    }

                    else if (recordTypeId == REFERENCECONTROL)
                    {
                        uint SizeTwiddled = br.ReadUInt32(); // 4 bytes
                        uint SizeOfLibidTwiddled = br.ReadUInt32(); // 4 bytes
                        byte[] LibidTwiddledBytes = br.ReadBytes((int)SizeOfLibidTwiddled); // Variable
                        string LibidTwiddled = Encoding.GetEncoding((int)codePage).GetString(LibidTwiddledBytes).TrimEnd('\0');

                        uint Reserved1 = br.ReadUInt32(); // (4 bytes): MUST be 0x00000000.MUST be ignored.
                        uint Reserved2 = br.ReadUInt16();  // (2 bytes): MUST be 0x0000.MUST be ignored.

                        //------------------
                        // REFERENCENAME record
                        ushort refNamelId = br.ReadUInt16(); // 2 bytes, should be 0x16 or 22 decimal
                        uint SizeOfName = br.ReadUInt32(); // 4 bytes

                        byte[] nameBytes = br.ReadBytes((int)SizeOfName); // Variable
                        string refName = Encoding.GetEncoding((int)codePage).GetString(nameBytes).TrimEnd('\0');


                        uint reserved = br.ReadUInt16(); // 2 bytes, MUST be 0x003E or 62 decimal. MUST be ignored.
                        uint SizeOfNameUnicode = br.ReadUInt32(); // 4 bytes SizeOfNameUnicode

                        byte[] NameUnicodeBytes = br.ReadBytes((int)SizeOfNameUnicode); // Variable
                        string refNameUnicode = Encoding.Unicode.GetString(NameUnicodeBytes).TrimEnd('\0');
                        //------------------

                        uint Reserved3 = br.ReadUInt16(); //(2 bytes): MUST be 0x0030 or 48 decimal. MUST be ignored.

                        uint SizeExtended = br.ReadUInt32(); // (4 bytes)
                        uint SizeOfLibidExtended = br.ReadUInt32(); // (4 bytes)
                        byte[] LibidExtendedBytes = br.ReadBytes((int)SizeOfLibidExtended); // Variable
                        string LibidExtended = Encoding.Unicode.GetString(LibidExtendedBytes).TrimEnd('\0');

                        uint Reserved4 = br.ReadUInt32(); // (4 bytes): MUST be 0x00000000.MUST be ignored.

                        uint Reserved5 = br.ReadUInt16(); // (2 bytes): MUST be 0x0000.MUST be ignored.

                        byte[] OriginalTypeLib = br.ReadBytes(16);

                        uint Cookie = br.ReadUInt32(); // (4 bytes)
                    }

                    else if (recordTypeId == REFERENCEORIGINAL)
                    {
                        uint SizeOfLibidOriginal = br.ReadUInt32(); // 4 bytes

                        byte[] LibidOriginalBytes = br.ReadBytes((int)SizeOfLibidOriginal); // Variable
                        string LibidOriginal = Encoding.GetEncoding((int)codePage).GetString(LibidOriginalBytes).TrimEnd('\0');

                        //---------------------------------------------------------
                        //---------------------------------------------------------
                        // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/d64485fa-8562-4726-9c5e-11e8f01a81c0
                        // What follows is a REFERENCECONTROL Record

                        ushort refcontrolId = br.ReadUInt16(); // 2 bytes
                        uint SizeTwiddled = br.ReadUInt32(); // 4 bytes
                        uint SizeOfLibidTwiddled = br.ReadUInt32(); // 4 bytes
                        byte[] LibidTwiddledBytes = br.ReadBytes((int)SizeOfLibidTwiddled); // Variable
                        string LibidTwiddled = Encoding.GetEncoding((int)codePage).GetString(LibidTwiddledBytes).TrimEnd('\0');

                        uint Reserved1 = br.ReadUInt32(); // (4 bytes): MUST be 0x00000000.MUST be ignored.
                        uint Reserved2 = br.ReadUInt16();  // (2 bytes): MUST be 0x0000.MUST be ignored.

                        //------------------
                        // REFERENCENAME record
                        ushort refNamelId = br.ReadUInt16(); // 2 bytes, should be 0x16 or 22 decimal
                        uint SizeOfName = br.ReadUInt32(); // 4 bytes

                        byte[] nameBytes = br.ReadBytes((int)SizeOfName); // Variable
                        string refName = Encoding.GetEncoding((int)codePage).GetString(nameBytes).TrimEnd('\0');


                        uint reserved = br.ReadUInt16(); // 2 bytes, MUST be 0x003E or 62 decimal. MUST be ignored.
                        uint SizeOfNameUnicode = br.ReadUInt32(); // 4 bytes SizeOfNameUnicode

                        byte[] NameUnicodeBytes = br.ReadBytes((int)SizeOfNameUnicode); // Variable
                        string refNameUnicode = Encoding.Unicode.GetString(NameUnicodeBytes).TrimEnd('\0');
                        //------------------

                        uint Reserved3 = br.ReadUInt16(); //(2 bytes): MUST be 0x0030 or 48 decimal. MUST be ignored.

                        uint SizeExtended = br.ReadUInt32(); // (4 bytes)
                        uint SizeOfLibidExtended = br.ReadUInt32(); // (4 bytes)
                        byte[] LibidExtendedBytes = br.ReadBytes((int)SizeOfLibidExtended); // Variable
                        string LibidExtended = Encoding.Unicode.GetString(LibidExtendedBytes).TrimEnd('\0');

                        uint Reserved4 = br.ReadUInt32(); // (4 bytes): MUST be 0x00000000.MUST be ignored.

                        uint Reserved5 = br.ReadUInt16(); // (2 bytes): MUST be 0x0000.MUST be ignored.

                        byte[] OriginalTypeLib = br.ReadBytes(16);

                        uint Cookie = br.ReadUInt32(); // (4 bytes)
                        //---------------------------------------------------------
                        //---------------------------------------------------------

                    }

                    else if (recordTypeId == REFERENCEREGISTERED)
                    {
                        uint SizeOfRefReg = br.ReadUInt32(); // 4 bytes
                        uint SizeOfLibid = br.ReadUInt32(); // 4 bytes

                        byte[] libidBytes = br.ReadBytes((int)SizeOfLibid); // Variable
                        string Libid = Encoding.GetEncoding((int)codePage).GetString(libidBytes).TrimEnd('\0');
                        // *\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINDOWS\system32\stdole2.tlb#OLE Automation

                        uint reserved1 = br.ReadUInt32(); // 4 bytes, MUST be 0x00000000. MUST be ignored.
                        uint reserved2 = br.ReadUInt16(); // 2 bytes, MUST be 0x0000. MUST be ignored.
                    }

                    else if (recordTypeId == REFERENCEPROJECT)
                    {
                        uint Size = br.ReadUInt32(); // 4 bytes
                        uint SizeOfLibidAbsolute = br.ReadUInt32(); // 4 bytes
                        byte[] LibidAbsoluteBytes = br.ReadBytes((int)SizeOfLibidAbsolute); // Variable
                        string LibidAbsolute = Encoding.GetEncoding((int)codePage).GetString(LibidAbsoluteBytes).TrimEnd('\0');
                        
                        uint SizeOfLibidRelative = br.ReadUInt32(); // 4 bytes
                        byte[] LibidRelativeBytes = br.ReadBytes((int)SizeOfLibidRelative); // Variable
                        string LibidRelative = Encoding.GetEncoding((int)codePage).GetString(LibidRelativeBytes).TrimEnd('\0');

                        uint MajorVersion = br.ReadUInt32(); // 4 bytes
                        uint MinorVersion = br.ReadUInt16(); // 2 bytes
                    }


                    else
                    {
                        Console.WriteLine("Unknown record Id: " + recordTypeId);
                    }

                }



                // Read the PROJECTMODULES
                while (decompressedStream.Position < decompressedStream.Length)
                {
                    // Work way through the records
                    ushort recordTypeId = br.ReadUInt16(); // 2 bytes


                    if (recordTypeId == PROJECTMODULES)
                    {
                        uint Size = br.ReadUInt32(); // 4 bytes
                        uint Count = br.ReadUInt16(); // 2 bytes, holds # of elements in Modules

                        //---------------------
                        // ProjectCookieRecord is a Cookie Record
                        ushort cookiedRecordId = br.ReadUInt16(); // 2 bytes
                        uint SizeForCookie = br.ReadUInt32(); // 4 bytes
                        ushort Cookie = br.ReadUInt16(); // 2 bytes, MUST be ignored on read. MUST be 0xFFFF on write.
                        //---------------------

                        // The number of elements in Modules
                        for (int i = 0; i < Count; i++)
                        {
                            //NameRecord(variable)-------------
                            ushort IdModuleName = br.ReadUInt16(); // 2 bytes, must be 0x0047
                            uint SizeOfModuleName = br.ReadUInt32(); // 4 bytes
                            byte[] ModuleNameBytes = br.ReadBytes((int)SizeOfModuleName); // Variable
                            string ModuleName = Encoding.GetEncoding((int)codePage).GetString(ModuleNameBytes).TrimEnd('\0');
                            //---------------------------------


                            //NameUnicodeRecord (variable)-----
                            ushort IdModuleNameUnicode = br.ReadUInt16(); // 2 bytes, must be 0x0019
                            uint SizeOfModuleNameUnicode = br.ReadUInt32(); // 4 bytes
                            byte[] ModuleNameUnicodeBytes = br.ReadBytes((int)SizeOfModuleNameUnicode); // Variable
                            string ModuleNameUnicode = Encoding.Unicode.GetString(ModuleNameUnicodeBytes).TrimEnd('\0');
                            //---------------------------------


                            //StreamNameRecord (variable)------
                            ushort IdModuleStreamName = br.ReadUInt16(); // 2 bytes, must be 0x001A.
                            uint SizeOfModuleStreamName = br.ReadUInt32(); // 4 bytes
                            byte[] ModuleStreamNameBytes = br.ReadBytes((int)SizeOfModuleStreamName); // Variable
                            string ModuleStreamName = Encoding.GetEncoding((int)codePage).GetString(ModuleStreamNameBytes).TrimEnd('\0');
                            
                            ushort ReservedModuleStreamName = br.ReadUInt16(); // (2 bytes): MUST be 0x0032.MUST be ignored.

                            uint SizeOfModuleStreamNameUnicode = br.ReadUInt32(); // 4 bytes
                            byte[] ModuleStreamNameUnicodeBytes = br.ReadBytes((int)SizeOfModuleStreamNameUnicode); // Variable
                            string ModuleStreamNameUnicode = Encoding.Unicode.GetString(ModuleStreamNameUnicodeBytes).TrimEnd('\0');
                            //---------------------------------


                            //DocStringRecord (variable)-------
                            ushort IdModuleDocString = br.ReadUInt16(); // 2 bytes, must be 0x001C
                            uint SizeOfModuleDocString = br.ReadUInt32(); // 4 bytes
                            byte[] ModuleDocStringBytes = br.ReadBytes((int)SizeOfModuleDocString); // Variable
                            string ModuleDocString = Encoding.GetEncoding((int)codePage).GetString(ModuleDocStringBytes).TrimEnd('\0');

                            ushort ReservedModuleDocString = br.ReadUInt16(); // (2 bytes): MUST be 0x0048 or 72 decimal.MUST be ignored.

                            uint SizeOfModuleDocStringUnicode = br.ReadUInt32(); // 4 bytes
                            byte[] ModuleDocStringUnicodeBytes = br.ReadBytes((int)SizeOfModuleDocStringUnicode); // Variable
                            string ModuleDocStringUnicode = Encoding.Unicode.GetString(ModuleDocStringUnicodeBytes).TrimEnd('\0');
                            //---------------------------------


                            //OffsetRecord---------------------
                            ushort IdModuleOffset = br.ReadUInt16(); // 2 bytes, must be 0x0031 or 49 decimal
                            uint SizeModuleOffset = br.ReadUInt32(); // (4 bytes): MUST be 0x00000004
                            //Contains the byte offset of the source code in the ModuleStream
                            uint ModuleTextOffset = br.ReadUInt32(); // (4 bytes)
                            //---------------------------------


                            //HelpContextRecord----------------
                            ushort IdModuleHelpContext = br.ReadUInt16(); // 2 bytes, must be 0x001E
                            uint SizeOfModuleHelpContext = br.ReadUInt32(); // (4 bytes): MUST be 0x00000004
                            //Help topic identifier in the Help file found in PROJECTHELPFILEPATH Record
                            uint ModuleHelpContext = br.ReadUInt32(); // (4 bytes)
                            //---------------------------------


                            //CookieRecord---------------------
                            ushort IdModuleCookie2 = br.ReadUInt16(); // 2 bytes
                            uint SizeOfModuleCookie2 = br.ReadUInt32(); // 4 bytes, MUST be 0x00000002.
                            ushort ModuleCookie2 = br.ReadUInt16(); // 2 bytes, MUST be ignored on read. MUST be 0xFFFF on write.
                            //---------------------------------


                            //TypeRecord-----------------------
                            ushort IdModuleType = br.ReadUInt16(); // 2 bytes
                            uint ReservedModuleType = br.ReadUInt32(); // 4 bytes, MUST be 0x00000000. MUST be ignored.
                            ModuleType finalType = IdModuleType switch
                            {
                                0x21 => ModuleType.StandardModule,
                                0x22 => ModuleType.ClassDocOrFormModule,
                                _ => ModuleType.Unknown
                            };
                            //---------------------------------


                            ushort NextIdOptional = br.ReadUInt16(); // 2 bytes
                            const ushort MODULEREADONLY = 0x0025;
                            //ReadOnlyRecord (optional)--------
                            if (NextIdOptional == MODULEREADONLY)
                            {
                                uint ReservedModuleReadOnly = br.ReadUInt32(); //MUST be 0x00000000. MUST be ignored.
                                NextIdOptional = br.ReadUInt16(); // 2 bytes
                            }
                            // PrivateRecord(optional)-------- -
                            const ushort MODULEPRIVATE = 0x0028;
                            if (NextIdOptional == MODULEPRIVATE)
                            {
                                uint ReservedModulePrivate = br.ReadUInt32(); //MUST be 0x00000000. MUST be ignored.
                                NextIdOptional = br.ReadUInt16(); // 2 bytes
                            }

                            //Terminator-----------------------
                            // The terminator must be reached meaning NextIdOptional hold terminator

                            // Save this modul into dictionary so we can return it later
                            if (!moduleDict.ContainsKey(ModuleStreamNameUnicode))
                            {
                                var vbaModuleInfo = new VbaModuleInfo(
                                    ModuleNameUnicode,
                                    ModuleStreamNameUnicode,
                                    codePage,
                                    ModuleTextOffset,
                                    finalType);

                                moduleDict.Add(ModuleStreamNameUnicode, vbaModuleInfo);
                            }

                            //Reserved-------------------------
                            uint ReservedFinal = br.ReadUInt32(); // 4 bytes, MUST be 0x00000000. MUST be ignored.
                        }


                    }

                    if (recordTypeId == PROJECTCOOKIE)
                    {
                        uint Size = br.ReadUInt32(); // 4 bytes
                        uint Cookie = br.ReadUInt16(); // 2 bytes, MUST be ignored on read. MUST be 0xFFFF on write.

                    }

                }


            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading dir stream: {ex.Message}");
               
            }

            return moduleDict;
        }
        


    }
}
