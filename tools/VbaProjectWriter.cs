// VbaProjectWriter.cs
// Builds xl/vbaProject.bin from VBA source files (.bas / .cls) and puts it into an .xlsx package,
// so that the .xlsm can be built without "Trust access to the VBA project object model".
// Formats: [MS-CFB] (compound file) and [MS-OVBA] (VBA project). Only the source text is stored;
// Excel compiles it when the workbook is opened (the _VBA_PROJECT stream has no performance cache).
// This file is loaded by tools/build.ps1 with Add-Type (C# 5, .NET Framework 4.x). Keep it ASCII only.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace BatchInvoicePdf.Build
{
    public enum VbaModuleKind { Standard, Class, Document }

    public class VbaModule
    {
        public string Name;
        public VbaModuleKind Kind;
        public string Code; // source as stored in the module stream (Attribute lines + code), CRLF
    }

    // ---------------------------------------------------------------- [MS-OVBA] 2.4.1 compression
    public static class OvbaCompression
    {
        const int ChunkSize = 4096;

        static int BitCount(int difference)
        {
            int bitCount = 4;
            while ((1 << bitCount) < difference) bitCount++;
            return bitCount;
        }

        public static byte[] Compress(byte[] data)
        {
            var output = new List<byte>(data.Length / 2 + 16);
            output.Add(0x01);
            int pos = 0;
            while (pos < data.Length)
            {
                int chunkStart = pos;
                int chunkEnd = Math.Min(pos + ChunkSize, data.Length);
                var chunk = new List<byte>(ChunkSize + 2);
                int p = chunkStart;
                while (p < chunkEnd)
                {
                    int flagIndex = chunk.Count;
                    chunk.Add(0);
                    int flags = 0;
                    for (int bit = 0; bit < 8 && p < chunkEnd; bit++)
                    {
                        int bitCount = BitCount(p - chunkStart);
                        int maxLength = (0xFFFF >> bitCount) + 3;
                        int bestLength = 0, bestOffset = 0;
                        for (int c = p - 1; c >= chunkStart; c--)
                        {
                            int len = 0;
                            while (p + len < chunkEnd && len < maxLength && data[c + len] == data[p + len]) len++;
                            if (len > bestLength) { bestLength = len; bestOffset = p - c; }
                            if (bestLength == maxLength) break;
                        }
                        if (bestLength >= 3)
                        {
                            int token = ((bestOffset - 1) << (16 - bitCount)) | (bestLength - 3);
                            chunk.Add((byte)(token & 0xFF));
                            chunk.Add((byte)((token >> 8) & 0xFF));
                            flags |= 1 << bit;
                            p += bestLength;
                        }
                        else
                        {
                            chunk.Add(data[p]);
                            p++;
                        }
                    }
                    chunk[flagIndex] = (byte)flags;
                }

                if (chunk.Count > ChunkSize)
                {
                    // Raw chunk: allowed only for a full 4096-byte chunk.
                    if (chunkEnd - chunkStart != ChunkSize)
                        throw new InvalidOperationException("OVBA compression: partial chunk did not compress.");
                    int header = 0x0FFF | (0x3 << 12);
                    output.Add((byte)(header & 0xFF));
                    output.Add((byte)(header >> 8));
                    for (int i = chunkStart; i < chunkEnd; i++) output.Add(data[i]);
                }
                else
                {
                    int header = ((chunk.Count + 2 - 3) & 0x0FFF) | (0x3 << 12) | 0x8000;
                    output.Add((byte)(header & 0xFF));
                    output.Add((byte)(header >> 8));
                    output.AddRange(chunk);
                }
                pos = chunkEnd;
            }
            return output.ToArray();
        }

        public static byte[] Decompress(byte[] data)
        {
            if (data.Length == 0 || data[0] != 0x01) throw new InvalidDataException("OVBA: bad signature byte.");
            var output = new List<byte>(data.Length * 2);
            int pos = 1;
            while (pos < data.Length)
            {
                int header = data[pos] | (data[pos + 1] << 8);
                int chunkEnd = pos + (header & 0x0FFF) + 3;
                bool compressed = (header & 0x8000) != 0;
                pos += 2;
                int chunkStart = output.Count;
                if (!compressed)
                {
                    for (int i = 0; i < ChunkSize; i++) output.Add(data[pos + i]);
                    pos += ChunkSize;
                    continue;
                }
                while (pos < chunkEnd)
                {
                    int flags = data[pos++];
                    for (int bit = 0; bit < 8 && pos < chunkEnd; bit++)
                    {
                        if (((flags >> bit) & 1) == 0)
                        {
                            output.Add(data[pos++]);
                        }
                        else
                        {
                            int token = data[pos] | (data[pos + 1] << 8);
                            pos += 2;
                            int bitCount = BitCount(output.Count - chunkStart);
                            int lengthMask = 0xFFFF >> bitCount;
                            int length = (token & lengthMask) + 3;
                            int offset = (token >> (16 - bitCount)) + 1;
                            int src = output.Count - offset;
                            for (int i = 0; i < length; i++) output.Add(output[src + i]);
                        }
                    }
                }
            }
            return output.ToArray();
        }
    }

    // ---------------------------------------------------------------- [MS-CFB] compound file writer
    public class CfbNode
    {
        public string Name;
        public bool IsStorage;
        public byte[] Data = new byte[0];
        public List<CfbNode> Children = new List<CfbNode>();

        internal int Id;
        internal int Left = -1, Right = -1, Child = -1;
        internal bool Red;
        internal uint StartSector;

        public static CfbNode Storage(string name) { return new CfbNode { Name = name, IsStorage = true }; }
        public static CfbNode Stream(string name, byte[] data) { return new CfbNode { Name = name, Data = data }; }
    }

    public static class CompoundFile
    {
        const int SectorSize = 512;
        const int MiniSectorSize = 64;
        const int MiniStreamCutoff = 4096;
        const uint FreeSect = 0xFFFFFFFF, EndOfChain = 0xFFFFFFFE, FatSect = 0xFFFFFFFD, NoStream = 0xFFFFFFFF;

        static int CompareNames(string a, string b)
        {
            if (a.Length != b.Length) return a.Length.CompareTo(b.Length);
            return string.CompareOrdinal(a.ToUpperInvariant(), b.ToUpperInvariant());
        }

        static void Flatten(CfbNode node, List<CfbNode> all)
        {
            node.Id = all.Count;
            all.Add(node);
            foreach (var child in node.Children) Flatten(child, all);
        }

        // Balanced tree from the sorted siblings. Nodes on the deepest level are red when the tree is
        // not perfect; this keeps the red-black rules (same black count on every path).
        static int BuildTree(List<CfbNode> sorted, int lo, int hi, int depth, Dictionary<CfbNode, int> depths)
        {
            if (lo > hi) return -1;
            int mid = (lo + hi) / 2;
            var node = sorted[mid];
            depths[node] = depth;
            node.Left = BuildTree(sorted, lo, mid - 1, depth + 1, depths);
            node.Right = BuildTree(sorted, mid + 1, hi, depth + 1, depths);
            return node.Id;
        }

        static void LinkChildren(CfbNode storage)
        {
            if (storage.Children.Count == 0) return;
            var sorted = new List<CfbNode>(storage.Children);
            sorted.Sort((x, y) => CompareNames(x.Name, y.Name));
            var depths = new Dictionary<CfbNode, int>();
            storage.Child = BuildTree(sorted, 0, sorted.Count - 1, 0, depths);
            int maxDepth = 0;
            foreach (var d in depths.Values) maxDepth = Math.Max(maxDepth, d);
            bool perfect = sorted.Count == (1 << (maxDepth + 1)) - 1;
            foreach (var n in sorted) n.Red = !perfect && depths[n] == maxDepth;
            foreach (var n in sorted) if (n.IsStorage) LinkChildren(n);
        }

        static void PutU32(byte[] buf, int offset, uint value)
        {
            buf[offset] = (byte)value; buf[offset + 1] = (byte)(value >> 8);
            buf[offset + 2] = (byte)(value >> 16); buf[offset + 3] = (byte)(value >> 24);
        }

        static void PutU16(byte[] buf, int offset, int value)
        {
            buf[offset] = (byte)value; buf[offset + 1] = (byte)(value >> 8);
        }

        public static byte[] Write(CfbNode root)
        {
            var all = new List<CfbNode>();
            Flatten(root, all);
            LinkChildren(root);

            // Mini stream (streams smaller than 4096 bytes) and its allocation table.
            var mini = new MemoryStream();
            var miniFat = new List<uint>();
            var regular = new List<CfbNode>();
            foreach (var n in all)
            {
                if (n == root || n.IsStorage) continue;
                if (n.Data.Length == 0) { n.StartSector = EndOfChain; continue; }
                if (n.Data.Length >= MiniStreamCutoff) { regular.Add(n); continue; }
                int count = (n.Data.Length + MiniSectorSize - 1) / MiniSectorSize;
                n.StartSector = (uint)miniFat.Count;
                for (int i = 0; i < count; i++)
                    miniFat.Add(i == count - 1 ? EndOfChain : (uint)(miniFat.Count + 1));
                mini.Write(n.Data, 0, n.Data.Length);
                int pad = count * MiniSectorSize - n.Data.Length;
                mini.Write(new byte[pad], 0, pad);
            }
            byte[] miniBytes = mini.ToArray();

            // Sector layout: regular streams, mini stream, mini FAT, directory, FAT.
            var sectors = new List<byte[]>();
            var fat = new List<uint>();
            Func<byte[], uint> place = bytes =>
            {
                if (bytes.Length == 0) return EndOfChain;
                int count = (bytes.Length + SectorSize - 1) / SectorSize;
                uint start = (uint)sectors.Count;
                for (int i = 0; i < count; i++)
                {
                    var s = new byte[SectorSize];
                    Array.Copy(bytes, i * SectorSize, s, 0, Math.Min(SectorSize, bytes.Length - i * SectorSize));
                    sectors.Add(s);
                    fat.Add(i == count - 1 ? EndOfChain : (uint)(sectors.Count));
                }
                return start;
            };

            foreach (var n in regular) n.StartSector = place(n.Data);
            uint miniStreamStart = place(miniBytes);

            var miniFatBytes = new byte[miniFat.Count * 4];
            for (int i = 0; i < miniFat.Count; i++) PutU32(miniFatBytes, i * 4, miniFat[i]);
            if (miniFatBytes.Length % SectorSize != 0)
            {
                var padded = new byte[(miniFatBytes.Length / SectorSize + 1) * SectorSize];
                for (int i = 0; i < padded.Length; i += 4) PutU32(padded, i, FreeSect);
                Array.Copy(miniFatBytes, padded, miniFatBytes.Length);
                miniFatBytes = padded;
            }
            uint miniFatStart = place(miniFatBytes);
            int miniFatSectors = miniFatBytes.Length / SectorSize;

            int dirEntries = ((all.Count + 3) / 4) * 4;
            var dir = new byte[dirEntries * 128];
            for (int i = 0; i < dirEntries; i++)
            {
                int o = i * 128;
                if (i >= all.Count)
                {
                    PutU32(dir, o + 68, NoStream); PutU32(dir, o + 72, NoStream); PutU32(dir, o + 76, NoStream);
                    continue;
                }
                var n = all[i];
                var nameBytes = Encoding.Unicode.GetBytes(n.Name);
                if (nameBytes.Length > 62) throw new InvalidOperationException("CFB: name too long: " + n.Name);
                Array.Copy(nameBytes, 0, dir, o, nameBytes.Length);
                PutU16(dir, o + 64, nameBytes.Length + 2);
                dir[o + 66] = (byte)(n == root ? 5 : (n.IsStorage ? 1 : 2));
                dir[o + 67] = (byte)(n.Red ? 0 : 1);
                PutU32(dir, o + 68, n.Left < 0 ? NoStream : (uint)n.Left);
                PutU32(dir, o + 72, n.Right < 0 ? NoStream : (uint)n.Right);
                PutU32(dir, o + 76, n.Child < 0 ? NoStream : (uint)n.Child);
                if (n == root)
                {
                    PutU32(dir, o + 116, miniStreamStart);
                    PutU32(dir, o + 120, (uint)miniBytes.Length);
                }
                else if (!n.IsStorage)
                {
                    PutU32(dir, o + 116, n.StartSector);
                    PutU32(dir, o + 120, (uint)n.Data.Length);
                }
            }
            uint dirStart = place(dir);

            int nonFat = sectors.Count;
            int fatSectors = 1;
            while ((nonFat + fatSectors) > fatSectors * (SectorSize / 4)) fatSectors++;
            if (fatSectors > 109) throw new InvalidOperationException("CFB: file too large for this writer.");
            var fatSectorIds = new List<uint>();
            for (int i = 0; i < fatSectors; i++) { fatSectorIds.Add((uint)(nonFat + i)); fat.Add(FatSect); }
            var fatBytes = new byte[fatSectors * SectorSize];
            for (int i = 0; i < fatBytes.Length / 4; i++) PutU32(fatBytes, i * 4, i < fat.Count ? fat[i] : FreeSect);

            var header = new byte[SectorSize];
            byte[] signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
            Array.Copy(signature, header, 8);
            PutU16(header, 0x18, 0x003E);
            PutU16(header, 0x1A, 0x0003);
            PutU16(header, 0x1C, 0xFFFE);
            PutU16(header, 0x1E, 0x0009);
            PutU16(header, 0x20, 0x0006);
            PutU32(header, 0x28, 0);
            PutU32(header, 0x2C, (uint)fatSectors);
            PutU32(header, 0x30, dirStart);
            PutU32(header, 0x34, 0);
            PutU32(header, 0x38, MiniStreamCutoff);
            PutU32(header, 0x3C, miniFatSectors == 0 ? EndOfChain : miniFatStart);
            PutU32(header, 0x40, (uint)miniFatSectors);
            PutU32(header, 0x44, EndOfChain);
            PutU32(header, 0x48, 0);
            for (int i = 0; i < 109; i++) PutU32(header, 0x4C + i * 4, i < fatSectorIds.Count ? fatSectorIds[i] : FreeSect);

            var result = new MemoryStream();
            result.Write(header, 0, header.Length);
            foreach (var s in sectors) result.Write(s, 0, s.Length);
            result.Write(fatBytes, 0, fatBytes.Length);
            return result.ToArray();
        }
    }

    // ---------------------------------------------------------------- [MS-OVBA] VBA project
    public static class VbaProjectWriter
    {
        const string ClassBase = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}";
        const string WorkbookBase = "0{00020819-0000-0000-C000-000000000046}";
        const string WorksheetBase = "0{00020820-0000-0000-C000-000000000046}";
        const string StdoleLibid = "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation";

        // Reads a source file as UTF-8 when it is valid UTF-8, otherwise as the project code page
        // (the VBA editor exports Shift_JIS on Japanese Windows).
        public static string ReadSource(string path, int codePage)
        {
            byte[] bytes = File.ReadAllBytes(path);
            try
            {
                var utf8 = new UTF8Encoding(false, true);
                string text = utf8.GetString(bytes);
                return text.Length > 0 && text[0] == '\uFEFF' ? text.Substring(1) : text;
            }
            catch (DecoderFallbackException)
            {
                return Encoding.GetEncoding(codePage).GetString(bytes);
            }
        }

        static string NormalizeNewLines(string text)
        {
            text = text.Replace("\r\n", "\n").Replace("\r", "\n").TrimEnd('\n') + "\n";
            return text.Replace("\n", "\r\n");
        }

        static string AttributeValue(string line)
        {
            int eq = line.IndexOf('=');
            return eq < 0 ? "" : line.Substring(eq + 1).Trim();
        }

        // .bas: stored as is. .cls: the VERSION/BEGIN..END header of an exported class is removed and
        // the attributes that the editor adds for class modules (VB_Base etc.) are put back.
        public static VbaModule LoadModule(string path, int codePage)
        {
            string text = NormalizeNewLines(ReadSource(path, codePage));
            var lines = new List<string>(text.Substring(0, text.Length - 2).Split(new[] { "\r\n" }, StringSplitOptions.None));
            bool isClass = Path.GetExtension(path).Equals(".cls", StringComparison.OrdinalIgnoreCase);

            int i = 0;
            if (isClass && lines.Count > 0 && lines[0].StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase))
            {
                i = 1;
                if (i < lines.Count && lines[i].Trim().Equals("BEGIN", StringComparison.OrdinalIgnoreCase))
                {
                    while (i < lines.Count && !lines[i].Trim().Equals("END", StringComparison.OrdinalIgnoreCase)) i++;
                    i++;
                }
            }
            var attrs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var order = new List<string>();
            while (i < lines.Count && lines[i].StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
            {
                string key = lines[i].Substring(10).Split('=')[0].Trim();
                if (!attrs.ContainsKey(key)) order.Add(key);
                attrs[key] = AttributeValue(lines[i]);
                i++;
            }
            if (!attrs.ContainsKey("VB_Name")) throw new InvalidDataException("No Attribute VB_Name in " + path);
            string name = attrs["VB_Name"].Trim('"');

            var sb = new StringBuilder();
            if (isClass)
            {
                string[] classOrder = { "VB_Name", "VB_Base", "VB_GlobalNameSpace", "VB_Creatable", "VB_PredeclaredId", "VB_Exposed", "VB_TemplateDerived", "VB_Customizable" };
                var defaults = new Dictionary<string, string> {
                    { "VB_Base", "\"" + ClassBase + "\"" }, { "VB_GlobalNameSpace", "False" }, { "VB_Creatable", "False" },
                    { "VB_PredeclaredId", "False" }, { "VB_Exposed", "False" }, { "VB_TemplateDerived", "False" }, { "VB_Customizable", "False" } };
                foreach (var key in classOrder)
                {
                    string value = attrs.ContainsKey(key) ? attrs[key] : defaults[key];
                    sb.Append("Attribute ").Append(key).Append(" = ").Append(value).Append("\r\n");
                }
                foreach (var key in order)
                    if (Array.IndexOf(classOrder, key) < 0)
                        sb.Append("Attribute ").Append(key).Append(" = ").Append(attrs[key]).Append("\r\n");
            }
            else
            {
                foreach (var key in order) sb.Append("Attribute ").Append(key).Append(" = ").Append(attrs[key]).Append("\r\n");
            }
            for (; i < lines.Count; i++) sb.Append(lines[i]).Append("\r\n");

            return new VbaModule { Name = name, Kind = isClass ? VbaModuleKind.Class : VbaModuleKind.Standard, Code = sb.ToString() };
        }

        public static VbaModule DocumentModule(string name, bool isWorkbook)
        {
            var sb = new StringBuilder();
            sb.Append("Attribute VB_Name = \"").Append(name).Append("\"\r\n");
            sb.Append("Attribute VB_Base = \"").Append(isWorkbook ? WorkbookBase : WorksheetBase).Append("\"\r\n");
            sb.Append("Attribute VB_GlobalNameSpace = False\r\n");
            sb.Append("Attribute VB_Creatable = False\r\n");
            sb.Append("Attribute VB_PredeclaredId = True\r\n");
            sb.Append("Attribute VB_Exposed = True\r\n");
            sb.Append("Attribute VB_TemplateDerived = False\r\n");
            sb.Append("Attribute VB_Customizable = True\r\n");
            return new VbaModule { Name = name, Kind = VbaModuleKind.Document, Code = sb.ToString() };
        }

        // [MS-OVBA] 2.4.3.2 data encryption, used for CMG / DPB / GC in the PROJECT stream.
        static string Encrypt(byte[] data, string projectId, byte seed)
        {
            byte projKey = 0;
            foreach (char ch in projectId) projKey = (byte)(projKey + (byte)ch);
            var output = new List<byte>();
            byte versionEnc = (byte)(seed ^ 2);
            byte projKeyEnc = (byte)(seed ^ projKey);
            output.Add(seed); output.Add(versionEnc); output.Add(projKeyEnc);
            byte unencrypted1 = projKey, encrypted1 = projKeyEnc, encrypted2 = versionEnc;
            var plain = new List<byte>();
            int ignoredLength = (seed & 6) / 2;
            for (int i = 0; i < ignoredLength; i++) plain.Add(0);
            plain.AddRange(BitConverter.GetBytes((uint)data.Length));
            plain.AddRange(data);
            foreach (byte b in plain)
            {
                byte enc = (byte)(b ^ (byte)(encrypted2 + unencrypted1));
                output.Add(enc);
                encrypted2 = encrypted1;
                encrypted1 = enc;
                unencrypted1 = b;
            }
            var sb = new StringBuilder();
            foreach (byte b in output) sb.Append(b.ToString("X2"));
            return sb.ToString();
        }

        static void Record(BinaryWriter w, ushort id, byte[] data)
        {
            w.Write(id); w.Write((uint)data.Length); w.Write(data);
        }

        static void RecordWithUnicode(BinaryWriter w, ushort id, ushort unicodeId, string text, Encoding mbcs)
        {
            Record(w, id, mbcs.GetBytes(text));
            Record(w, unicodeId, Encoding.Unicode.GetBytes(text));
        }

        static byte[] DirStream(List<VbaModule> modules, Encoding mbcs, int codePage)
        {
            var ms = new MemoryStream();
            var w = new BinaryWriter(ms);
            // PROJECTINFORMATION
            Record(w, 0x0001, BitConverter.GetBytes((uint)3));          // SysKind: 64-bit Windows
            Record(w, 0x0002, BitConverter.GetBytes((uint)0x409));      // Lcid
            Record(w, 0x0014, BitConverter.GetBytes((uint)0x409));      // LcidInvoke
            Record(w, 0x0003, BitConverter.GetBytes((ushort)codePage)); // CodePage
            Record(w, 0x0004, mbcs.GetBytes("VBAProject"));             // Name
            RecordWithUnicode(w, 0x0005, 0x0040, "", mbcs);             // DocString
            RecordWithUnicode(w, 0x0006, 0x003D, "", mbcs);             // HelpFilePath
            Record(w, 0x0007, BitConverter.GetBytes((uint)0));          // HelpContext
            Record(w, 0x0008, BitConverter.GetBytes((uint)0));          // LibFlags
            w.Write((ushort)0x0009); w.Write((uint)4); w.Write((uint)1); w.Write((ushort)0); // Version
            RecordWithUnicode(w, 0x000C, 0x003C, "", mbcs);             // Constants
            // PROJECTREFERENCES: OLE Automation only (VBA and Excel come from the host)
            RecordWithUnicode(w, 0x0016, 0x003E, "stdole", mbcs);
            byte[] libid = mbcs.GetBytes(StdoleLibid);
            w.Write((ushort)0x000D); w.Write((uint)(4 + libid.Length + 4 + 2)); w.Write((uint)libid.Length); w.Write(libid); w.Write((uint)0); w.Write((ushort)0);
            // PROJECTMODULES
            w.Write((ushort)0x000F); w.Write((uint)2); w.Write((ushort)modules.Count);
            w.Write((ushort)0x0013); w.Write((uint)2); w.Write((ushort)0xFFFF);
            foreach (var m in modules)
            {
                Record(w, 0x0019, mbcs.GetBytes(m.Name));
                Record(w, 0x0047, Encoding.Unicode.GetBytes(m.Name));
                RecordWithUnicode(w, 0x001A, 0x0032, m.Name, mbcs);     // stream name = module name
                RecordWithUnicode(w, 0x001C, 0x0048, "", mbcs);         // DocString
                Record(w, 0x0031, BitConverter.GetBytes((uint)0));      // source starts at offset 0
                Record(w, 0x001E, BitConverter.GetBytes((uint)0));      // HelpContext
                Record(w, 0x002C, BitConverter.GetBytes((ushort)0xFFFF));
                // MODULETYPE: 0x0021 = standard module, 0x0022 = document or class module
                w.Write((ushort)(m.Kind == VbaModuleKind.Standard ? 0x0021 : 0x0022)); w.Write((uint)0);
                w.Write((ushort)0x002B); w.Write((uint)0);
            }
            w.Write((ushort)0x0010); w.Write((uint)0);
            w.Flush();
            return ms.ToArray();
        }

        public static byte[] Build(List<VbaModule> modules, int codePage, string projectGuid)
        {
            Encoding mbcs = Encoding.GetEncoding(codePage);
            string projectId = "{" + projectGuid.Trim('{', '}').ToUpperInvariant() + "}";

            var project = new StringBuilder();
            project.Append("ID=\"").Append(projectId).Append("\"\r\n");
            foreach (var m in modules)
            {
                if (m.Kind == VbaModuleKind.Document) project.Append("Document=").Append(m.Name).Append("/&H00000000\r\n");
                else if (m.Kind == VbaModuleKind.Class) project.Append("Class=").Append(m.Name).Append("\r\n");
                else project.Append("Module=").Append(m.Name).Append("\r\n");
            }
            project.Append("Name=\"VBAProject\"\r\n");
            project.Append("HelpContextID=\"0\"\r\n");
            project.Append("VersionCompatible32=\"393222000\"\r\n");
            project.Append("CMG=\"").Append(Encrypt(new byte[] { 0, 0, 0, 0 }, projectId, 0x35)).Append("\"\r\n");
            project.Append("DPB=\"").Append(Encrypt(new byte[] { 0 }, projectId, 0x37)).Append("\"\r\n");
            project.Append("GC=\"").Append(Encrypt(new byte[] { 0xFF }, projectId, 0x39)).Append("\"\r\n");
            project.Append("\r\n[Host Extender Info]\r\n");
            project.Append("&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n");

            var wm = new MemoryStream();
            foreach (var m in modules)
            {
                byte[] a = mbcs.GetBytes(m.Name); wm.Write(a, 0, a.Length); wm.WriteByte(0);
                byte[] u = Encoding.Unicode.GetBytes(m.Name); wm.Write(u, 0, u.Length); wm.WriteByte(0); wm.WriteByte(0);
            }
            wm.WriteByte(0); wm.WriteByte(0);

            var root = CfbNode.Storage("Root Entry");
            var vba = CfbNode.Storage("VBA");
            root.Children.Add(vba);
            root.Children.Add(CfbNode.Stream("PROJECT", mbcs.GetBytes(project.ToString())));
            root.Children.Add(CfbNode.Stream("PROJECTwm", wm.ToArray()));
            vba.Children.Add(CfbNode.Stream("_VBA_PROJECT", new byte[] { 0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00 }));
            byte[] dir = DirStream(modules, mbcs, codePage);
            vba.Children.Add(CfbNode.Stream("dir", Compress(dir)));
            foreach (var m in modules)
            {
                byte[] source = mbcs.GetBytes(m.Code);
                if (mbcs.GetString(source) != m.Code)
                    throw new InvalidDataException("Module " + m.Name + " has characters that code page " + codePage + " cannot store.");
                vba.Children.Add(CfbNode.Stream(m.Name, Compress(source)));
            }
            return CompoundFile.Write(root);
        }

        // Compress and check that it comes back the same (guards the compressor).
        static byte[] Compress(byte[] data)
        {
            byte[] packed = OvbaCompression.Compress(data);
            byte[] back = OvbaCompression.Decompress(packed);
            if (back.Length != data.Length) throw new InvalidOperationException("OVBA round trip failed (length).");
            for (int i = 0; i < data.Length; i++)
                if (back[i] != data[i]) throw new InvalidOperationException("OVBA round trip failed at " + i + ".");
            return packed;
        }
    }

    // ---------------------------------------------------------------- .xlsx -> .xlsm
    public static class XlsmPackager
    {
        const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string NsPkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string NsTypes = "http://schemas.openxmlformats.org/package/2006/content-types";

        static XmlDocument Load(ZipArchive zip, string entryName)
        {
            var entry = zip.GetEntry(entryName);
            if (entry == null) throw new InvalidDataException("Missing part: " + entryName);
            var doc = new XmlDocument { PreserveWhitespace = true };
            using (var s = entry.Open()) doc.Load(s);
            return doc;
        }

        static void Save(ZipArchive zip, string entryName, XmlDocument doc)
        {
            var old = zip.GetEntry(entryName);
            if (old != null) old.Delete();
            var entry = zip.CreateEntry(entryName, CompressionLevel.Optimal);
            using (var s = entry.Open())
            using (var w = XmlWriter.Create(s, new XmlWriterSettings { Encoding = new UTF8Encoding(false) }))
                doc.Save(w);
        }

        static void SetCodeName(XmlDocument doc, string prElementName, string codeName)
        {
            var root = doc.DocumentElement;
            XmlElement pr = null;
            foreach (XmlNode n in root.ChildNodes)
                if (n is XmlElement && n.LocalName == prElementName && n.NamespaceURI == NsMain) { pr = (XmlElement)n; break; }
            if (pr == null)
            {
                pr = doc.CreateElement(prElementName, NsMain);
                XmlNode before = null;
                if (prElementName == "workbookPr")
                {
                    // CT_Workbook order: fileVersion, fileSharing, workbookPr, ... (elements of other
                    // namespaces, such as mc:AlternateContent, are skipped)
                    foreach (XmlNode n in root.ChildNodes)
                        if (n is XmlElement && n.NamespaceURI == NsMain && n.LocalName != "fileVersion" && n.LocalName != "fileSharing") { before = n; break; }
                }
                else
                {
                    foreach (XmlNode n in root.ChildNodes) if (n is XmlElement) { before = n; break; }
                }
                if (before != null) root.InsertBefore(pr, before); else root.AppendChild(pr);
            }
            pr.SetAttribute("codeName", codeName);
        }

        // sheetCodeNames: sheet name (as shown on the tab) -> code name
        public static void Inject(string xlsxPath, string xlsmPath, byte[] vbaProject, string workbookCodeName, IDictionary<string, string> sheetCodeNames)
        {
            File.Copy(xlsxPath, xlsmPath, true);
            using (var zip = ZipFile.Open(xlsmPath, ZipArchiveMode.Update))
            {
                var types = Load(zip, "[Content_Types].xml");
                var tns = new XmlNamespaceManager(types.NameTable);
                tns.AddNamespace("t", NsTypes);
                var wbType = (XmlElement)types.SelectSingleNode("/t:Types/t:Override[@PartName='/xl/workbook.xml']", tns);
                wbType.SetAttribute("ContentType", "application/vnd.ms-excel.sheet.macroEnabled.main+xml");
                if (types.SelectSingleNode("/t:Types/t:Override[@PartName='/xl/vbaProject.bin']", tns) == null)
                {
                    var o = types.CreateElement("Override", NsTypes);
                    o.SetAttribute("PartName", "/xl/vbaProject.bin");
                    o.SetAttribute("ContentType", "application/vnd.ms-office.vbaProject");
                    types.DocumentElement.AppendChild(o);
                }
                Save(zip, "[Content_Types].xml", types);

                var rels = Load(zip, "xl/_rels/workbook.xml.rels");
                var rns = new XmlNamespaceManager(rels.NameTable);
                rns.AddNamespace("r", NsPkgRel);
                var targets = new Dictionary<string, string>();
                var ids = new HashSet<string>();
                foreach (XmlElement r in rels.SelectNodes("/r:Relationships/r:Relationship", rns))
                {
                    ids.Add(r.GetAttribute("Id"));
                    targets[r.GetAttribute("Id")] = r.GetAttribute("Target");
                }
                string vbaId = "rIdVba";
                while (ids.Contains(vbaId)) vbaId += "1";
                var rel = rels.CreateElement("Relationship", NsPkgRel);
                rel.SetAttribute("Id", vbaId);
                rel.SetAttribute("Type", "http://schemas.microsoft.com/office/2006/relationships/vbaProject");
                rel.SetAttribute("Target", "vbaProject.bin");
                rels.DocumentElement.AppendChild(rel);
                Save(zip, "xl/_rels/workbook.xml.rels", rels);

                var wb = Load(zip, "xl/workbook.xml");
                SetCodeName(wb, "workbookPr", workbookCodeName);
                var wns = new XmlNamespaceManager(wb.NameTable);
                wns.AddNamespace("m", NsMain);
                var found = new HashSet<string>();
                foreach (XmlElement sheet in wb.SelectNodes("/m:workbook/m:sheets/m:sheet", wns))
                {
                    string sheetName = sheet.GetAttribute("name");
                    if (!sheetCodeNames.ContainsKey(sheetName)) continue;
                    string target = targets[sheet.GetAttribute("id", NsRel)];
                    string part = target.StartsWith("/") ? target.Substring(1) : "xl/" + target;
                    var ws = Load(zip, part);
                    SetCodeName(ws, "sheetPr", sheetCodeNames[sheetName]);
                    Save(zip, part, ws);
                    found.Add(sheetName);
                }
                foreach (var name in sheetCodeNames.Keys)
                    if (!found.Contains(name)) throw new InvalidDataException("Sheet not found in workbook: " + name);
                Save(zip, "xl/workbook.xml", wb);

                var vbaEntry = zip.CreateEntry("xl/vbaProject.bin", CompressionLevel.Optimal);
                using (var s = vbaEntry.Open()) s.Write(vbaProject, 0, vbaProject.Length);
            }
        }
    }
}
