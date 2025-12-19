using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
namespace Disgaea_DS_Manager;

public enum ArchiveType { DSARC, MSND }
public class ImportResult
{
    public ArchiveType FileType { get; set; }
    public Collection<Entry> Entries { get; } = [];
    public required string SourceFolder { get; set; }
}
public static class Helpers
{
    public const int NAMESZ = 40;
    public static byte[] PadName(string n)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(n ?? string.Empty);
        if (bytes.Length > NAMESZ)
        {
            bytes = bytes.Take(NAMESZ).ToArray();
        }

        return bytes.Concat(new byte[NAMESZ - bytes.Length]).ToArray();
    }
    public static string GuessExtByMagic(ReadOnlySpan<byte> data, string defaultExt)
    {
        return data.Length < 4
            ? defaultExt
            : (data[0], data[1], data[2], data[3]) switch
            {
                (0x53, 0x57, 0x41, 0x56) => ".swav",
                (0x53, 0x54, 0x52, 0x4D) => ".strm",
                _ => defaultExt
            };
    }
}
public class Entry
{
    public FileInfo Path { get; set; }
    public int Size { get; set; }
    public int Offset { get; set; }
    public bool IsMsnd { get; set; }
    public Collection<Entry> Children { get; } = [];
    public Entry(FileInfo path, int size = 0, int offset = 0, bool isMsnd = false)
    {
        Path = path;
        Size = size;
        Offset = offset;
        IsMsnd = isMsnd;
    }
}
public class TupleComparer : IEqualityComparer<Tuple<string, string>>
{
    public bool Equals(Tuple<string, string> x, Tuple<string, string> y)
    {
        return string.Equals(x?.Item1, y?.Item1, StringComparison.Ordinal) &&
        string.Equals(x?.Item2, y?.Item2, StringComparison.Ordinal);
    }

    public int GetHashCode(Tuple<string, string> obj)
    {
        return HashCode.Combine(
                StringComparer.Ordinal.GetHashCode(obj?.Item1 ?? string.Empty),
                StringComparer.Ordinal.GetHashCode(obj?.Item2 ?? string.Empty));
    }
}
public static class Detector
{
    private static readonly byte[] MAGIC_DSARC = [0x44, 0x53, 0x41, 0x52, 0x43, 0x20, 0x46, 0x4C];
    private static readonly byte[] MAGIC_MSND = [0x44, 0x53, 0x45, 0x51];
    public static ArchiveType FromFile(string path)
    {
        using FileStream fs = new(path, FileMode.Open, FileAccess.Read);
        Span<byte> magic8 = stackalloc byte[8];
        int read = fs.Read(magic8);
        return read < 4 ? throw new InvalidDataException("Unable to read file header.")
            : magic8[..4].SequenceEqual(MAGIC_MSND) ? ArchiveType.MSND
            : read >= 8 && magic8.SequenceEqual(MAGIC_DSARC) ? ArchiveType.DSARC
            : throw new InvalidDataException("Unknown archive format (magic mismatch).");
    }
}
public static class Msnd
{
    public const int HDR_MSND = 48;
    public static readonly string[] MSND_ORDER = [".sseq", ".sbnk", ".swar"];
    public static readonly byte[] MAGIC_MSND = [0x44, 0x53, 0x45, 0x51];
    public static Collection<Entry> Parse(ReadOnlySpan<byte> buf, string baseName)
    {
        if (!buf[..4].SequenceEqual(MAGIC_MSND))
        {
            throw new InvalidDataException("Not an MSND archive.");
        }

        if (buf.Length < HDR_MSND)
        {
            throw new InvalidDataException("MSND too small.");
        }

        int sseqOff = BitConverter.ToInt32(buf[16..]) - 16;
        int sbnkOff = BitConverter.ToInt32(buf[20..]);
        int swarOff = BitConverter.ToInt32(buf[24..]);
        int sseqSz = BitConverter.ToInt32(buf[32..]);
        int sbnkSz = BitConverter.ToInt32(buf[36..]);
        int swarSz = BitConverter.ToInt32(buf[40..]);
        (string name, int offset, int size)[] chunks = new (string name, int offset, int size)[]
        {
            ("SSEQ", sseqOff, sseqSz),
            ("SBNK", sbnkOff, sbnkSz),
            ("SWAR", swarOff, swarSz)
        };
        foreach ((string name, int offset, int size) in chunks)
        {
            if (offset < 0 || size < 0 || offset + size > buf.Length)
            {
                throw new InvalidDataException($"{name} file exceeds bounds.");
            }
        }
        return
        [
            new(new FileInfo($"{baseName}.sseq"), sseqSz, sseqOff),
            new(new FileInfo($"{baseName}.sbnk"), sbnkSz, sbnkOff),
            new(new FileInfo($"{baseName}.swar"), swarSz, swarOff)
        ];
    }
    public static byte[] Build(Dictionary<string, byte[]> chunks, byte[]? txtBytes = null)
    {
        foreach (string ext in MSND_ORDER)
        {
            if (!chunks.ContainsKey(ext))
            {
                throw new InvalidOperationException($"Missing {ext}");
            }
        }

        byte[] sseq = chunks[".sseq"], sbnk = chunks[".sbnk"], swar = chunks[".swar"];
        int sseqOff = HDR_MSND;
        int sbnkOff = sseqOff + sseq.Length;
        int swarOff = sbnkOff + sbnk.Length;
        Span<byte> hdr = stackalloc byte[HDR_MSND];
        MAGIC_MSND.CopyTo(hdr);
        _ = BitConverter.TryWriteBytes(hdr[16..], sseqOff + 16);
        _ = BitConverter.TryWriteBytes(hdr[20..], sbnkOff);
        _ = BitConverter.TryWriteBytes(hdr[24..], swarOff);
        _ = BitConverter.TryWriteBytes(hdr[32..], sseq.Length);
        _ = BitConverter.TryWriteBytes(hdr[36..], sbnk.Length);
        _ = BitConverter.TryWriteBytes(hdr[40..], swar.Length);
        if (txtBytes is { Length: 4 })
        {
            txtBytes.CopyTo(hdr[44..]);
        }

        using MemoryStream ms = new();
        using BinaryWriter bw = new(ms);
        bw.Write(hdr);
        bw.Write(sseq);
        bw.Write(sbnk);
        bw.Write(swar);
        return ms.ToArray();
    }
    public static byte[] ReplaceChunk(byte[] msndBuf, string ext, byte[] newData)
    {
        ArgumentNullException.ThrowIfNull(msndBuf);
        ArgumentNullException.ThrowIfNull(ext);
        ArgumentNullException.ThrowIfNull(newData);
        Collection<Entry> entries = Parse(msndBuf, "temp");
        Dictionary<string, Entry> map = new(StringComparer.OrdinalIgnoreCase)
        {
            [".sseq"] = entries[0],
            [".sbnk"] = entries[1],
            [".swar"] = entries[2]
        };
        if (!map.TryGetValue(ext.ToLowerInvariant(), out Entry? entry))
        {
            throw new ArgumentException("Unsupported file", nameof(ext));
        }

        Dictionary<string, byte[]> chunks = MSND_ORDER.ToDictionary(
            e => e,
            e => e == ext ? newData : msndBuf.Skip(map[e].Offset).Take(map[e].Size).ToArray());
        byte[]? txt = msndBuf.Length >= HDR_MSND ? msndBuf[44..48] : null;
        return Build(chunks, txt);
    }
}
public static class Dsarc
{
    public const int HDR_DSARC = 16;
    public const int ENTRY_INFO_SZ = 8;
    public const int VERSION = 1;
    public static readonly byte[] MAGIC_DSARC = [0x44, 0x53, 0x41, 0x52, 0x43, 0x20, 0x46, 0x4C];
    private static readonly char[] Separator = ['='];
    public static Collection<Entry> Parse(string path)
    {
        using FileStream fs = new(path, FileMode.Open, FileAccess.Read);
        using BinaryReader br = new(fs, Encoding.UTF8, true);
        if (!br.ReadBytes(8).SequenceEqual(MAGIC_DSARC))
        {
            throw new InvalidDataException("Not a DSARC.");
        }

        int count = br.ReadInt32();
        int version = br.ReadInt32();
        if (version != VERSION)
        {
            throw new NotSupportedException($"Unsupported DSARC version {version}");
        }

        _ = fs.Seek(HDR_DSARC, SeekOrigin.Begin);
        long archiveSize = new FileInfo(path).Length;
        List<(string name, int size, int offset)> entriesMeta = [];
        for (int i = 0; i < count; i++)
        {
            byte[] raw = br.ReadBytes(Helpers.NAMESZ);
            if (raw.Length < Helpers.NAMESZ)
            {
                throw new InvalidDataException("Corrupted entry name");
            }

            string name = Encoding.UTF8.GetString(raw).Split('\0')[0].Trim();
            if (string.IsNullOrEmpty(name))
            {
                name = $"file_{i}";
            }

            byte[] info = br.ReadBytes(ENTRY_INFO_SZ);
            if (info.Length < ENTRY_INFO_SZ)
            {
                throw new InvalidDataException("Corrupted entry info");
            }

            int size = BitConverter.ToInt32(info, 0);
            int offset = BitConverter.ToInt32(info, 4);
            if (offset + (long)size > archiveSize)
            {
                throw new InvalidDataException($"{name} exceeds bounds");
            }

            entriesMeta.Add((name, size, offset));
        }
        Collection<Entry> outList = [];
        foreach ((string name, int size, int offset) in entriesMeta)
        {
            Entry entry = new(new FileInfo(name), size, offset);
            _ = fs.Seek(offset, SeekOrigin.Begin);
            byte[] magicAttempt = br.ReadBytes(4);
            if (magicAttempt.SequenceEqual(Msnd.MAGIC_MSND))
            {
                _ = fs.Seek(offset, SeekOrigin.Begin);
                byte[] msndBuf = br.ReadBytes(size);
                entry.IsMsnd = true;
                foreach (Entry child in Msnd.Parse(msndBuf, System.IO.Path.GetFileNameWithoutExtension(name)))
                {
                    entry.Children.Add(child);
                }
            }
            outList.Add(entry);
        }
        return outList;
    }

    public static Collection<Entry> ParseFromBuffer(byte[] buf)
    {
        ArgumentNullException.ThrowIfNull(buf);
        using MemoryStream ms = new(buf);
        using BinaryReader br = new(ms, Encoding.UTF8, true);
        if (!br.ReadBytes(8).SequenceEqual(MAGIC_DSARC))
        {
            throw new InvalidDataException("Not a DSARC in buffer.");
        }

        int count = br.ReadInt32();
        int version = br.ReadInt32();
        if (version != VERSION)
        {
            throw new NotSupportedException($"Unsupported DSARC version {version}");
        }

        _ = ms.Seek(HDR_DSARC, SeekOrigin.Begin);
        List<(string name, int size, int offset)> entriesMeta = [];
        for (int i = 0; i < count; i++)
        {
            byte[] raw = br.ReadBytes(Helpers.NAMESZ);
            if (raw.Length < Helpers.NAMESZ)
            {
                throw new InvalidDataException("Corrupted entry name");
            }

            string name = Encoding.UTF8.GetString(raw).Split('\0')[0].Trim();
            if (string.IsNullOrEmpty(name))
            {
                name = $"file_{i}";
            }

            byte[] info = br.ReadBytes(ENTRY_INFO_SZ);
            if (info.Length < ENTRY_INFO_SZ)
            {
                throw new InvalidDataException("Corrupted entry info");
            }

            int size = BitConverter.ToInt32(info, 0);
            int offset = BitConverter.ToInt32(info, 4);
            if (offset + (long)size > buf.Length)
            {
                throw new InvalidDataException($"{name} exceeds bounds in buffer");
            }

            entriesMeta.Add((name, size, offset));
        }
        Collection<Entry> outList = [];
        foreach ((string name, int size, int offset) in entriesMeta)
        {
            Entry entry = new(new FileInfo(name), size, offset);
            _ = ms.Seek(offset, SeekOrigin.Begin);
            byte[] magicAttempt = br.ReadBytes(4);
            if (magicAttempt.SequenceEqual(Msnd.MAGIC_MSND))
            {
                _ = ms.Seek(offset, SeekOrigin.Begin);
                byte[] msndBuf = br.ReadBytes(size);
                entry.IsMsnd = true;
                foreach (Entry child in Msnd.Parse(msndBuf, System.IO.Path.GetFileNameWithoutExtension(name)))
                {
                    entry.Children.Add(child);
                }
            }
            outList.Add(entry);
        }
        return outList;
    }
    public static byte[] BuildFromPairs(Collection<Tuple<string, byte[]>> pairs)
    {
        ArgumentNullException.ThrowIfNull(pairs);
        using MemoryStream ms = new();
        using BinaryWriter bw = new(ms, Encoding.UTF8, true);
        bw.Write(MAGIC_DSARC);
        bw.Write(pairs.Count);
        bw.Write(VERSION);
        int offset = HDR_DSARC + (pairs.Count * (Helpers.NAMESZ + ENTRY_INFO_SZ));
        foreach ((string name, byte[] data) in pairs)
        {
            bw.Write(Helpers.PadName(name));
            bw.Write(data.Length);
            bw.Write(offset);
            offset += data.Length;
        }
        foreach ((string _, byte[] data) in pairs)
        {
            bw.Write(data);
        }

        return ms.ToArray();
    }
}