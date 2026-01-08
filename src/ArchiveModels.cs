using System.Buffers;
using System.Text;
namespace Disgaea_DS_Manager.Models;

public enum ArchiveType { DSARC, MSND }
public sealed record ArchiveDocument
{
    public string? FilePath { get; init; }
    public ArchiveType FileType { get; init; }
    public ArchiveEntry RootEntry { get; init; } = new() { Name = "[Root]" };
    public bool IsModified { get; init; }
    public string? OriginalFilePath { get; init; }
    public bool HasContent { get; init; }
    public List<ArchiveEntry> GetAllEntries()
    {
        List<ArchiveEntry> result = new(16);
        CollectEntries(RootEntry, result);
        return result;
    }
    private static void CollectEntries(ArchiveEntry entry, List<ArchiveEntry> list)
    {
        foreach (ArchiveEntry child in entry.Children)
        {
            list.Add(child);
            if (child.Children.Count > 0)
            {
                CollectEntries(child, list);
            }
        }
    }
    public ArchiveDocument DeepCopy()
    {
        return this with { RootEntry = RootEntry.DeepCopy() };
    }
}
public sealed class ArchiveEntry
{
    private static long _nextId;
    public long Id { get; private set; } = Interlocked.Increment(ref _nextId);
    public string Name { get; set; } = string.Empty;
    public int Size { get; set; }
    public int Offset { get; set; }
    public ArchiveType? NestedType { get; set; }
    public IDataSource? DataSource { get; set; }
    public List<ArchiveEntry> Children { get; set; } = [];
    public bool IsModified { get; set; }
    public int ImportOrder { get; set; }
    public bool IsNested => NestedType.HasValue;
    public string DisplayName => IsModified ? $"{Name} *" : Name;
    public async Task<byte[]> GetDataAsync(CancellationToken ct = default)
    {
        return DataSource is null
            ? throw new InvalidOperationException($"Entry '{Name}' has no data source.")
            : await DataSource.GetDataAsync(ct).ConfigureAwait(false);
    }
    public ArchiveEntry DeepCopy()
    {
        return new()
        {
            Id = Id,
            Name = Name,
            Size = Size,
            Offset = Offset,
            NestedType = NestedType,
            DataSource = DataSource,
            IsModified = IsModified,
            ImportOrder = ImportOrder,
            Children = [.. Children.Select(c => c.DeepCopy())]
        };
    }
    public ArchiveEntry DeepCopyWithNewId()
    {
        return new()
        {
            Name = Name,
            Size = Size,
            Offset = Offset,
            NestedType = NestedType,
            DataSource = DataSource,
            IsModified = IsModified,
            ImportOrder = ImportOrder,
            Children = [.. Children.Select(c => c.DeepCopyWithNewId())]
        };
    }
}
public interface IDataSource
{
    Task<byte[]> GetDataAsync(CancellationToken ct = default);
}
public sealed class FileRangeSource(string path, long offset, int size) : IDataSource
{
    public async Task<byte[]> GetDataAsync(CancellationToken ct = default)
    {
        byte[] buffer = new byte[size];
        await using FileStream fs = new(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
        _ = fs.Seek(offset, SeekOrigin.Begin);
        await fs.ReadExactlyAsync(buffer, ct).ConfigureAwait(false);
        return buffer;
    }
}
public sealed class BufferSource(byte[] data) : IDataSource
{
    public Task<byte[]> GetDataAsync(CancellationToken ct = default)
    {
        return Task.FromResult(data);
    }
}
public sealed class ImportResult
{
    public ArchiveType FileType { get; set; }
    public List<ArchiveEntry> Entries { get; } = [];
    public required string SourceFolder { get; set; }
}
public static class Formats
{
    public const int NameSize = 40;
    public const int DsarcHeader = 16;
    public const int DsarcEntryInfo = 8;
    public const int DsarcVersion = 1;
    public const int MsndHeader = 48;
    public const int MsndUnknownOffset = 0x2C;
    public static readonly byte[] MagicDsarc = [0x44, 0x53, 0x41, 0x52, 0x43, 0x20, 0x46, 0x4C];
    public static readonly byte[] MagicMsnd = [0x44, 0x53, 0x45, 0x51];
    public static readonly string[] MsndOrder = [".sseq", ".sbnk", ".swar"];
    private static readonly ArrayPool<byte> Pool = ArrayPool<byte>.Shared;
    public static byte[] PadName(string name)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(name ?? string.Empty);
        if (bytes.Length > NameSize)
        {
            bytes = bytes[..NameSize];
        }
        byte[] result = new byte[NameSize];
        bytes.CopyTo(result, 0);
        return result;
    }
    public static string GuessExtension(ReadOnlySpan<byte> data, string fallback)
    {
        return data.Length < 4
            ? fallback
            : (data[0], data[1], data[2], data[3]) switch
            {
                (0x53, 0x57, 0x41, 0x56) => ".swav",
                (0x53, 0x54, 0x52, 0x4D) => ".strm",
                _ => fallback
            };
    }
    public static ArchiveType DetectType(string path)
    {
        byte[] magic = Pool.Rent(8);
        try
        {
            using FileStream fs = new(path, FileMode.Open, FileAccess.Read, FileShare.Read, 8, FileOptions.SequentialScan);
            int read = fs.Read(magic, 0, 8);
            return read < 4
                ? throw new InvalidDataException("Cannot read file header.")
                : magic.AsSpan(0, 4).SequenceEqual(MagicMsnd)
                ? ArchiveType.MSND
                : read >= 8 && magic.AsSpan(0, 8).SequenceEqual(MagicDsarc)
                ? ArchiveType.DSARC
                : throw new InvalidDataException("Unknown archive format.");
        }
        finally { Pool.Return(magic); }
    }
    public static ArchiveType? DetectTypeFromBuffer(ReadOnlySpan<byte> data)
    {
        return data.Length >= 8 && data[..8].SequenceEqual(MagicDsarc)
            ? ArchiveType.DSARC
            : data.Length >= 4 && data[..4].SequenceEqual(MagicMsnd) ? ArchiveType.MSND : null;
    }
    public static MsndOffsets ParseMsndOffsets(ReadOnlySpan<byte> data)
    {
        if (data.Length < MsndHeader)
        {
            throw new InvalidDataException("Buffer too small for MSND header.");
        }
        int sseqOff = BitConverter.ToInt32(data[16..]) - 16;
        int sbnkOff = BitConverter.ToInt32(data[20..]);
        int swarOff = BitConverter.ToInt32(data[24..]);
        int sseqSz = BitConverter.ToInt32(data[32..]);
        int sbnkSz = BitConverter.ToInt32(data[36..]);
        int swarSz = BitConverter.ToInt32(data[40..]);
        return new MsndOffsets(sseqOff, sbnkOff, swarOff, sseqSz, sbnkSz, swarSz);
    }
    public static bool AreMsndBoundsValid(int length, MsndOffsets offsets)
    {
        return offsets.SseqOffset >= 0 && offsets.SseqSize >= 0 && offsets.SseqOffset + offsets.SseqSize <= length &&
               offsets.SbnkOffset >= 0 && offsets.SbnkSize >= 0 && offsets.SbnkOffset + offsets.SbnkSize <= length &&
               offsets.SwarOffset >= 0 && offsets.SwarSize >= 0 && offsets.SwarOffset + offsets.SwarSize <= length;
    }
    public static void PopulateMsndChildren(byte[] data, ArchiveEntry parent, string baseName)
    {
        MsndOffsets offsets = ParseMsndOffsets(data);
        if (!AreMsndBoundsValid(data.Length, offsets))
        {
            return;
        }
        parent.Children.Clear();
        parent.Children.Add(CreateMsndChildEntry($"{baseName}.sseq", data, offsets.SseqOffset, offsets.SseqSize));
        parent.Children.Add(CreateMsndChildEntry($"{baseName}.sbnk", data, offsets.SbnkOffset, offsets.SbnkSize));
        parent.Children.Add(CreateMsndChildEntry($"{baseName}.swar", data, offsets.SwarOffset, offsets.SwarSize));
    }
    public static ArchiveEntry CreateMsndChildEntry(string name, byte[] buf, int offset, int size)
    {
        return new()
        {
            Name = name,
            Size = size,
            Offset = offset,
            DataSource = new BufferSource(buf[offset..(offset + size)])
        };
    }
    public static byte[] BuildMsnd(Dictionary<string, byte[]> chunks, byte[]? unknownBytes = null)
    {
        foreach (string ext in MsndOrder)
        {
            if (!chunks.ContainsKey(ext))
            {
                throw new InvalidOperationException($"Missing MSND chunk: {ext}");
            }
        }
        byte[] sseq = chunks[".sseq"];
        byte[] sbnk = chunks[".sbnk"];
        byte[] swar = chunks[".swar"];
        int sseqOff = MsndHeader;
        int sbnkOff = sseqOff + sseq.Length;
        int swarOff = sbnkOff + sbnk.Length;
        using MemoryStream ms = new(MsndHeader + sseq.Length + sbnk.Length + swar.Length);
        using BinaryWriter bw = new(ms);
        bw.Write(MagicMsnd);
        bw.Write(new byte[12]);
        bw.Write(sseqOff + 16);
        bw.Write(sbnkOff);
        bw.Write(swarOff);
        bw.Write(0);
        bw.Write(sseq.Length);
        bw.Write(sbnk.Length);
        bw.Write(swar.Length);
        bw.Write(unknownBytes is { Length: 4 } ? unknownBytes : new byte[4]);
        bw.Write(sseq);
        bw.Write(sbnk);
        bw.Write(swar);
        return ms.ToArray();
    }
    public static byte[] BuildEmptyMsnd()
    {
        using MemoryStream ms = new(MsndHeader);
        using BinaryWriter bw = new(ms);
        bw.Write(MagicMsnd);
        bw.Write(new byte[12]);
        int sseqOff = MsndHeader;
        bw.Write(sseqOff + 16);
        bw.Write(sseqOff);
        bw.Write(sseqOff);
        bw.Write(0);
        bw.Write(0);
        bw.Write(0);
        bw.Write(0);
        bw.Write(new byte[4]);
        return ms.ToArray();
    }
    public static byte[] ExtractMsndUnknownBytes(ReadOnlySpan<byte> msndData)
    {
        return msndData.Length < MsndHeader ? new byte[4] : msndData[MsndUnknownOffset..(MsndUnknownOffset + 4)].ToArray();
    }
}
public readonly record struct MsndOffsets(int SseqOffset, int SbnkOffset, int SwarOffset, int SseqSize, int SbnkSize, int SwarSize);