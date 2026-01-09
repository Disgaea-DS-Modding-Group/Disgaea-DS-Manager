using Disgaea_DS_Manager.Models;
using System.Text;
namespace Disgaea_DS_Manager.Services;

public sealed class ArchiveReaderService
{
    public async Task<ArchiveDocument> ParseAsync(string path, CancellationToken ct = default)
    {
        ArchiveType type = Formats.DetectType(path);
        byte[] data = await File.ReadAllBytesAsync(path, ct).ConfigureAwait(false);
        return type switch
        {
            ArchiveType.MSND => ParseMsnd(data, path),
            ArchiveType.DSARC => ParseDsarc(data, path),
            _ => throw new NotSupportedException("Unknown archive type")
        };
    }
    public ArchiveDocument ParseFromBuffer(byte[] data, string virtualPath)
    {
        ArchiveType type = Formats.DetectTypeFromBuffer(data)
            ?? throw new InvalidDataException("Unknown buffer format.");
        return type switch
        {
            ArchiveType.MSND => ParseMsnd(data, virtualPath),
            ArchiveType.DSARC => ParseDsarc(data, virtualPath),
            _ => throw new NotSupportedException()
        };
    }
    private static ArchiveDocument ParseDsarc(byte[] buf, string path)
    {
        if (buf.Length < Formats.DsarcHeader || !buf.AsSpan(0, 8).SequenceEqual(Formats.MagicDsarc))
        {
            throw new InvalidDataException("Invalid DSARC header.");
        }
        int count = BitConverter.ToInt32(buf, 8);
        int version = BitConverter.ToInt32(buf, 12);
        if (version != Formats.DsarcVersion)
        {
            throw new NotSupportedException($"Unsupported DSARC version: {version}");
        }
        ArchiveEntry root = new() { Name = Path.GetFileName(path) };
        int pos = Formats.DsarcHeader;
        for (int i = 0; i < count; i++)
        {
            ReadOnlySpan<byte> nameBytes = buf.AsSpan(pos, Formats.NameSize);
            string name = Encoding.UTF8.GetString(nameBytes).Split('\0')[0].Trim();
            if (string.IsNullOrEmpty(name))
            {
                name = $"file_{i}";
            }
            pos += Formats.NameSize;
            int size = BitConverter.ToInt32(buf, pos);
            int offset = BitConverter.ToInt32(buf, pos + 4);
            pos += Formats.DsarcEntryInfo;
            if (offset < 0 || size < 0 || offset + size > buf.Length)
            {
                throw new InvalidDataException($"Entry '{name}' exceeds archive bounds.");
            }
            ArchiveEntry entry = new()
            {
                Name = name,
                Size = size,
                Offset = offset,
                ImportOrder = i,
                DataSource = new BufferSource(buf[offset..(offset + size)])
            };
            ReadOnlySpan<byte> entryData = buf.AsSpan(offset, Math.Min(size, 8));
            if (size >= 4 && entryData[..4].SequenceEqual(Formats.MagicMsnd))
            {
                entry.NestedType = ArchiveType.MSND;
                ParseMsndChildren(buf.AsSpan(offset, size), entry, name);
            }
            else if (size >= 8 && entryData[..8].SequenceEqual(Formats.MagicDsarc))
            {
                entry.NestedType = ArchiveType.DSARC;
                ParseDsarcChildren(buf.AsSpan(offset, size), entry, name);
            }
            root.Children.Add(entry);
        }
        return new ArchiveDocument
        {
            FilePath = path,
            FileType = ArchiveType.DSARC,
            OriginalFilePath = path,
            RootEntry = root
        };
    }
    private static ArchiveDocument ParseMsnd(byte[] buf, string path)
    {
        if (buf.Length < Formats.MsndHeader || !buf.AsSpan(0, 4).SequenceEqual(Formats.MagicMsnd))
        {
            throw new InvalidDataException("Invalid MSND header.");
        }
        string baseName = Path.GetFileNameWithoutExtension(path);
        ArchiveEntry root = new() { Name = Path.GetFileName(path), NestedType = ArchiveType.MSND };
        MsndOffsets offsets = Formats.ParseMsndOffsets(buf);
        ValidateMsndBounds(buf.Length, offsets);
        root.Children.Add(Formats.CreateMsndChildEntry($"{baseName}.sseq", buf, offsets.SseqOffset, offsets.SseqSize));
        root.Children.Add(Formats.CreateMsndChildEntry($"{baseName}.sbnk", buf, offsets.SbnkOffset, offsets.SbnkSize));
        root.Children.Add(Formats.CreateMsndChildEntry($"{baseName}.swar", buf, offsets.SwarOffset, offsets.SwarSize));
        return new ArchiveDocument
        {
            FilePath = path,
            FileType = ArchiveType.MSND,
            OriginalFilePath = path,
            RootEntry = root
        };
    }
    private static void ParseMsndChildren(ReadOnlySpan<byte> msndData, ArchiveEntry parent, string parentName)
    {
        if (msndData.Length < Formats.MsndHeader)
        {
            return;
        }
        string baseName = Path.GetFileNameWithoutExtension(parentName);
        byte[] fullData = msndData.ToArray();
        MsndOffsets offsets = Formats.ParseMsndOffsets(fullData);
        if (!Formats.AreMsndBoundsValid(msndData.Length, offsets))
        {
            return;
        }
        parent.Children.Add(Formats.CreateMsndChildEntry($"{baseName}.sseq", fullData, offsets.SseqOffset, offsets.SseqSize));
        parent.Children.Add(Formats.CreateMsndChildEntry($"{baseName}.sbnk", fullData, offsets.SbnkOffset, offsets.SbnkSize));
        parent.Children.Add(Formats.CreateMsndChildEntry($"{baseName}.swar", fullData, offsets.SwarOffset, offsets.SwarSize));
    }
    private static void ParseDsarcChildren(ReadOnlySpan<byte> dsarcData, ArchiveEntry parent, string parentName)
    {
        try
        {
            if (dsarcData.Length < Formats.DsarcHeader || !dsarcData[..8].SequenceEqual(Formats.MagicDsarc))
            {
                return;
            }
            int count = BitConverter.ToInt32(dsarcData[8..12]);
            int version = BitConverter.ToInt32(dsarcData[12..16]);
            if (version != Formats.DsarcVersion)
            {
                return;
            }
            parent.Children.Clear();
            byte[] fullData = dsarcData.ToArray();
            int pos = Formats.DsarcHeader;
            for (int i = 0; i < count; i++)
            {
                ReadOnlySpan<byte> nameBytes = dsarcData[pos..(pos + Formats.NameSize)];
                string name = Encoding.UTF8.GetString(nameBytes).Split('\0')[0].Trim();
                if (string.IsNullOrEmpty(name))
                {
                    name = $"file_{i}";
                }
                pos += Formats.NameSize;
                int size = BitConverter.ToInt32(dsarcData[pos..(pos + 4)]);
                int offset = BitConverter.ToInt32(dsarcData[(pos + 4)..(pos + 8)]);
                pos += Formats.DsarcEntryInfo;
                if (offset < 0 || size < 0 || offset + size > dsarcData.Length)
                {
                    continue;
                }
                ArchiveEntry child = new()
                {
                    Name = name,
                    Size = size,
                    Offset = offset,
                    ImportOrder = i,
                    DataSource = new BufferSource(fullData[offset..(offset + size)])
                };
                int checkSize = Math.Min(size, 8);
                ReadOnlySpan<byte> entryData = dsarcData[offset..(offset + checkSize)];
                if (size >= 4 && entryData[..Math.Min(4, entryData.Length)].SequenceEqual(Formats.MagicMsnd))
                {
                    child.NestedType = ArchiveType.MSND;
                    ParseMsndChildren(dsarcData[offset..(offset + size)], child, name);
                }
                else if (size >= 8 && entryData[..Math.Min(8, entryData.Length)].SequenceEqual(Formats.MagicDsarc))
                {
                    child.NestedType = ArchiveType.DSARC;
                    ParseDsarcChildren(dsarcData[offset..(offset + size)], child, name);
                }
                parent.Children.Add(child);
            }
        }
        catch { }
    }
    private static void ValidateMsndBounds(int length, MsndOffsets offsets)
    {
        if (!Formats.AreMsndBoundsValid(length, offsets))
        {
            throw new InvalidDataException("MSND chunk exceeds bounds.");
        }
    }
}