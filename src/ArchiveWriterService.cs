using Disgaea_DS_Manager.Models;
namespace Disgaea_DS_Manager.Services;

public sealed class ArchiveWriterService
{
    public async Task<byte[]> SerializeAsync(ArchiveDocument doc, IProgress<double>? progress = null, CancellationToken ct = default)
    {
        return doc.FileType switch
        {
            ArchiveType.DSARC => await SerializeDsarcAsync(doc, progress, ct).ConfigureAwait(false),
            ArchiveType.MSND => await SerializeMsndAsync(doc, progress, ct).ConfigureAwait(false),
            _ => throw new NotSupportedException($"Unsupported archive type: {doc.FileType}")
        };
    }
    public async Task<byte[]> SerializeEntryAsync(ArchiveEntry entry, CancellationToken ct = default)
    {
        return !entry.IsNested || entry.Children.Count == 0
            ? await entry.GetDataAsync(ct).ConfigureAwait(false)
            : entry.NestedType switch
            {
                ArchiveType.MSND => await BuildMsndFromEntryAsync(entry, ct).ConfigureAwait(false),
                ArchiveType.DSARC => await BuildDsarcFromEntryAsync(entry, ct).ConfigureAwait(false),
                _ => await entry.GetDataAsync(ct).ConfigureAwait(false)
            };
    }
    private async Task<byte[]> SerializeDsarcAsync(ArchiveDocument doc, IProgress<double>? progress, CancellationToken ct)
    {
        List<ArchiveEntry> entries = doc.RootEntry.Children;
        int count = entries.Count;
        int headerSize = Formats.DsarcHeader + (count * (Formats.NameSize + Formats.DsarcEntryInfo));
        List<(string Name, byte[] Data)> dataList = new(count);
        int totalDataSize = 0;
        for (int i = 0; i < count; i++)
        {
            ct.ThrowIfCancellationRequested();
            ArchiveEntry entry = entries[i];
            byte[] data = await SerializeEntryAsync(entry, ct).ConfigureAwait(false);
            dataList.Add((entry.Name, data));
            totalDataSize += data.Length;
            progress?.Report((double)(i + 1) / (count * 2));
        }
        using MemoryStream ms = new(headerSize + totalDataSize);
        using BinaryWriter bw = new(ms);
        bw.Write(Formats.MagicDsarc);
        bw.Write(count);
        bw.Write(Formats.DsarcVersion);
        int offset = headerSize;
        foreach ((string name, byte[] data) in dataList)
        {
            bw.Write(Formats.PadName(name));
            bw.Write(data.Length);
            bw.Write(offset);
            offset += data.Length;
        }
        for (int i = 0; i < dataList.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            bw.Write(dataList[i].Data);
            progress?.Report(0.5 + ((double)(i + 1) / (count * 2)));
            await Task.Yield();
        }
        return ms.ToArray();
    }
    private async Task<byte[]> SerializeMsndAsync(ArchiveDocument doc, IProgress<double>? progress, CancellationToken ct)
    {
        List<ArchiveEntry> children = doc.RootEntry.Children;
        if (children.Count != 3)
        {
            throw new InvalidOperationException("MSND requires exactly 3 chunks (sseq, sbnk, swar).");
        }
        Dictionary<string, byte[]> chunks = new(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < 3; i++)
        {
            ct.ThrowIfCancellationRequested();
            string ext = Path.GetExtension(children[i].Name).ToLowerInvariant();
            chunks[ext] = await children[i].GetDataAsync(ct).ConfigureAwait(false);
            progress?.Report((double)(i + 1) / 3);
        }
        return Formats.BuildMsnd(chunks);
    }
    private async Task<byte[]> BuildMsndFromEntryAsync(ArchiveEntry entry, CancellationToken ct)
    {
        if (entry.Children.Count != 3)
        {
            throw new InvalidOperationException("MSND entry requires exactly 3 children.");
        }
        Dictionary<string, byte[]> chunks = new(StringComparer.OrdinalIgnoreCase);
        foreach (ArchiveEntry child in entry.Children)
        {
            ct.ThrowIfCancellationRequested();
            string ext = Path.GetExtension(child.Name).ToLowerInvariant();
            chunks[ext] = await child.GetDataAsync(ct).ConfigureAwait(false);
        }
        byte[]? unknownBytes = await TryExtractUnknownBytesAsync(entry, ct).ConfigureAwait(false);
        return Formats.BuildMsnd(chunks, unknownBytes);
    }
    private static async Task<byte[]?> TryExtractUnknownBytesAsync(ArchiveEntry entry, CancellationToken ct)
    {
        if (entry.DataSource is null)
        {
            return null;
        }
        try
        {
            byte[] originalData = await entry.DataSource.GetDataAsync(ct).ConfigureAwait(false);
            if (originalData.Length >= Formats.MsndHeader && originalData.AsSpan(0, 4).SequenceEqual(Formats.MagicMsnd))
            {
                return Formats.ExtractMsndUnknownBytes(originalData);
            }
        }
        catch { }
        return null;
    }
    private async Task<byte[]> BuildDsarcFromEntryAsync(ArchiveEntry entry, CancellationToken ct)
    {
        ArchiveDocument tempDoc = new() { FileType = ArchiveType.DSARC, RootEntry = entry };
        return await SerializeDsarcAsync(tempDoc, null, ct).ConfigureAwait(false);
    }
    public byte[] ReplaceChunk(byte[] msndBuf, string ext, byte[] newData)
    {
        ArgumentNullException.ThrowIfNull(msndBuf);
        ArgumentNullException.ThrowIfNull(newData);
        if (msndBuf.Length < Formats.MsndHeader || !msndBuf.AsSpan(0, 4).SequenceEqual(Formats.MagicMsnd))
        {
            throw new InvalidDataException("Invalid MSND buffer.");
        }
        MsndOffsets offsets = Formats.ParseMsndOffsets(msndBuf);
        byte[] unknownBytes = Formats.ExtractMsndUnknownBytes(msndBuf);
        Dictionary<string, byte[]> chunks = new(StringComparer.OrdinalIgnoreCase)
        {
            [".sseq"] = msndBuf[offsets.SseqOffset..(offsets.SseqOffset + offsets.SseqSize)],
            [".sbnk"] = msndBuf[offsets.SbnkOffset..(offsets.SbnkOffset + offsets.SbnkSize)],
            [".swar"] = msndBuf[offsets.SwarOffset..(offsets.SwarOffset + offsets.SwarSize)]
        };
        string normalizedExt = ext.ToLowerInvariant();
        if (!normalizedExt.StartsWith('.'))
        {
            normalizedExt = "." + normalizedExt;
        }
        if (!chunks.ContainsKey(normalizedExt))
        {
            throw new ArgumentException($"Invalid chunk extension: {ext}", nameof(ext));
        }
        chunks[normalizedExt] = newData;
        return Formats.BuildMsnd(chunks, unknownBytes);
    }
}