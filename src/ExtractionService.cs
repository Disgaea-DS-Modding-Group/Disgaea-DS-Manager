using Disgaea_DS_Manager.Models;
using System.Text;
namespace Disgaea_DS_Manager.Services;

public sealed class ExtractionService
{
    private readonly ArchiveWriterService _writer = new();
    public async Task ExtractAllAsync(
        ArchiveDocument doc,
        string destFolder,
        bool nested,
        IProgress<(int Current, int Total)>? progress = null,
        CancellationToken ct = default)
    {
        string baseOut = Path.Combine(destFolder, Path.GetFileNameWithoutExtension(doc.FilePath ?? "archive"));
        _ = Directory.CreateDirectory(baseOut);
        List<ArchiveEntry> entries = doc.RootEntry.Children;
        if (doc.FileType == ArchiveType.MSND)
        {
            await ExtractMsndDocumentAsync(doc, baseOut, entries, progress, ct).ConfigureAwait(false);
        }
        else
        {
            await ExtractDsarcDocumentAsync(baseOut, entries, nested, progress, ct).ConfigureAwait(false);
        }
    }
    private async Task ExtractMsndDocumentAsync(
        ArchiveDocument doc,
        string baseOut,
        List<ArchiveEntry> entries,
        IProgress<(int Current, int Total)>? progress,
        CancellationToken ct)
    {
        string baseName = Path.GetFileNameWithoutExtension(doc.FilePath ?? "archive");
        byte[] msndData = await _writer.SerializeAsync(doc, null, ct).ConfigureAwait(false);
        byte[] unknownBytes = Formats.ExtractMsndUnknownBytes(msndData);
        for (int i = 0; i < entries.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            ArchiveEntry entry = entries[i];
            byte[] data = await entry.GetDataAsync(ct).ConfigureAwait(false);
            await File.WriteAllBytesAsync(Path.Combine(baseOut, entry.Name), data, ct).ConfigureAwait(false);
            progress?.Report((i + 1, entries.Count));
        }
        await File.WriteAllBytesAsync(Path.Combine(baseOut, $"{baseName}.txt"), unknownBytes, ct).ConfigureAwait(false);
    }
    private async Task ExtractDsarcDocumentAsync(
        string baseOut,
        List<ArchiveEntry> entries,
        bool nested,
        IProgress<(int Current, int Total)>? progress,
        CancellationToken ct)
    {
        List<string> mapperLines = new(entries.Count);
        Dictionary<(string, string), int> counters = [];
        for (int i = 0; i < entries.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            ArchiveEntry entry = entries[i];
            string outputName = await ExtractEntryAsync(entry, baseOut, nested, counters, ct).ConfigureAwait(false);
            mapperLines.Add($"{entry.Name}={outputName}");
            progress?.Report((i + 1, entries.Count));
        }
        await File.WriteAllLinesAsync(Path.Combine(baseOut, "mapper.txt"), mapperLines, Encoding.UTF8, ct).ConfigureAwait(false);
    }
    public async Task ExtractSingleAsync(ArchiveEntry entry, string destFolder, CancellationToken ct = default)
    {
        _ = Directory.CreateDirectory(destFolder);
        byte[] data = await entry.GetDataAsync(ct).ConfigureAwait(false);
        string ext = Formats.GuessExtension(data, Path.GetExtension(entry.Name));
        string fileName = Path.GetFileNameWithoutExtension(entry.Name) + ext;
        await File.WriteAllBytesAsync(Path.Combine(destFolder, fileName), data, ct).ConfigureAwait(false);
    }
    public async Task ExtractNestedAsync(
        ArchiveEntry entry,
        string destFolder,
        IProgress<(int Current, int Total)>? progress = null,
        CancellationToken ct = default)
    {
        string baseOut = Path.Combine(destFolder, Path.GetFileNameWithoutExtension(entry.Name));
        _ = Directory.CreateDirectory(baseOut);
        byte[] data = await _writer.SerializeEntryAsync(entry, ct).ConfigureAwait(false);
        List<ArchiveEntry> children = entry.Children.Count > 0 ? entry.Children : ParseChildren(data, entry.Name);
        List<string> mapperLines = new(children.Count);
        Dictionary<(string, string), int> counters = [];
        for (int i = 0; i < children.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            ArchiveEntry child = children[i];
            byte[] childData = data[child.Offset..(child.Offset + child.Size)];
            string outputName = await ExtractChildDataAsync(child.Name, childData, baseOut, counters, ct).ConfigureAwait(false);
            mapperLines.Add($"{child.Name}={outputName}");
            progress?.Report((i + 1, children.Count));
        }
        if (entry.NestedType == ArchiveType.MSND)
        {
            string baseName = Path.GetFileNameWithoutExtension(entry.Name);
            byte[] unknownBytes = Formats.ExtractMsndUnknownBytes(data);
            await File.WriteAllBytesAsync(Path.Combine(baseOut, $"{baseName}.txt"), unknownBytes, ct).ConfigureAwait(false);
        }
        else
        {
            await File.WriteAllLinesAsync(Path.Combine(baseOut, "mapper.txt"), mapperLines, Encoding.UTF8, ct).ConfigureAwait(false);
        }
    }
    private async Task<string> ExtractEntryAsync(
        ArchiveEntry entry,
        string outDir,
        bool nested,
        Dictionary<(string, string), int> counters,
        CancellationToken ct)
    {
        byte[] data = await _writer.SerializeEntryAsync(entry, ct).ConfigureAwait(false);
        return await ExtractChildDataAsync(entry.Name, data, outDir, counters, ct, nested).ConfigureAwait(false);
    }
    private async Task<string> ExtractChildDataAsync(
        string name,
        byte[] data,
        string outDir,
        Dictionary<(string, string), int> counters,
        CancellationToken ct,
        bool nested = true)
    {
        ArchiveType? type = Formats.DetectTypeFromBuffer(data);
        if (nested && type.HasValue)
        {
            string folderName = GetUniqueName(Path.GetFileNameWithoutExtension(name), "", outDir, counters);
            string childDir = Path.Combine(outDir, folderName);
            _ = Directory.CreateDirectory(childDir);
            await ExtractNestedBufferAsync(data, childDir, Path.GetFileNameWithoutExtension(name), ct).ConfigureAwait(false);
            return folderName;
        }
        string ext = Formats.GuessExtension(data, Path.GetExtension(name));
        string fileName = GetUniqueName(Path.GetFileNameWithoutExtension(name), ext, outDir, counters);
        await File.WriteAllBytesAsync(Path.Combine(outDir, fileName), data, ct).ConfigureAwait(false);
        return fileName;
    }
    private async Task ExtractNestedBufferAsync(byte[] data, string outDir, string baseName, CancellationToken ct)
    {
        ArchiveType type = Formats.DetectTypeFromBuffer(data)
            ?? throw new InvalidDataException("Unknown nested archive format.");
        List<ArchiveEntry> children = ParseChildrenFromBuffer(data, baseName);
        List<string> mapperLines = new(children.Count);
        Dictionary<(string, string), int> counters = [];
        foreach (ArchiveEntry child in children)
        {
            ct.ThrowIfCancellationRequested();
            byte[] childData = data[child.Offset..(child.Offset + child.Size)];
            string outputName = await ExtractChildDataAsync(child.Name, childData, outDir, counters, ct).ConfigureAwait(false);
            mapperLines.Add($"{child.Name}={outputName}");
        }
        if (type == ArchiveType.MSND)
        {
            byte[] unknownBytes = Formats.ExtractMsndUnknownBytes(data);
            await File.WriteAllBytesAsync(Path.Combine(outDir, $"{baseName}.txt"), unknownBytes, ct).ConfigureAwait(false);
        }
        else
        {
            await File.WriteAllLinesAsync(Path.Combine(outDir, "mapper.txt"), mapperLines, Encoding.UTF8, ct).ConfigureAwait(false);
        }
    }
    private static List<ArchiveEntry> ParseChildren(byte[] data, string parentName)
    {
        ArchiveType? type = Formats.DetectTypeFromBuffer(data);
        return type.HasValue ? ParseChildrenFromBuffer(data, Path.GetFileNameWithoutExtension(parentName)) : [];
    }
    private static List<ArchiveEntry> ParseChildrenFromBuffer(byte[] data, string baseName)
    {
        ArchiveType? type = Formats.DetectTypeFromBuffer(data);
        if (type == ArchiveType.MSND && data.Length >= Formats.MsndHeader)
        {
            MsndOffsets offsets = Formats.ParseMsndOffsets(data);
            return
            [
                new ArchiveEntry { Name = $"{baseName}.sseq", Size = offsets.SseqSize, Offset = offsets.SseqOffset },
                new ArchiveEntry { Name = $"{baseName}.sbnk", Size = offsets.SbnkSize, Offset = offsets.SbnkOffset },
                new ArchiveEntry { Name = $"{baseName}.swar", Size = offsets.SwarSize, Offset = offsets.SwarOffset }
            ];
        }
        if (type == ArchiveType.DSARC && data.Length >= Formats.DsarcHeader)
        {
            int count = BitConverter.ToInt32(data, 8);
            List<ArchiveEntry> entries = new(count);
            int pos = Formats.DsarcHeader;
            for (int i = 0; i < count; i++)
            {
                ReadOnlySpan<byte> nameBytes = data.AsSpan(pos, Formats.NameSize);
                string name = Encoding.UTF8.GetString(nameBytes).Split('\0')[0].Trim();
                if (string.IsNullOrEmpty(name))
                {
                    name = $"file_{i}";
                }
                pos += Formats.NameSize;
                int size = BitConverter.ToInt32(data, pos);
                int offset = BitConverter.ToInt32(data, pos + 4);
                pos += Formats.DsarcEntryInfo;
                entries.Add(new ArchiveEntry { Name = name, Size = size, Offset = offset });
            }
            return entries;
        }
        return [];
    }
    private static string GetUniqueName(string baseName, string ext, string outDir, Dictionary<(string, string), int> counters)
    {
        (string, string) key = (baseName, ext ?? "");
        counters[key] = counters.GetValueOrDefault(key) + 1;
        int count = counters[key];
        string candidate = count == 1 ? $"{baseName}{ext}" : $"{baseName}_{count}{ext}";
        int extra = 1;
        while (File.Exists(Path.Combine(outDir, candidate)) || Directory.Exists(Path.Combine(outDir, candidate)))
        {
            candidate = count == 1 ? $"{baseName}_{extra++}{ext}" : $"{baseName}_{count}_{extra++}{ext}";
        }
        return candidate;
    }
}