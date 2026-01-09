using Disgaea_DS_Manager.Models;
using System.Text;
namespace Disgaea_DS_Manager.Services;

public sealed class FolderImportService
{
    private static readonly char[] Separator = ['='];
    private static readonly HashSet<string> MsndExtensionsSet = new(Formats.MsndOrder, StringComparer.OrdinalIgnoreCase);
    public async Task<ImportResult> AnalyzeFolderAsync(string folder, CancellationToken ct = default)
    {
        if (!Directory.Exists(folder))
        {
            throw new DirectoryNotFoundException($"Folder not found: {folder}");
        }
        ImportResult result = new() { SourceFolder = folder };
        string mapperPath = Path.Combine(folder, "mapper.txt");
        if (File.Exists(mapperPath))
        {
            await ProcessMapperFileAsync(mapperPath, folder, result, ct).ConfigureAwait(false);
            result.FileType = ArchiveType.DSARC;
        }
        else
        {
            await ProcessDirectoryAsync(folder, result, ct).ConfigureAwait(false);
        }
        return result;
    }
    private async Task ProcessMapperFileAsync(string mapperPath, string folder, ImportResult result, CancellationToken ct)
    {
        string[] lines = await File.ReadAllLinesAsync(mapperPath, Encoding.UTF8, ct).ConfigureAwait(false);
        foreach (string line in lines)
        {
            ct.ThrowIfCancellationRequested();
            if (!line.Contains('='))
            {
                continue;
            }
            string[] parts = line.Split(Separator, 2);
            string archiveName = parts[0].Trim();
            string sourceName = parts[1].Trim();
            string sourcePath = Path.Combine(folder, sourceName);
            ArchiveEntry entry = await CreateEntryFromSourceAsync(archiveName, sourcePath, ct).ConfigureAwait(false);
            result.Entries.Add(entry);
        }
    }
    private async Task ProcessDirectoryAsync(string folder, ImportResult result, CancellationToken ct)
    {
        List<string> allFiles = Directory.GetFiles(folder, "*", SearchOption.AllDirectories)
            .Where(f => !f.EndsWith("mapper.txt", StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f)
            .ToList();
        HashSet<string> extensions = allFiles.Select(f => Path.GetExtension(f).ToUpperInvariant()).ToHashSet();
        HashSet<string> msndExtensions = Formats.MsndOrder.Select(x => x.ToUpperInvariant()).ToHashSet();
        bool isMsndStructure = extensions.SetEquals(msndExtensions) ||
            (allFiles.Select(Path.GetFileNameWithoutExtension).Distinct().Count() == 1 && msndExtensions.IsSubsetOf(extensions));
        if (isMsndStructure)
        {
            result.FileType = ArchiveType.MSND;
            foreach (string ext in Formats.MsndOrder)
            {
                ct.ThrowIfCancellationRequested();
                string file = allFiles.FirstOrDefault(f => Path.GetExtension(f).Equals(ext, StringComparison.OrdinalIgnoreCase))
                    ?? throw new FileNotFoundException($"Missing MSND file: {ext}");
                byte[] data = await File.ReadAllBytesAsync(file, ct).ConfigureAwait(false);
                result.Entries.Add(new ArchiveEntry
                {
                    Name = Path.GetFileName(file),
                    Size = data.Length,
                    ImportOrder = result.Entries.Count,
                    DataSource = new BufferSource(data)
                });
            }
            return;
        }
        result.FileType = ArchiveType.DSARC;
        IOrderedEnumerable<string> topItems = Directory.GetFileSystemEntries(folder, "*", SearchOption.TopDirectoryOnly)
            .Where(p => !Path.GetFileName(p).Equals("mapper.txt", StringComparison.OrdinalIgnoreCase))
            .OrderBy(p => p);
        int importOrder = 0;
        foreach (string item in topItems)
        {
            ct.ThrowIfCancellationRequested();
            ArchiveEntry entry = await CreateEntryFromSourceAsync(Path.GetFileName(item), item, ct).ConfigureAwait(false);
            entry.ImportOrder = importOrder++;
            result.Entries.Add(entry);
        }
    }
    private async Task<ArchiveEntry> CreateEntryFromSourceAsync(string name, string sourcePath, CancellationToken ct)
    {
        if (Directory.Exists(sourcePath))
        {
            return await CreateEntryFromDirectoryAsync(name, sourcePath, ct).ConfigureAwait(false);
        }
        if (!File.Exists(sourcePath))
        {
            return new ArchiveEntry { Name = name, Size = 0 };
        }
        byte[] data = await File.ReadAllBytesAsync(sourcePath, ct).ConfigureAwait(false);
        return new ArchiveEntry { Name = name, Size = data.Length, DataSource = new BufferSource(data) };
    }
    private async Task<ArchiveEntry> CreateEntryFromDirectoryAsync(string name, string directory, CancellationToken ct)
    {
        ArchiveEntry entry = new() { Name = name, NestedType = DetectContainerType(directory) };
        try
        {
            byte[] data = await BuildNestedArchiveAsync(directory, ct).ConfigureAwait(false);
            entry.Size = data.Length;
            entry.DataSource = new BufferSource(data);
            ArchiveType? detectedType = Formats.DetectTypeFromBuffer(data);
            if (detectedType == ArchiveType.MSND)
            {
                entry.NestedType = ArchiveType.MSND;
                Formats.PopulateMsndChildren(data, entry, Path.GetFileNameWithoutExtension(entry.Name));
            }
            else if (detectedType == ArchiveType.DSARC)
            {
                entry.NestedType = ArchiveType.DSARC;
                PopulateDsarcChildren(data, entry);
            }
        }
        catch { }
        return entry;
    }
    private static void PopulateDsarcChildren(byte[] dsarcData, ArchiveEntry parent)
    {
        try
        {
            if (dsarcData.Length < Formats.DsarcHeader || !dsarcData.AsSpan(0, 8).SequenceEqual(Formats.MagicDsarc))
            {
                return;
            }
            int count = BitConverter.ToInt32(dsarcData, 8);
            int version = BitConverter.ToInt32(dsarcData, 12);
            if (version != Formats.DsarcVersion)
            {
                return;
            }
            parent.Children.Clear();
            int pos = Formats.DsarcHeader;
            for (int i = 0; i < count; i++)
            {
                ReadOnlySpan<byte> nameBytes = dsarcData.AsSpan(pos, Formats.NameSize);
                string name = System.Text.Encoding.UTF8.GetString(nameBytes).Split('\0')[0].Trim();
                if (string.IsNullOrEmpty(name))
                {
                    name = $"file_{i}";
                }
                pos += Formats.NameSize;
                int size = BitConverter.ToInt32(dsarcData, pos);
                int offset = BitConverter.ToInt32(dsarcData, pos + 4);
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
                    DataSource = new BufferSource(dsarcData[offset..(offset + size)])
                };
                ReadOnlySpan<byte> entryData = dsarcData.AsSpan(offset, Math.Min(size, 8));
                if (size >= 4 && entryData[..Math.Min(4, entryData.Length)].SequenceEqual(Formats.MagicMsnd))
                {
                    child.NestedType = ArchiveType.MSND;
                    string baseName = Path.GetFileNameWithoutExtension(name);
                    Formats.PopulateMsndChildren(dsarcData[offset..(offset + size)], child, baseName);
                }
                else if (size >= 8 && entryData[..Math.Min(8, entryData.Length)].SequenceEqual(Formats.MagicDsarc))
                {
                    child.NestedType = ArchiveType.DSARC;
                }
                parent.Children.Add(child);
            }
        }
        catch { }
    }
    private async Task<byte[]> BuildNestedArchiveAsync(string folder, CancellationToken ct)
    {
        List<string> allFiles = Directory.GetFiles(folder, "*", SearchOption.AllDirectories)
            .Where(f => !f.EndsWith("mapper.txt", StringComparison.OrdinalIgnoreCase) && !f.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
            .ToList();
        HashSet<string> extensions = allFiles.Select(f => Path.GetExtension(f).ToUpperInvariant()).ToHashSet();
        HashSet<string> msndExtensions = Formats.MsndOrder.Select(x => x.ToUpperInvariant()).ToHashSet();
        if (extensions.SetEquals(msndExtensions) || msndExtensions.IsSubsetOf(extensions))
        {
            Dictionary<string, byte[]> chunks = new(StringComparer.OrdinalIgnoreCase);
            foreach (string ext in Formats.MsndOrder)
            {
                string file = allFiles.FirstOrDefault(f => Path.GetExtension(f).Equals(ext, StringComparison.OrdinalIgnoreCase))
                    ?? throw new FileNotFoundException($"Missing: {ext}");
                chunks[ext] = await File.ReadAllBytesAsync(file, ct).ConfigureAwait(false);
            }
            byte[]? unknownBytes = await ReadMsndUnknownBytesAsync(folder, ct).ConfigureAwait(false);
            return Formats.BuildMsnd(chunks, unknownBytes);
        }
        string mapperPath = Path.Combine(folder, "mapper.txt");
        return File.Exists(mapperPath)
            ? await BuildDsarcFromMapperAsync(mapperPath, folder, ct).ConfigureAwait(false)
            : await BuildDsarcFromFilesAsync(folder, ct).ConfigureAwait(false);
    }
    private static async Task<byte[]?> ReadMsndUnknownBytesAsync(string folder, CancellationToken ct)
    {
        List<string> txtFiles = Directory.GetFiles(folder, "*.txt", SearchOption.TopDirectoryOnly)
            .Where(f => !Path.GetFileName(f).Equals("mapper.txt", StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (txtFiles.Count == 0)
        {
            return null;
        }
        byte[] data = await File.ReadAllBytesAsync(txtFiles[0], ct).ConfigureAwait(false);
        if (data.Length >= 4)
        {
            return data[..4];
        }
        if (data.Length > 0)
        {
            byte[] padded = new byte[4];
            data.CopyTo(padded, 0);
            return padded;
        }
        return null;
    }
    private async Task<byte[]> BuildDsarcFromMapperAsync(string mapperPath, string folder, CancellationToken ct)
    {
        string[] lines = await File.ReadAllLinesAsync(mapperPath, Encoding.UTF8, ct).ConfigureAwait(false);
        List<(string Name, byte[] Data)> pairs = new(lines.Length);
        foreach (string line in lines)
        {
            if (!line.Contains('='))
            {
                continue;
            }
            string[] parts = line.Split(Separator, 2);
            string archiveName = parts[0].Trim();
            string sourceName = parts[1].Trim();
            string sourcePath = Path.Combine(folder, sourceName);
            byte[] data = Directory.Exists(sourcePath)
                ? await BuildNestedArchiveAsync(sourcePath, ct).ConfigureAwait(false)
                : File.Exists(sourcePath)
                ? await File.ReadAllBytesAsync(sourcePath, ct).ConfigureAwait(false)
                : throw new FileNotFoundException($"Source not found: {sourcePath}");
            pairs.Add((archiveName, data));
        }
        return BuildDsarc(pairs);
    }
    private async Task<byte[]> BuildDsarcFromFilesAsync(string folder, CancellationToken ct)
    {
        IEnumerable<string> items = Directory.GetFileSystemEntries(folder, "*", SearchOption.TopDirectoryOnly)
            .Where(p => !Path.GetFileName(p).Equals("mapper.txt", StringComparison.OrdinalIgnoreCase));
        List<(string Name, byte[] Data)> pairs = [];
        foreach (string item in items)
        {
            ct.ThrowIfCancellationRequested();
            string name = Path.GetFileName(item);
            byte[] data = Directory.Exists(item)
                ? await BuildNestedArchiveAsync(item, ct).ConfigureAwait(false)
                : await File.ReadAllBytesAsync(item, ct).ConfigureAwait(false);
            pairs.Add((name, data));
        }
        return BuildDsarc(pairs);
    }
    private static byte[] BuildDsarc(List<(string Name, byte[] Data)> pairs)
    {
        int count = pairs.Count;
        int headerSize = Formats.DsarcHeader + (count * (Formats.NameSize + Formats.DsarcEntryInfo));
        int totalDataSize = pairs.Sum(p => p.Data.Length);
        using MemoryStream ms = new(headerSize + totalDataSize);
        using BinaryWriter bw = new(ms);
        bw.Write(Formats.MagicDsarc);
        bw.Write(count);
        bw.Write(Formats.DsarcVersion);
        int offset = headerSize;
        foreach ((string name, byte[] data) in pairs)
        {
            bw.Write(Formats.PadName(name));
            bw.Write(data.Length);
            bw.Write(offset);
            offset += data.Length;
        }
        foreach ((_, byte[] data) in pairs)
        {
            bw.Write(data);
        }
        return ms.ToArray();
    }
    public static ArchiveType DetectContainerType(string directoryPath)
    {
        if (!Directory.Exists(directoryPath))
        {
            return ArchiveType.DSARC;
        }
        try
        {
            List<string> files = Directory.GetFiles(directoryPath, "*", SearchOption.TopDirectoryOnly)
                .Where(f => !f.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                .ToList();
            HashSet<string> extensions = files.Select(f => Path.GetExtension(f).ToLowerInvariant()).ToHashSet();
            if (extensions.Count > 0 && extensions.All(MsndExtensionsSet.Contains))
            {
                if (MsndExtensionsSet.IsSubsetOf(extensions) || extensions.SetEquals(MsndExtensionsSet))
                {
                    return ArchiveType.MSND;
                }
            }
            List<string> txtFiles = Directory.GetFiles(directoryPath, "*.txt", SearchOption.TopDirectoryOnly)
                .Where(f => !Path.GetFileName(f).Equals("mapper.txt", StringComparison.OrdinalIgnoreCase))
                .ToList();
            HashSet<string> msndExtUpper = Formats.MsndOrder.Select(x => x.ToLowerInvariant()).ToHashSet();
            if (txtFiles.Count > 0 && extensions.SetEquals(msndExtUpper))
            {
                return ArchiveType.MSND;
            }
        }
        catch { }
        return ArchiveType.DSARC;
    }
}