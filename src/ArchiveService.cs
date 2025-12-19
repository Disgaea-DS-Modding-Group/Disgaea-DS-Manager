using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
namespace Disgaea_DS_Manager;

public interface IArchiveService
{
    Task<Collection<Entry>> LoadArchiveAsync(string archivePath, CancellationToken ct = default);
    Task SaveArchiveAsync(string path, ArchiveType type, IList<Entry> entries, string srcFolder,
        IProgress<(int current, int total)>? progress, CancellationToken ct, string? originalArchivePath = null);
    Task<ImportResult> InspectFolderForImportAsync(string folder, CancellationToken ct = default);
    Task<(string outBase, List<string> mapperLines)> ExtractAllAsync(string archivePath, ArchiveType filetype, string destFolder, IProgress<(int current, int total)>? progress = null, CancellationToken ct = default);
    Task ExtractItemAsync(string archivePath, ArchiveType filetype, Entry entry, string destFolder, CancellationToken ct = default);
    Task ExtractChunkItemAsync(string archivePath, Entry parentEntry, Entry chunkEntry, string destFolder, CancellationToken ct = default);
    Task<byte[]> ReplaceChunkItemAsync(string archivePath, Entry parentEntry, Entry chunkEntry, string replacementFilePath, string srcFolder, CancellationToken ct = default);
    Task<byte[]> RebuildNestedFromFolderAsync(string folder, CancellationToken ct = default);
    Task NestedExtractBufferAsync(byte[] buf, string outdir, string baseLabel, IProgress<(int current, int total)>? progress = null, CancellationToken ct = default);
    Task<Collection<Entry>> ParseDsarcFromBufferAsync(byte[] buf, CancellationToken ct = default);
    Task<byte[]> BuildDsarcFromFolderAsync(string folder, CancellationToken ct = default);
    Task<byte[]> BuildMsndFromFolderAsync(string folder, CancellationToken ct = default);
    Task WriteFileAsync(string path, byte[] data, CancellationToken ct = default);
    Task<byte[]> ReadFileAsync(string path, CancellationToken ct = default);
    Task<byte[]> ReadRangeAsync(string path, long offset, int size, CancellationToken ct = default);
    Task CopyFileToFolderAsync(string sourceFilePath, string destFolder, CancellationToken ct = default);
}
internal class ArchiveService : IArchiveService
{
    private readonly TupleComparer _tupleComparer = new();
    private static readonly char[] Separator = ['='];
    private static readonly string[] SourceExtensions = [".sseq", ".sbnk", ".swar"];
    private static void AppendLog(string msg)
    {
        System.Diagnostics.Debug.WriteLine($"[ArchiveService] {msg}");
    }

    public async Task<Collection<Entry>> LoadArchiveAsync(string archivePath, CancellationToken ct = default)
    {
        return await Task.Run(() =>
            {
                ArchiveType type = Detector.FromFile(archivePath);
                return type == ArchiveType.MSND
                    ? Msnd.Parse(File.ReadAllBytes(archivePath), Path.GetFileNameWithoutExtension(archivePath))
                    : Dsarc.Parse(archivePath);
            }, ct).ConfigureAwait(false);
    }

    public async Task SaveArchiveAsync(string path, ArchiveType type, IList<Entry> entries, string srcFolder,
        IProgress<(int current, int total)>? progress, CancellationToken ct, string? originalArchivePath = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        ArgumentNullException.ThrowIfNull(entries);
        if (entries.Count == 0)
        {
            throw new ArgumentException("No entries to save", nameof(entries));
        }

        await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();
            if (type == ArchiveType.MSND)
            {
                SaveMsndArchive(path, srcFolder, entries, progress, ct);
            }
            else
            {
                SaveDsarcArchive(path, srcFolder, entries, progress, ct, originalArchivePath);
            }
        }, ct).ConfigureAwait(false);
    }
    private void SaveMsndArchive(string path, string srcFolder, IList<Entry> entries,
        IProgress<(int current, int total)>? progress, CancellationToken ct)
    {
        Dictionary<string, byte[]> chunks = new(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, string> sourceFiles = GetSourceFilesByExtension(srcFolder);
        int total = Msnd.MSND_ORDER.Length;
        for (int i = 0; i < total; i++)
        {
            ct.ThrowIfCancellationRequested();
            string ext = Msnd.MSND_ORDER[i];
            string? sourcePath = FindSourceFileForExtension(ext, entries, srcFolder, sourceFiles);
            if (sourcePath is null)
            {
                throw new FileNotFoundException($"Missing {ext} file. Searched in: {srcFolder}");
            }

            chunks[ext] = File.ReadAllBytes(sourcePath);
            progress?.Report((i + 1, total));
        }
        File.WriteAllBytes(path, Msnd.Build(chunks));
    }
    private void SaveDsarcArchive(string path, string srcFolder, IList<Entry> entries,
        IProgress<(int current, int total)>? progress, CancellationToken ct, string? originalArchivePath)
    {
        string mappingPath = Path.Combine(srcFolder, "mapper.txt");
        Collection<Tuple<string, byte[]>> pairs = File.Exists(mappingPath)
            ? BuildPairsFromMapper(mappingPath, srcFolder, entries, progress, ct, originalArchivePath)
            : BuildPairsFromEntries(entries, srcFolder, progress, ct, originalArchivePath);
        File.WriteAllBytes(path, Dsarc.BuildFromPairs(pairs));
        progress?.Report((1, 1));
    }
    private Collection<Tuple<string, byte[]>> BuildPairsFromMapper(string mappingPath, string srcFolder,
        IList<Entry> entries, IProgress<(int current, int total)>? progress, CancellationToken ct, string? originalArchivePath)
    {
        string[] lines = File.ReadAllLines(mappingPath, Encoding.UTF8);
        Collection<Tuple<string, byte[]>> pairs = [];
        Dictionary<string, string> sourceFiles = GetSourceFilesByName(srcFolder);
        List<string> missingFiles = [];
        for (int i = 0; i < lines.Length; i++)
        {
            ct.ThrowIfCancellationRequested();
            string ln = lines[i];
            if (!ln.Contains('='))
            {
                progress?.Report((i + 1, lines.Length));
                continue;
            }
            string[] parts = ln.Split(Separator, 2);
            string left = parts[0].Trim();
            string right = parts[1].Trim();
            if (!TryGetFileData(left, right, srcFolder, sourceFiles, entries, originalArchivePath, out byte[]? data))
            {
                missingFiles.Add(right);
            }
            else if (data is not null)
            {
                pairs.Add(Tuple.Create(left, data));
            }
            progress?.Report((i + 1, lines.Length));
        }
        return missingFiles.Count > 0 ? throw new FileNotFoundException($"Missing files:\n{string.Join("\n", missingFiles)}") : pairs;
    }
    private Collection<Tuple<string, byte[]>> BuildPairsFromEntries(IList<Entry> entries, string srcFolder,
        IProgress<(int current, int total)>? progress, CancellationToken ct, string? originalArchivePath)
    {
        Collection<Tuple<string, byte[]>> pairs = [];
        Dictionary<string, string> sourceFiles = GetSourceFilesByName(srcFolder);
        for (int i = 0; i < entries.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            Entry entry = entries[i];
            if (!TryGetFileData(entry.Path.ToString(), entry.Path.Name, srcFolder, sourceFiles, entries, originalArchivePath, out byte[]? data))
            {
                throw new FileNotFoundException($"Missing source for: {entry.Path}");
            }

            pairs.Add(Tuple.Create(entry.Path.ToString(), data!));
            progress?.Report((i + 1, entries.Count));
        }
        return pairs;
    }
    private bool TryGetFileData(string key, string fileName, string srcFolder, Dictionary<string, string> sourceFiles,
        IList<Entry> entries, string? originalArchivePath, out byte[]? data)
    {
        data = null;
        if (sourceFiles.TryGetValue(fileName, out string? sourceFile))
        {
            data = File.ReadAllBytes(sourceFile);
            return true;
        }
        string candidate = Path.Combine(srcFolder, fileName);
        if (File.Exists(candidate))
        {
            data = File.ReadAllBytes(candidate);
            return true;
        }
        string candidateDir = Path.Combine(srcFolder, fileName);
        if (Directory.Exists(candidateDir))
        {
            data = RebuildNestedFromFolderAsync(candidateDir).GetAwaiter().GetResult();
            return data != null;
        }
        if (!string.IsNullOrEmpty(originalArchivePath) && File.Exists(originalArchivePath))
        {
            Entry? originalEntry = entries.FirstOrDefault(e =>
                string.Equals(e.Path.Name, key, StringComparison.OrdinalIgnoreCase));
            if (originalEntry is { Offset: > 0, Size: > 0 })
            {
                try
                {
                    using FileStream fs = new(originalArchivePath, FileMode.Open, FileAccess.Read);
                    _ = fs.Seek(originalEntry.Offset, SeekOrigin.Begin);
                    data = new byte[originalEntry.Size];
                    return fs.Read(data, 0, data.Length) == originalEntry.Size;
                }
                catch (Exception ex)
                {
                    AppendLog($"Failed to read original entry {key}: {ex.Message}");
                }
            }
        }
        return false;
    }
    public async Task<ImportResult> InspectFolderForImportAsync(string folder, CancellationToken ct = default)
    {
        return await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();
            if (!Directory.Exists(folder))
            {
                throw new DirectoryNotFoundException(folder);
            }

            ImportResult result = new() { SourceFolder = folder };
            string mappingFile = Path.Combine(folder, "mapper.txt");
            if (File.Exists(mappingFile))
            {
                ProcessMapperFile(mappingFile, folder, result, ct);
                result.FileType = ArchiveType.DSARC;
            }
            else
            {
                ProcessDirectoryContents(folder, result, ct);
            }
            return result;
        }, ct).ConfigureAwait(false);
    }

    private void ProcessMapperFile(string mappingFile, string folder, ImportResult result, CancellationToken ct)
    {
        foreach (string ln in File.ReadAllLines(mappingFile, Encoding.UTF8))
        {
            ct.ThrowIfCancellationRequested();
            if (!ln.Contains('='))
            {
                continue;
            }

            string[] parts = ln.Split(Separator, 2);
            string left = parts[0].Trim(), right = parts[1].Trim();
            string candidate = Path.Combine(folder, right);
            result.Entries.Add(Directory.Exists(candidate)
                ? CreateEntryFromDirectory(left, candidate, ct)
                : new Entry(new FileInfo(left), File.Exists(candidate) ? (int)new FileInfo(candidate).Length : 0));
        }
    }
    private void ProcessDirectoryContents(string folder, ImportResult result, CancellationToken ct)
    {
        List<string> allFiles = Directory.GetFiles(folder, "*", SearchOption.AllDirectories)
            .Where(f => !f.EndsWith("mapper.txt", StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f).ToList();
        List<string> stems = allFiles.Select(f => Path.GetFileNameWithoutExtension(f)).Distinct().ToList();
        HashSet<string> exts = allFiles.Select(f => Path.GetExtension(f).ToUpperInvariant()).ToHashSet();
        HashSet<string> msndOrderUpper = Msnd.MSND_ORDER.Select(x => x.ToUpperInvariant()).ToHashSet();
        if ((stems.Count == 1 && msndOrderUpper.IsSubsetOf(exts)) || exts.SetEquals(msndOrderUpper))
        {
            result.FileType = ArchiveType.MSND;
            foreach (string ext in Msnd.MSND_ORDER)
            {
                string? file = allFiles.FirstOrDefault(f =>
                    Path.GetExtension(f).Equals(ext, StringComparison.OrdinalIgnoreCase));
                if (file is null)
                {
                    throw new FileNotFoundException($"Missing MSND file {ext}");
                }

                result.Entries.Add(new Entry(new FileInfo(Path.GetFileName(file))));
            }
            return;
        }
        result.FileType = ArchiveType.DSARC;
        foreach (string item in Directory.GetFileSystemEntries(folder, "*", SearchOption.TopDirectoryOnly))
        {
            ct.ThrowIfCancellationRequested();
            if (Path.GetFileName(item)?.Equals("mapper.txt", StringComparison.OrdinalIgnoreCase) == true)
            {
                continue;
            }

            result.Entries.Add(Directory.Exists(item)
                ? CreateEntryFromDirectory(Path.GetFileName(item), item, ct)
                : new Entry(new FileInfo(GetRelativePath(folder, item))));
        }
    }
    private Entry CreateEntryFromDirectory(string name, string directory, CancellationToken ct)
    {
        Entry entry = new(new FileInfo(name));
        try
        {
            byte[] buf = RebuildNestedFromFolderAsync(directory, ct).GetAwaiter().GetResult();
            if (buf?.Length >= 4 && buf.AsSpan(0, 4).SequenceEqual(Msnd.MAGIC_MSND))
            {
                entry.IsMsnd = true;
                foreach (Entry child in Msnd.Parse(buf, Path.GetFileNameWithoutExtension(name)))
                {
                    entry.Children.Add(child);
                }
            }
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
        return entry;
    }
    public async Task<(string outBase, List<string> mapperLines)> ExtractAllAsync(string archivePath,
        ArchiveType filetype, string destFolder, IProgress<(int current, int total)>? progress = null, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(archivePath);
        ArgumentException.ThrowIfNullOrEmpty(destFolder);
        return await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();
            Collection<Entry> entries = filetype == ArchiveType.MSND
                ? Msnd.Parse(File.ReadAllBytes(archivePath), Path.GetFileNameWithoutExtension(archivePath))
                : Dsarc.Parse(archivePath);
            string baseOut = Path.Combine(destFolder, Path.GetFileNameWithoutExtension(archivePath));
            _ = Directory.CreateDirectory(baseOut);
            return filetype == ArchiveType.MSND
                ? ExtractMsndArchive(archivePath, baseOut, entries, progress, ct)
                : ExtractDsarcArchive(archivePath, baseOut, entries, progress, ct);
        }, ct).ConfigureAwait(false);
    }
    private (string outBase, List<string> mapperLines) ExtractMsndArchive(string archivePath, string baseOut,
        Collection<Entry> entries, IProgress<(int current, int total)>? progress, CancellationToken ct)
    {
        byte[] buf = File.ReadAllBytes(archivePath);
        string baseName = Path.GetFileNameWithoutExtension(archivePath);
        if (buf.Length >= Msnd.HDR_MSND)
        {
            byte[] txtData = buf.AsSpan(44, Math.Min(4, buf.Length - 44)).ToArray();
            File.WriteAllBytes(Path.Combine(baseOut, baseName + ".txt"), txtData);
        }
        for (int i = 0; i < entries.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            Entry entry = entries[i];
            byte[] data = buf.AsSpan(entry.Offset, entry.Size).ToArray();
            string target = Path.Combine(baseOut, entry.Path.Name);
            _ = Directory.CreateDirectory(Path.GetDirectoryName(target) ?? baseOut);
            File.WriteAllBytes(target, data);
            progress?.Report((i + 1, entries.Count));
        }
        return (baseOut, []);
    }
    private (string outBase, List<string> mapperLines) ExtractDsarcArchive(string archivePath, string baseOut,
        Collection<Entry> entries, IProgress<(int current, int total)>? progress, CancellationToken ct)
    {
        long archiveSize = new FileInfo(archivePath).Length;
        int dataStart = Dsarc.HDR_DSARC + (entries.Count * (Helpers.NAMESZ + Dsarc.ENTRY_INFO_SZ));
        var ranges = entries.Select((e, idx) => new { Entry = e, Index = idx, Start = (long)e.Offset, End = (long)e.Offset + e.Size })
            .Where(x => x.Start >= dataStart && x.End <= archiveSize && x.Entry.Size >= 0 && x.End >= x.Start)
            .OrderBy(x => x.Start).ToList();
        string[] extractedNames = new string[entries.Count];
        Dictionary<Tuple<string, string>, int> counters = new(_tupleComparer);
        using FileStream fs = new(archivePath, FileMode.Open, FileAccess.Read);
        for (int i = 0; i < ranges.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var range = ranges[i];
            _ = fs.Seek(range.Start, SeekOrigin.Begin);
            byte[] data = new byte[range.End - range.Start];
            _ = fs.Read(data, 0, data.Length);
            string ext = Helpers.GuessExtByMagic(data, range.Entry.Path.Extension);
            string finalName = GetUniqueOutputName(Path.GetFileNameWithoutExtension(range.Entry.Path.Name), ext, baseOut, counters);
            string target = Path.Combine(baseOut, finalName);
            _ = Directory.CreateDirectory(Path.GetDirectoryName(target) ?? baseOut);
            File.WriteAllBytes(target, data);
            extractedNames[range.Index] = finalName;
            progress?.Report((i + 1, ranges.Count));
        }
        List<string> mappingLines = entries.Select((e, idx) => $"{e.Path.Name}={extractedNames[idx] ?? e.Path.Name}").ToList();
        File.WriteAllText(Path.Combine(baseOut, "mapper.txt"), string.Join("\n", mappingLines), Encoding.UTF8);
        return (baseOut, mappingLines);
    }
    public async Task ExtractItemAsync(string archivePath, ArchiveType filetype, Entry entry, string destFolder, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(archivePath);
        ArgumentNullException.ThrowIfNull(entry);
        ArgumentException.ThrowIfNullOrEmpty(destFolder);
        await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();
            byte[] data = filetype == ArchiveType.MSND
                ? File.ReadAllBytes(archivePath).AsSpan(entry.Offset, entry.Size).ToArray()
                : ReadFileRange(archivePath, entry.Offset, entry.Size);
            string target = Path.Combine(destFolder, filetype == ArchiveType.MSND
                ? entry.Path.Name
                : Path.GetFileNameWithoutExtension(entry.Path.Name) + Helpers.GuessExtByMagic(data, entry.Path.Extension));
            _ = Directory.CreateDirectory(Path.GetDirectoryName(target) ?? destFolder);
            File.WriteAllBytes(target, data);
        }, ct).ConfigureAwait(false);
    }
    public async Task ExtractChunkItemAsync(string archivePath, Entry parentEntry, Entry chunkEntry, string destFolder, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(parentEntry);
        ArgumentNullException.ThrowIfNull(chunkEntry);
        ArgumentException.ThrowIfNullOrEmpty(archivePath);
        ArgumentException.ThrowIfNullOrEmpty(destFolder);
        await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();
            byte[] msndBuf = ReadFileRange(archivePath, parentEntry.Offset, parentEntry.Size);
            byte[] data = msndBuf.AsSpan(chunkEntry.Offset, chunkEntry.Size).ToArray();
            string target = Path.Combine(destFolder, chunkEntry.Path.Name);
            _ = Directory.CreateDirectory(Path.GetDirectoryName(target) ?? destFolder);
            File.WriteAllBytes(target, data);
        }, ct).ConfigureAwait(false);
    }
    public async Task<byte[]> ReplaceChunkItemAsync(string archivePath, Entry parentEntry, Entry chunkEntry,
        string replacementFilePath, string srcFolder, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(parentEntry);
        ArgumentNullException.ThrowIfNull(chunkEntry);
        ArgumentException.ThrowIfNullOrEmpty(archivePath);
        ArgumentException.ThrowIfNullOrEmpty(replacementFilePath);
        return await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();
            byte[] newData = File.ReadAllBytes(replacementFilePath);
            byte[] msndBuf = ReadFileRange(archivePath, parentEntry.Offset, parentEntry.Size);
            byte[] rebuilt = Msnd.ReplaceChunk(msndBuf, Path.GetExtension(chunkEntry.Path.Name).ToLowerInvariant(), newData);
            if (!string.IsNullOrEmpty(srcFolder))
            {
                string outPath = Path.Combine(srcFolder, parentEntry.Path.Name);
                _ = Directory.CreateDirectory(Path.GetDirectoryName(outPath) ?? srcFolder);
                File.WriteAllBytes(outPath, rebuilt);
            }
            return rebuilt;
        }, ct).ConfigureAwait(false);
    }
    public async Task<byte[]> RebuildNestedFromFolderAsync(string folder, CancellationToken ct = default)
    {
        return await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();
            if (!Directory.Exists(folder))
            {
                throw new DirectoryNotFoundException(folder);
            }

            string mapperPath = Path.Combine(folder, "mapper.txt");
            return File.Exists(mapperPath)
                ? BuildFromMapper(mapperPath, folder, ct)
                : BuildFromFolderContents(folder);
        }, ct).ConfigureAwait(false);
    }

    private byte[] BuildFromMapper(string mapperPath, string folder, CancellationToken ct)
    {
        Collection<Tuple<string, byte[]>> pairs = [];
        foreach (string ln in File.ReadAllLines(mapperPath, Encoding.UTF8))
        {
            ct.ThrowIfCancellationRequested();
            if (!ln.Contains('='))
            {
                continue;
            }

            string[] parts = ln.Split(Separator, 2);
            string left = parts[0].Trim(), right = parts[1].Trim();
            string candidatePath = Path.Combine(folder, right);
            pairs.Add(Tuple.Create(left, Directory.Exists(candidatePath)
                ? RebuildNestedFromFolderAsync(candidatePath, ct).GetAwaiter().GetResult()
                : File.ReadAllBytes(FindFile(folder, right))));
        }
        return Dsarc.BuildFromPairs(pairs);
    }
    private byte[] BuildFromFolderContents(string folder)
    {
        string baseName = Path.GetFileName(folder);
        string?[] msndFiles = SourceExtensions.Select(ext =>
            Directory.GetFiles(folder, $"*{ext}", SearchOption.TopDirectoryOnly).FirstOrDefault()).ToArray();
        if (msndFiles.Any(f => !string.IsNullOrEmpty(f)))
        {
            Dictionary<string, byte[]> chunks = new(StringComparer.OrdinalIgnoreCase);
            foreach (string ext in Msnd.MSND_ORDER)
            {
                string exact = Path.Combine(folder, baseName + ext);
                string? chosen = File.Exists(exact) ? exact : Directory.GetFiles(folder, $"*{ext}", SearchOption.TopDirectoryOnly).FirstOrDefault();
                if (chosen is null)
                {
                    throw new FileNotFoundException($"Missing MSND file {ext} in {folder}");
                }

                chunks[ext] = File.ReadAllBytes(chosen);
            }
            string txtPath = Path.Combine(folder, baseName + ".txt");
            byte[]? txtBytes = File.Exists(txtPath) ? File.ReadAllBytes(txtPath) : null;
            return Msnd.Build(chunks, txtBytes);
        }
        string[] files = Directory.GetFiles(folder, "*", SearchOption.TopDirectoryOnly);
        return files.Length == 1 ? File.ReadAllBytes(files[0])
            : throw new InvalidOperationException($"Cannot determine archive type at {folder}");
    }
    public async Task NestedExtractBufferAsync(byte[] buf, string outdir, string baseLabel,
        IProgress<(int current, int total)>? progress = null, CancellationToken ct = default)
    {
        await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();
            if (buf is null || outdir is null)
            {
                return;
            }

            _ = Directory.CreateDirectory(outdir);
            if (buf.Length >= 4 && buf.AsSpan(0, 4).SequenceEqual(Msnd.MAGIC_MSND))
            {
                ExtractMsndBuffer(buf, outdir, baseLabel, progress, ct);
            }
            else if (buf.Length >= 8 && buf.AsSpan(0, 8).SequenceEqual(Dsarc.MAGIC_DSARC))
            {
                ExtractDsarcBuffer(buf, outdir, baseLabel, progress, ct);
            }
            else
            {
                string fallbackName = GetUniqueOutputName(baseLabel, Path.GetExtension(baseLabel), outdir, new Dictionary<Tuple<string, string>, int>(_tupleComparer));
                File.WriteAllBytes(Path.Combine(outdir, fallbackName), buf);
            }
        }, ct).ConfigureAwait(false);
    }
    private void ExtractMsndBuffer(byte[] buf, string outdir, string baseLabel,
        IProgress<(int current, int total)>? progress, CancellationToken ct)
    {
        if (buf.Length >= Msnd.HDR_MSND)
        {
            byte[] txtData = buf.AsSpan(44, Math.Min(4, buf.Length - 44)).ToArray();
            File.WriteAllBytes(Path.Combine(outdir, baseLabel + ".txt"), txtData);
        }
        Collection<Entry> children = Msnd.Parse(buf, baseLabel);
        Dictionary<Tuple<string, string>, int> counters = new(_tupleComparer);
        for (int i = 0; i < children.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            Entry child = children[i];
            byte[] data = buf.AsSpan(child.Offset, child.Size).ToArray();
            _ = ProcessChildEntry(data, child, outdir, counters, progress, ct, i, children.Count);
        }
    }
    private void ExtractDsarcBuffer(byte[] buf, string outdir, string baseLabel,
        IProgress<(int current, int total)>? progress, CancellationToken ct)
    {
        Collection<Entry> parsed = Dsarc.ParseFromBuffer(buf);
        Dictionary<Tuple<string, string>, int> counters = new(_tupleComparer);
        List<string> mappingLines = [];
        for (int i = 0; i < parsed.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            Entry entry = parsed[i];
            byte[] data = buf.AsSpan(entry.Offset, entry.Size).ToArray();
            string outputName = ProcessChildEntry(data, entry, outdir, counters, progress, ct, i, parsed.Count);
            mappingLines.Add($"{entry.Path.Name}={outputName}");
        }
        File.WriteAllText(Path.Combine(outdir, "mapper.txt"), string.Join("\n", mappingLines), Encoding.UTF8);
    }
    private string ProcessChildEntry(byte[] data, Entry entry, string outdir,
        Dictionary<Tuple<string, string>, int> counters, IProgress<(int current, int total)>? progress,
        CancellationToken ct, int index, int total)
    {
        bool isDsarc = data.Length >= 8 && data.AsSpan(0, 8).SequenceEqual(Dsarc.MAGIC_DSARC);
        bool isMsnd = data.Length >= 4 && data.AsSpan(0, 4).SequenceEqual(Msnd.MAGIC_MSND);
        if (isDsarc || isMsnd)
        {
            string folderName = GetUniqueOutputName(Path.GetFileNameWithoutExtension(entry.Path.Name), string.Empty, outdir, counters);
            string childFolder = Path.Combine(outdir, folderName);
            _ = Directory.CreateDirectory(childFolder);
            NestedExtractBufferAsync(data, childFolder, Path.GetFileNameWithoutExtension(entry.Path.Name), progress, ct).GetAwaiter().GetResult();
            File.WriteAllBytes(Path.Combine(childFolder, isMsnd ? Path.GetFileNameWithoutExtension(entry.Path.Name) + ".txt" : "mapper.txt"), []);
            progress?.Report((index + 1, total));
            return folderName;
        }
        else
        {
            string ext = Helpers.GuessExtByMagic(data, Path.GetExtension(entry.Path.Name));
            string finalName = GetUniqueOutputName(Path.GetFileNameWithoutExtension(entry.Path.Name), ext, outdir, counters);
            File.WriteAllBytes(Path.Combine(outdir, finalName), data);
            progress?.Report((index + 1, total));
            return finalName;
        }
    }
    public Task<Collection<Entry>> ParseDsarcFromBufferAsync(byte[] buf, CancellationToken ct = default)
    {
        return Task.Run(() => Dsarc.ParseFromBuffer(buf), ct);
    }

    public Task<byte[]> BuildDsarcFromFolderAsync(string folder, CancellationToken ct = default)
    {
        return Task.Run(() => BuildFromMapper(Path.Combine(folder, "mapper.txt"), folder, ct), ct);
    }

    public Task<byte[]> BuildMsndFromFolderAsync(string folder, CancellationToken ct = default)
    {
        return Task.Run(() => BuildFromFolderContents(folder), ct);
    }

    public async Task WriteFileAsync(string path, byte[] data, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();
            _ = Directory.CreateDirectory(Path.GetDirectoryName(path) ?? ".");
            File.WriteAllBytes(path, data);
        }, ct).ConfigureAwait(false);
    }
    public async Task<byte[]> ReadFileAsync(string path, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        return await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();
            return File.ReadAllBytes(path);
        }, ct).ConfigureAwait(false);
    }
    public async Task<byte[]> ReadRangeAsync(string path, long offset, int size, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        return await Task.Run(() => ReadFileRange(path, offset, size), ct).ConfigureAwait(false);
    }
    public async Task CopyFileToFolderAsync(string sourceFilePath, string destFolder, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(sourceFilePath);
        ArgumentException.ThrowIfNullOrEmpty(destFolder);
        await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();
            _ = Directory.CreateDirectory(destFolder);
            File.Copy(sourceFilePath, Path.Combine(destFolder, Path.GetFileName(sourceFilePath)), true);
        }, ct).ConfigureAwait(false);
    }
    private static byte[] ReadFileRange(string path, long offset, int size)
    {
        using FileStream fs = new(path, FileMode.Open, FileAccess.Read);
        _ = fs.Seek(offset, SeekOrigin.Begin);
        byte[] buf = new byte[size];
        int read = fs.Read(buf, 0, size);
        return read == size ? buf : buf[..read];
    }
    private static Dictionary<string, string> GetSourceFilesByExtension(string folder)
    {
        return Directory.Exists(folder)
            ? Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories)
                .Where(f => Msnd.MSND_ORDER.Contains(Path.GetExtension(f).ToLowerInvariant()))
                .ToDictionary(f => Path.GetExtension(f).ToLowerInvariant(), f => f)
            : [];
    }

    private static Dictionary<string, string> GetSourceFilesByName(string folder)
    {
        return Directory.Exists(folder)
                ? Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories)
                    .ToDictionary(f => Path.GetFileName(f), f => f, StringComparer.OrdinalIgnoreCase)
                : [];
    }

    private static string? FindSourceFileForExtension(string ext, IList<Entry> entries, string srcFolder, Dictionary<string, string> sourceFiles)
    {
        Entry? entryForExt = entries.FirstOrDefault(e =>
            Path.GetExtension(e.Path.Name).Equals(ext, StringComparison.OrdinalIgnoreCase));
        if (entryForExt != null)
        {
            string candidate = Path.Combine(srcFolder, entryForExt.Path.Name);
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }
        return sourceFiles.GetValueOrDefault(ext) ?? Directory.GetFiles(srcFolder, $"*{ext}", SearchOption.AllDirectories).FirstOrDefault();
    }
    private static string FindFile(string folder, string fileName)
    {
        string candidate = Path.Combine(folder, fileName);
        return File.Exists(candidate) ? candidate
            : Directory.GetFiles(folder, Path.GetFileName(fileName), SearchOption.AllDirectories).FirstOrDefault()
            ?? throw new FileNotFoundException($"File not found: {fileName}");
    }
    private static string GetUniqueOutputName(string baseName, string ext, string outdir, Dictionary<Tuple<string, string>, int> counters)
    {
        Tuple<string, string> key = Tuple.Create(baseName, ext ?? string.Empty);
        counters[key] = counters.GetValueOrDefault(key) + 1;
        int count = counters[key];
        string candidate = count == 1 ? $"{baseName}{ext}" : $"{baseName}_{count}{ext}";
        string finalName = candidate;
        int extra = 1;
        while (File.Exists(Path.Combine(outdir, finalName)) || Directory.Exists(Path.Combine(outdir, finalName)))
        {
            finalName = count == 1 ? $"{baseName}_{extra++}{ext}" : $"{baseName}_{count}_{extra++}{ext}";
        }
        return finalName;
    }
    private static string GetRelativePath(string baseDir, string fullPath)
    {
        Uri baseUri = new(AppendDirectorySeparator(baseDir));
        Uri fullUri = new(fullPath);
        return Uri.UnescapeDataString(baseUri.MakeRelativeUri(fullUri).ToString().Replace('/', Path.DirectorySeparatorChar));
    }
    private static string AppendDirectorySeparator(string path)
    {
        return path.EndsWith(Path.DirectorySeparatorChar) ? path : path + Path.DirectorySeparatorChar;
    }
}