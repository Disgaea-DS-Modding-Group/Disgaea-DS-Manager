using Disgaea_DS_Manager.Models;
using Disgaea_DS_Manager.Services;
namespace Disgaea_DS_Manager.Commands;

public interface ICommand
{
    Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default);
}
public static class EntryFinder
{
    public static ArchiveEntry? FindById(ArchiveEntry root, long id)
    {
        if (root.Id == id)
        {
            return root;
        }
        foreach (ArchiveEntry child in root.Children)
        {
            ArchiveEntry? found = FindById(child, id);
            if (found is not null)
            {
                return found;
            }
        }
        return null;
    }
    public static ArchiveEntry? FindParent(ArchiveEntry root, long childId)
    {
        foreach (ArchiveEntry child in root.Children)
        {
            if (child.Id == childId)
            {
                return root;
            }
            ArchiveEntry? found = FindParent(child, childId);
            if (found is not null)
            {
                return found;
            }
        }
        return null;
    }
    public static bool RemoveById(ArchiveEntry parent, long id)
    {
        for (int i = 0; i < parent.Children.Count; i++)
        {
            if (parent.Children[i].Id == id)
            {
                parent.Children.RemoveAt(i);
                return true;
            }
            if (RemoveById(parent.Children[i], id))
            {
                return true;
            }
        }
        return false;
    }
    public static bool HasDuplicateName(ArchiveEntry parent, string name, long excludeId)
    {
        return parent.Children.Any(c => c.Id != excludeId && c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    }
    public static void MarkParentsModified(ArchiveEntry root, long childId)
    {
        ArchiveEntry? parent = FindParent(root, childId);
        if (parent is not null && parent != root)
        {
            parent.IsModified = true;
            MarkParentsModified(root, parent.Id);
        }
    }
    public static void ClearModifiedFlags(ArchiveEntry entry)
    {
        entry.IsModified = false;
        foreach (ArchiveEntry child in entry.Children)
        {
            ClearModifiedFlags(child);
        }
    }
    public static void PopulateMsndChildrenWithModified(byte[] msndData, ArchiveEntry entry, bool markModified = true, string? replacedExt = null)
    {
        string baseName = Path.GetFileNameWithoutExtension(entry.Name);
        MsndOffsets offsets = Formats.ParseMsndOffsets(msndData);
        entry.Children.Clear();
        entry.Children.Add(CreateChild($"{baseName}.sseq", msndData, offsets.SseqOffset, offsets.SseqSize, markModified || replacedExt == ".sseq"));
        entry.Children.Add(CreateChild($"{baseName}.sbnk", msndData, offsets.SbnkOffset, offsets.SbnkSize, markModified || replacedExt == ".sbnk"));
        entry.Children.Add(CreateChild($"{baseName}.swar", msndData, offsets.SwarOffset, offsets.SwarSize, markModified || replacedExt == ".swar"));
    }
    private static ArchiveEntry CreateChild(string name, byte[] data, int offset, int size, bool modified)
    {
        return new()
        {
            Name = name,
            Size = size,
            Offset = offset,
            DataSource = new BufferSource(data[offset..(offset + size)]),
            IsModified = modified
        };
    }
}
public sealed class LoadArchiveCommand(ArchiveReaderService reader, string path) : ICommand
{
    public async Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        ArchiveDocument doc = await reader.ParseAsync(path, ct).ConfigureAwait(false);
        return doc with { HasContent = true };
    }
}
public sealed class ImportFolderCommand(FolderImportService importer, string folderPath) : ICommand
{
    public async Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        ImportResult result = await importer.AnalyzeFolderAsync(folderPath, ct).ConfigureAwait(false);
        return new ArchiveDocument
        {
            FileType = result.FileType,
            IsModified = true,
            HasContent = true,
            RootEntry = new ArchiveEntry { Name = Path.GetFileName(folderPath), Children = result.Entries }
        };
    }
}
public sealed class ReplaceEntryCommand(ArchiveDocument document, long targetId, byte[] newData) : ICommand
{
    public Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        ArchiveDocument clone = document.DeepCopy();
        ArchiveEntry entry = EntryFinder.FindById(clone.RootEntry, targetId)
            ?? throw new InvalidOperationException("Entry not found.");
        ArchiveType? detectedType = Formats.DetectTypeFromBuffer(newData);
        entry.DataSource = new BufferSource(newData);
        entry.Size = newData.Length;
        entry.IsModified = true;
        if (detectedType == ArchiveType.MSND && newData.Length >= Formats.MsndHeader)
        {
            entry.NestedType = ArchiveType.MSND;
            EntryFinder.PopulateMsndChildrenWithModified(newData, entry);
        }
        else if (detectedType == ArchiveType.DSARC)
        {
            entry.NestedType = ArchiveType.DSARC;
        }
        else if (entry.IsNested && entry.Children.Count > 0)
        {
            entry.NestedType = null;
            entry.Children.Clear();
        }
        EntryFinder.MarkParentsModified(clone.RootEntry, targetId);
        return Task.FromResult(clone with { IsModified = true });
    }
}
public sealed class ReplaceChunkCommand(
    ArchiveDocument document,
    long parentId,
    long chunkId,
    byte[] newChunkData,
    ArchiveWriterService writer) : ICommand
{
    public async Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        ArchiveDocument clone = document.DeepCopy();
        ArchiveEntry parent = EntryFinder.FindById(clone.RootEntry, parentId)
            ?? throw new InvalidOperationException("Parent not found.");
        ArchiveEntry chunk = EntryFinder.FindById(clone.RootEntry, chunkId)
            ?? throw new InvalidOperationException("Chunk not found.");
        byte[] msndData = await parent.GetDataAsync(ct).ConfigureAwait(false);
        if (msndData.Length < Formats.MsndHeader || !msndData.AsSpan(0, 4).SequenceEqual(Formats.MagicMsnd))
        {
            throw new InvalidDataException("Invalid MSND buffer.");
        }
        string replacedExt = Path.GetExtension(chunk.Name).ToLowerInvariant();
        byte[] newMsndData = writer.ReplaceChunk(msndData, replacedExt, newChunkData);
        parent.DataSource = new BufferSource(newMsndData);
        parent.Size = newMsndData.Length;
        parent.IsModified = true;
        EntryFinder.PopulateMsndChildrenWithModified(newMsndData, parent, false, replacedExt);
        EntryFinder.MarkParentsModified(clone.RootEntry, parentId);
        return clone with { IsModified = true };
    }
}
public sealed class SaveArchiveCommand(
    ArchiveWriterService writer,
    ArchiveDocument document,
    string path,
    IProgress<double>? progress = null) : ICommand
{
    public async Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        IProgress<double>? serializeProgress = progress is not null
            ? new Progress<double>(p => progress.Report(p * 0.5))
            : null;
        byte[] data = await writer.SerializeAsync(document, serializeProgress, ct).ConfigureAwait(false);
        await WriteFileWithProgressAsync(path, data, progress, ct).ConfigureAwait(false);
        ArchiveDocument clone = document.DeepCopy();
        EntryFinder.ClearModifiedFlags(clone.RootEntry);
        return clone with { FilePath = path, OriginalFilePath = path, IsModified = false };
    }
    private static async Task WriteFileWithProgressAsync(string path, byte[] data, IProgress<double>? progress, CancellationToken ct)
    {
        const int BufferSize = 1024 * 1024;
        long totalBytes = data.Length;
        long bytesWritten = 0;
        using FileStream fileStream = new(path, FileMode.Create, FileAccess.Write, FileShare.None, BufferSize, useAsync: true);
        while (bytesWritten < totalBytes)
        {
            ct.ThrowIfCancellationRequested();
            int bytesToWrite = (int)Math.Min(BufferSize, totalBytes - bytesWritten);
            await fileStream.WriteAsync(data, (int)bytesWritten, bytesToWrite, ct).ConfigureAwait(false);
            bytesWritten += bytesToWrite;
            double progressValue = 0.5 + ((double)bytesWritten / totalBytes * 0.5);
            progress?.Report(progressValue);
        }
        await fileStream.FlushAsync(ct).ConfigureAwait(false);
    }
}
public sealed class ImportToNestedCommand(
    ArchiveDocument document,
    long targetId,
    string folderPath,
    FolderImportService importer) : ICommand
{
    public async Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        ArchiveDocument clone = document.DeepCopy();
        ArchiveEntry entry = EntryFinder.FindById(clone.RootEntry, targetId)
            ?? throw new InvalidOperationException("Entry not found.");
        byte[]? unknownBytes = await TryGetExistingUnknownBytesAsync(entry, ct).ConfigureAwait(false)
            ?? await ReadMsndUnknownBytesFromFolderAsync(folderPath, ct).ConfigureAwait(false);
        ImportResult result = await importer.AnalyzeFolderAsync(folderPath, ct).ConfigureAwait(false);
        Dictionary<string, byte[]> chunks = new(StringComparer.OrdinalIgnoreCase);
        foreach (ArchiveEntry child in result.Entries)
        {
            string ext = Path.GetExtension(child.Name).ToLowerInvariant();
            chunks[ext] = await child.GetDataAsync(ct).ConfigureAwait(false);
        }
        byte[] newMsndData = Formats.BuildMsnd(chunks, unknownBytes);
        entry.DataSource = new BufferSource(newMsndData);
        entry.Size = newMsndData.Length;
        entry.NestedType = ArchiveType.MSND;
        entry.IsModified = true;
        EntryFinder.PopulateMsndChildrenWithModified(newMsndData, entry);
        EntryFinder.MarkParentsModified(clone.RootEntry, targetId);
        return clone with { IsModified = true };
    }
    private static async Task<byte[]?> TryGetExistingUnknownBytesAsync(ArchiveEntry entry, CancellationToken ct)
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
    private static async Task<byte[]?> ReadMsndUnknownBytesFromFolderAsync(string folder, CancellationToken ct)
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
}
public sealed class AddBlankEntryCommand(
    ArchiveDocument document,
    long? parentId,
    string name,
    bool isContainer,
    ArchiveType? containerType) : ICommand
{
    public Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        ArchiveDocument clone = document.DeepCopy();
        byte[] entryData = isContainer && containerType == ArchiveType.MSND ? Formats.BuildEmptyMsnd() : [];
        ArchiveEntry newEntry = new()
        {
            Name = name,
            Size = entryData.Length,
            IsModified = true,
            ImportOrder = int.MaxValue,
            DataSource = new BufferSource(entryData),
            NestedType = isContainer ? containerType : null
        };
        if (isContainer && containerType == ArchiveType.MSND)
        {
            string baseName = Path.GetFileNameWithoutExtension(name);
            int sseqOff = Formats.MsndHeader;
            newEntry.Children.Add(new ArchiveEntry { Name = $"{baseName}.sseq", Size = 0, Offset = sseqOff, DataSource = new BufferSource([]), IsModified = true });
            newEntry.Children.Add(new ArchiveEntry { Name = $"{baseName}.sbnk", Size = 0, Offset = sseqOff, DataSource = new BufferSource([]), IsModified = true });
            newEntry.Children.Add(new ArchiveEntry { Name = $"{baseName}.swar", Size = 0, Offset = sseqOff, DataSource = new BufferSource([]), IsModified = true });
        }
        if (parentId is null)
        {
            clone.RootEntry.Children.Add(newEntry);
        }
        else
        {
            ArchiveEntry? targetParent = EntryFinder.FindById(clone.RootEntry, parentId.Value);
            if (targetParent is not null)
            {
                targetParent.Children.Add(newEntry);
                targetParent.IsModified = true;
                EntryFinder.MarkParentsModified(clone.RootEntry, parentId.Value);
            }
        }
        return Task.FromResult(clone with { IsModified = true });
    }
}
public sealed class DeleteEntryCommand(ArchiveDocument document, long targetId) : ICommand
{
    public Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        ArchiveDocument clone = document.DeepCopy();
        ArchiveEntry? parent = EntryFinder.FindParent(clone.RootEntry, targetId);
        bool removed = EntryFinder.RemoveById(clone.RootEntry, targetId);
        if (removed && parent is not null && parent != clone.RootEntry)
        {
            parent.IsModified = true;
            EntryFinder.MarkParentsModified(clone.RootEntry, parent.Id);
        }
        return Task.FromResult(clone with { IsModified = removed || clone.IsModified });
    }
}
public sealed class SortEntriesCommand(
    ArchiveDocument document,
    long? containerId = null,
    bool sortAlphabetically = true,
    bool recursive = false) : ICommand
{
    public Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        ArchiveDocument clone = document.DeepCopy();
        ArchiveEntry container = containerId is null
            ? clone.RootEntry
            : EntryFinder.FindById(clone.RootEntry, containerId.Value) ?? clone.RootEntry;
        if (container.NestedType == ArchiveType.MSND)
        {
            return Task.FromResult(clone);
        }
        if (recursive)
        {
            SortRecursive(container, sortAlphabetically);
        }
        else
        {
            SortContainer(container, sortAlphabetically);
        }
        if (container != clone.RootEntry)
        {
            container.IsModified = true;
            EntryFinder.MarkParentsModified(clone.RootEntry, container.Id);
        }
        return Task.FromResult(clone with { IsModified = true });
    }
    private static void SortRecursive(ArchiveEntry entry, bool sortAlphabetically)
    {
        foreach (ArchiveEntry child in entry.Children)
        {
            if (child.NestedType == ArchiveType.DSARC)
            {
                SortRecursive(child, sortAlphabetically);
            }
        }
        SortContainer(entry, sortAlphabetically);
    }
    private static void SortContainer(ArchiveEntry container, bool sortAlphabetically)
    {
        if (container.NestedType == ArchiveType.MSND)
        {
            return;
        }
        List<ArchiveEntry> sorted = sortAlphabetically
            ? [.. container.Children.OrderBy(e => e.Name, StringComparer.OrdinalIgnoreCase)]
            : [.. container.Children.OrderBy(e => e.ImportOrder)];
        container.Children.Clear();
        container.Children.AddRange(sorted);
    }
}
public sealed class RenameEntryCommand(ArchiveDocument document, long targetId, string newName) : ICommand
{
    public Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        ArchiveDocument clone = document.DeepCopy();
        ArchiveEntry? entry = EntryFinder.FindById(clone.RootEntry, targetId);
        if (entry is not null)
        {
            string newBaseName = Path.GetFileNameWithoutExtension(newName);
            entry.Name = newName;
            entry.IsModified = true;
            if (entry.NestedType == ArchiveType.MSND && entry.Children.Count == 3)
            {
                foreach (ArchiveEntry child in entry.Children)
                {
                    string ext = Path.GetExtension(child.Name);
                    child.Name = $"{newBaseName}{ext}";
                }
            }
            EntryFinder.MarkParentsModified(clone.RootEntry, targetId);
        }
        return Task.FromResult(clone with { IsModified = true });
    }
}
public sealed class MoveEntryCommand(ArchiveDocument document, long targetId, long? destinationParentId, int newIndex) : ICommand
{
    public Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        ArchiveDocument clone = document.DeepCopy();
        ArchiveEntry? sourceParent = EntryFinder.FindParent(clone.RootEntry, targetId);
        if (sourceParent is null)
        {
            return Task.FromResult(clone);
        }
        ArchiveEntry? entryToMove = null;
        int sourceIndex = -1;
        for (int i = 0; i < sourceParent.Children.Count; i++)
        {
            if (sourceParent.Children[i].Id == targetId)
            {
                entryToMove = sourceParent.Children[i];
                sourceIndex = i;
                sourceParent.Children.RemoveAt(i);
                break;
            }
        }
        if (entryToMove is null)
        {
            return Task.FromResult(clone);
        }
        ArchiveEntry destParent = destinationParentId is null
            ? clone.RootEntry
            : EntryFinder.FindById(clone.RootEntry, destinationParentId.Value) ?? clone.RootEntry;
        int insertIndex = newIndex;
        if (sourceParent.Id == destParent.Id && sourceIndex < newIndex)
        {
            insertIndex--;
        }
        insertIndex = Math.Clamp(insertIndex, 0, destParent.Children.Count);
        destParent.Children.Insert(insertIndex, entryToMove);
        if (sourceParent != clone.RootEntry)
        {
            sourceParent.IsModified = true;
        }
        if (destParent != clone.RootEntry && destParent != sourceParent)
        {
            destParent.IsModified = true;
        }
        return Task.FromResult(clone with { IsModified = true });
    }
}
public sealed class MoveEntryToContainerCommand(ArchiveDocument document, long targetId, long containerId) : ICommand
{
    public Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        ArchiveDocument clone = document.DeepCopy();
        ArchiveEntry? sourceParent = EntryFinder.FindParent(clone.RootEntry, targetId);
        if (sourceParent is null)
        {
            return Task.FromResult(clone);
        }
        ArchiveEntry? entryToMove = null;
        for (int i = 0; i < sourceParent.Children.Count; i++)
        {
            if (sourceParent.Children[i].Id == targetId)
            {
                entryToMove = sourceParent.Children[i];
                sourceParent.Children.RemoveAt(i);
                break;
            }
        }
        if (entryToMove is null)
        {
            return Task.FromResult(clone);
        }
        ArchiveEntry? targetContainer = EntryFinder.FindById(clone.RootEntry, containerId);
        if (targetContainer?.NestedType == ArchiveType.DSARC)
        {
            targetContainer.Children.Add(entryToMove);
            targetContainer.IsModified = true;
            if (sourceParent != clone.RootEntry)
            {
                sourceParent.IsModified = true;
            }
            EntryFinder.MarkParentsModified(clone.RootEntry, containerId);
        }
        else
        {
            sourceParent.Children.Add(entryToMove);
            return Task.FromResult(clone);
        }
        return Task.FromResult(clone with { IsModified = true });
    }
}
public sealed class MoveEntryUpCommand(ArchiveDocument document, long targetId) : ICommand
{
    public Task<ArchiveDocument> ExecuteAsync(CancellationToken ct = default)
    {
        ArchiveDocument clone = document.DeepCopy();
        ArchiveEntry? sourceParent = EntryFinder.FindParent(clone.RootEntry, targetId);
        if (sourceParent is null || sourceParent == clone.RootEntry)
        {
            return Task.FromResult(clone);
        }
        ArchiveEntry? grandParent = EntryFinder.FindParent(clone.RootEntry, sourceParent.Id);
        if (grandParent is null)
        {
            return Task.FromResult(clone);
        }
        ArchiveEntry? entryToMove = null;
        for (int i = 0; i < sourceParent.Children.Count; i++)
        {
            if (sourceParent.Children[i].Id == targetId)
            {
                entryToMove = sourceParent.Children[i];
                sourceParent.Children.RemoveAt(i);
                break;
            }
        }
        if (entryToMove is null)
        {
            return Task.FromResult(clone);
        }
        int parentIndex = grandParent.Children.FindIndex(c => c.Id == sourceParent.Id);
        grandParent.Children.Insert(parentIndex + 1, entryToMove);
        sourceParent.IsModified = true;
        if (grandParent != clone.RootEntry)
        {
            grandParent.IsModified = true;
        }
        return Task.FromResult(clone with { IsModified = true });
    }
}
public sealed class CommandProcessor
{
    private readonly LinkedList<ArchiveDocument> _undoStack = new();
    private readonly Stack<ArchiveDocument> _redoStack = new();
    private const int MaxUndoLevels = 50;
    public ArchiveDocument Current { get; private set; } = new();
    public bool CanUndo => _undoStack.Count > 0;
    public bool CanRedo => _redoStack.Count > 0;
    public async Task<ArchiveDocument> ExecuteAsync(ICommand command, CancellationToken ct = default)
    {
        _ = _undoStack.AddFirst(Current.DeepCopy());
        while (_undoStack.Count > MaxUndoLevels)
        {
            _undoStack.RemoveLast();
        }
        _redoStack.Clear();
        Current = await command.ExecuteAsync(ct).ConfigureAwait(false);
        return Current;
    }
    public ArchiveDocument Undo()
    {
        if (!CanUndo)
        {
            throw new InvalidOperationException("Nothing to undo.");
        }
        _redoStack.Push(Current);
        Current = _undoStack.First!.Value;
        _undoStack.RemoveFirst();
        return Current;
    }
    public ArchiveDocument Redo()
    {
        if (!CanRedo)
        {
            throw new InvalidOperationException("Nothing to redo.");
        }
        _ = _undoStack.AddFirst(Current);
        Current = _redoStack.Pop();
        return Current;
    }
    public void SetCurrent(ArchiveDocument doc)
    {
        _undoStack.Clear();
        _redoStack.Clear();
        Current = doc;
    }
    public void Clear()
    {
        _undoStack.Clear();
        _redoStack.Clear();
        Current = new ArchiveDocument();
    }
}