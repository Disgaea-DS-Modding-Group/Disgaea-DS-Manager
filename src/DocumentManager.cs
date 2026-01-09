using Disgaea_DS_Manager.Commands;
using Disgaea_DS_Manager.Models;
namespace Disgaea_DS_Manager.Services;

public sealed class DocumentManager
{
    private readonly ArchiveReaderService _reader = new();
    private readonly ArchiveWriterService _writer = new();
    private readonly FolderImportService _importer = new();
    private readonly ExtractionService _extractor = new();
    private readonly CommandProcessor _processor = new();
    public ArchiveDocument Current => _processor.Current;
    public bool CanUndo => _processor.CanUndo;
    public bool CanRedo => _processor.CanRedo;
    public event Action? DocumentChanged;
    public async Task<ArchiveDocument> LoadArchiveAsync(string path, CancellationToken ct = default)
    {
        LoadArchiveCommand command = new(_reader, path);
        ArchiveDocument doc = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
        return doc;
    }
    public async Task<ArchiveDocument> ImportFolderAsync(string folderPath, CancellationToken ct = default)
    {
        ImportFolderCommand command = new(_importer, folderPath);
        ArchiveDocument doc = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
        return doc;
    }
    public async Task<ArchiveDocument> ReplaceEntryAsync(ArchiveEntry target, byte[] newData, CancellationToken ct = default)
    {
        ReplaceEntryCommand command = new(Current, target.Id, newData);
        ArchiveDocument doc = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
        return doc;
    }
    public async Task<ArchiveDocument> ReplaceChunkAsync(ArchiveEntry parentEntry, ArchiveEntry chunkEntry, byte[] newChunkData, CancellationToken ct = default)
    {
        ReplaceChunkCommand command = new(Current, parentEntry.Id, chunkEntry.Id, newChunkData, _writer);
        ArchiveDocument doc = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
        return doc;
    }
    public async Task<ArchiveDocument> ImportToNestedAsync(ArchiveEntry target, string folderPath, CancellationToken ct = default)
    {
        ImportToNestedCommand command = new(Current, target.Id, folderPath, _importer);
        ArchiveDocument doc = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
        return doc;
    }
    public async Task<ArchiveDocument> SaveArchiveAsync(string? path = null, IProgress<double>? progress = null, CancellationToken ct = default)
    {
        path ??= Current.FilePath ?? throw new InvalidOperationException("No save path specified.");
        SaveArchiveCommand command = new(_writer, Current, path, progress);
        ArchiveDocument doc = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
        return doc;
    }
    public async Task ExtractAllAsync(string destFolder, bool nested, IProgress<(int Current, int Total)>? progress = null, CancellationToken ct = default)
    {
        await _extractor.ExtractAllAsync(Current, destFolder, nested, progress, ct).ConfigureAwait(false);
    }
    public async Task ExtractSingleAsync(ArchiveEntry entry, string destFolder, CancellationToken ct = default)
    {
        await _extractor.ExtractSingleAsync(entry, destFolder, ct).ConfigureAwait(false);
    }
    public async Task ExtractNestedAsync(ArchiveEntry entry, string destFolder, IProgress<(int Current, int Total)>? progress = null, CancellationToken ct = default)
    {
        await _extractor.ExtractNestedAsync(entry, destFolder, progress, ct).ConfigureAwait(false);
    }
    public ArchiveDocument CreateNew(string path)
    {
        ArchiveType type = path.EndsWith(".msnd", StringComparison.OrdinalIgnoreCase) ? ArchiveType.MSND : ArchiveType.DSARC;
        ArchiveDocument doc = new()
        {
            FilePath = path,
            FileType = type,
            IsModified = true,
            HasContent = true,
            RootEntry = new ArchiveEntry { Name = Path.GetFileName(path) }
        };
        _processor.SetCurrent(doc);
        DocumentChanged?.Invoke();
        return doc;
    }
    public ArchiveDocument CreateNewEmpty(ArchiveType type)
    {
        string displayName = type == ArchiveType.DSARC ? "New DSARC" : "New DSEQ";
        ArchiveEntry rootEntry = new() { Name = displayName };
        if (type == ArchiveType.MSND)
        {
            rootEntry.NestedType = ArchiveType.MSND;
            string baseName = "new_sound";
            byte[] emptyMsnd = Formats.BuildEmptyMsnd();
            rootEntry.DataSource = new BufferSource(emptyMsnd);
            rootEntry.Size = emptyMsnd.Length;
            int sseqOff = Formats.MsndHeader;
            rootEntry.Children.Add(new ArchiveEntry { Name = $"{baseName}.sseq", Size = 0, Offset = sseqOff, DataSource = new BufferSource([]), IsModified = true });
            rootEntry.Children.Add(new ArchiveEntry { Name = $"{baseName}.sbnk", Size = 0, Offset = sseqOff, DataSource = new BufferSource([]), IsModified = true });
            rootEntry.Children.Add(new ArchiveEntry { Name = $"{baseName}.swar", Size = 0, Offset = sseqOff, DataSource = new BufferSource([]), IsModified = true });
        }
        ArchiveDocument doc = new()
        {
            FilePath = null,
            FileType = type,
            IsModified = type == ArchiveType.MSND,
            HasContent = true,
            RootEntry = rootEntry
        };
        _processor.SetCurrent(doc);
        DocumentChanged?.Invoke();
        return doc;
    }
    public async Task AddBlankEntryAsync(ArchiveEntry? parent, string name, bool isContainer, ArchiveType? containerType, CancellationToken ct = default)
    {
        AddBlankEntryCommand command = new(Current, parent?.Id, name, isContainer, containerType);
        _ = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
    }
    public async Task DeleteEntryAsync(ArchiveEntry target, CancellationToken ct = default)
    {
        DeleteEntryCommand command = new(Current, target.Id);
        _ = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
    }
    public async Task SortAlphabeticallyAsync(ArchiveEntry? container = null, CancellationToken ct = default)
    {
        if (container?.NestedType == ArchiveType.MSND)
        {
            return;
        }
        SortEntriesCommand command = new(Current, container?.Id, sortAlphabetically: true);
        _ = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
    }
    public async Task SortAlphabeticallyRecursiveAsync(ArchiveEntry? container = null, CancellationToken ct = default)
    {
        if (container?.NestedType == ArchiveType.MSND)
        {
            return;
        }
        SortEntriesCommand command = new(Current, container?.Id, sortAlphabetically: true, recursive: true);
        _ = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
    }
    public async Task SortByImportOrderAsync(ArchiveEntry? container = null, CancellationToken ct = default)
    {
        if (container?.NestedType == ArchiveType.MSND)
        {
            return;
        }
        SortEntriesCommand command = new(Current, container?.Id, sortAlphabetically: false);
        _ = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
    }
    public bool IsRootEmpty()
    {
        return Current.FileType == ArchiveType.DSARC
            ? Current.RootEntry.Children.Count == 0
            : Current.FileType == ArchiveType.MSND && Current.RootEntry.Children.All(c => c.Size == 0);
    }
    public bool CanSortContainer(ArchiveEntry? container)
    {
        return container is null
            ? Current.FileType == ArchiveType.DSARC && Current.RootEntry.Children.Count > 1
            : container.NestedType == ArchiveType.DSARC && container.Children.Count > 1;
    }
    public bool HasNestedArchives()
    {
        return Current.FileType == ArchiveType.DSARC &&
               Current.RootEntry.Children.Any(c => c.NestedType == ArchiveType.DSARC && c.Children.Count > 0);
    }
    public bool HasBlankFiles()
    {
        return HasBlankEntries(Current.RootEntry);
    }
    // <summary>
    // Checks if an entry or any of its descendants contain blank/empty content.
    // </summary>
    private static bool HasBlankEntries(ArchiveEntry entry)
    {
        foreach (ArchiveEntry child in entry.Children)
        {
            if (IsFullyEmpty(child))
            {
                return true;
            }
            if (child.NestedType == ArchiveType.DSARC && HasBlankEntries(child))
            {
                return true;
            }
        }
        return false;
    }
    // <summary>
    // Determines if an entry is "fully empty" - meaning it has no meaningful content.
    // For containers, this means all descendants are also fully empty.
    // </summary>
    private static bool IsFullyEmpty(ArchiveEntry entry)
    {
        return entry.NestedType == ArchiveType.MSND
            ? entry.Children.Count == 0 || entry.Children.All(c => c.Size == 0)
            : entry.NestedType == ArchiveType.DSARC ? entry.Children.Count == 0 || entry.Children.All(IsFullyEmpty) : entry.Size == 0;
    }
    // <summary>
    // Removes all empty entries from the archive.
    // Prioritizes deleting containers that are fully empty (all descendants empty)
    // before deleting individual empty files. This is more efficient and matches
    // user expectations - if a whole folder is empty, delete the folder.
    // </summary>
    public async Task RemoveAllEmptyEntriesAsync(CancellationToken ct = default)
    {
        bool changed = true;
        while (changed)
        {
            ct.ThrowIfCancellationRequested();
            changed = false;
            ArchiveEntry? toDelete = FindBestDeletionTarget(Current.RootEntry);
            if (toDelete is not null)
            {
                await DeleteEntryAsync(toDelete, ct).ConfigureAwait(false);
                changed = true;
            }
        }
    }
    // <summary>
    // Finds the best entry to delete next. Prioritizes:
    // 1. Fully empty containers (delete whole container instead of individual children)
    // 2. Individual empty files in containers that have mixed content
    // </summary>
    private static ArchiveEntry? FindBestDeletionTarget(ArchiveEntry root)
    {
        foreach (ArchiveEntry child in root.Children)
        {
            if (IsFullyEmpty(child))
            {
                return child;
            }
            if (child.NestedType == ArchiveType.DSARC)
            {
                ArchiveEntry? nested = FindBestDeletionTarget(child);
                if (nested is not null)
                {
                    return nested;
                }
            }
        }
        return null;
    }
    public async Task RemoveBlankEntriesRecursiveAsync(CancellationToken ct = default)
    {
        await RemoveAllEmptyEntriesAsync(ct).ConfigureAwait(false);
    }
    public async Task RemoveEmptyDsarcEntriesAsync(CancellationToken ct = default)
    {
        await Task.CompletedTask;
    }
    public async Task RenameEntryAsync(ArchiveEntry target, string newName, CancellationToken ct = default)
    {
        RenameEntryCommand command = new(Current, target.Id, newName);
        _ = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
    }
    public async Task MoveEntryAsync(ArchiveEntry entry, long? destinationParentId, int newIndex, CancellationToken ct = default)
    {
        if (Current.FileType == ArchiveType.MSND)
        {
            return;
        }
        MoveEntryCommand command = new(Current, entry.Id, destinationParentId, newIndex);
        _ = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
    }
    public async Task MoveEntryToContainerAsync(ArchiveEntry entry, ArchiveEntry targetContainer, CancellationToken ct = default)
    {
        if (targetContainer.NestedType != ArchiveType.DSARC)
        {
            return;
        }
        MoveEntryToContainerCommand command = new(Current, entry.Id, targetContainer.Id);
        _ = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
    }
    public async Task MoveEntryUpAsync(ArchiveEntry entry, CancellationToken ct = default)
    {
        if (Current.FileType == ArchiveType.MSND)
        {
            return;
        }
        MoveEntryUpCommand command = new(Current, entry.Id);
        _ = await _processor.ExecuteAsync(command, ct).ConfigureAwait(false);
        DocumentChanged?.Invoke();
    }
    public bool HasDuplicateName(ArchiveEntry entry, string newName)
    {
        ArchiveEntry? parent = EntryFinder.FindParent(Current.RootEntry, entry.Id);
        return parent is not null && parent.Children.Any(c => c.Id != entry.Id && c.Name.Equals(newName, StringComparison.OrdinalIgnoreCase));
    }
    public List<ArchiveEntry> GetEmptyDsarcEntries()
    {
        List<ArchiveEntry> emptyEntries = [];
        FindEmptyContainers(Current.RootEntry, emptyEntries);
        return emptyEntries;
    }
    private static void FindEmptyContainers(ArchiveEntry entry, List<ArchiveEntry> results)
    {
        foreach (ArchiveEntry child in entry.Children)
        {
            if (IsFullyEmpty(child))
            {
                results.Add(child);
            }
            else if (child.NestedType == ArchiveType.DSARC)
            {
                FindEmptyContainers(child, results);
            }
        }
    }
    public ArchiveDocument Undo()
    {
        ArchiveDocument doc = _processor.Undo();
        DocumentChanged?.Invoke();
        return doc;
    }
    public ArchiveDocument Redo()
    {
        ArchiveDocument doc = _processor.Redo();
        DocumentChanged?.Invoke();
        return doc;
    }
    public void Clear()
    {
        _processor.Clear();
        DocumentChanged?.Invoke();
    }
}