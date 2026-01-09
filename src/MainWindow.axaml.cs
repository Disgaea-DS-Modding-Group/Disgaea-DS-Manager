using Avalonia;
using Avalonia.Collections;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Disgaea_DS_Manager.Commands;
using Disgaea_DS_Manager.Models;
using Disgaea_DS_Manager.Services;
namespace Disgaea_DS_Manager;

public partial class MainWindow : Window
{
    private readonly DocumentManager _manager = new();
    private CancellationTokenSource? _cts;
    private TreeViewItem? _rootItem;
    private MenuItem _undoMenu = null!;
    private MenuItem _redoMenu = null!;
    private MenuItem _sortMenu = null!;
    private MenuItem _sortRecursiveMenu = null!;
    private TreeView _treeView = null!;
    private TextBlock _statusLabel = null!;
    private TextBlock _modifiedLabel = null!;
    private ProgressBar _progressBar = null!;
    private TextBox _logTextBox = null!;
    private Canvas _dragOverlay = null!;
    private bool _isDragPending;
    private bool _isDragging;
    private Point _dragStartPos;
    private TreeViewItem? _dragSourceItem;
    private ArchiveEntry? _dragSourceEntry;
    private Border? _dragGhost;
    private Rectangle? _dropIndicator;
    private Border? _dropHighlight;
    private DropTarget? _currentDropTarget;
    private readonly Dictionary<long, bool> _expansionState = [];
    private const double DragThreshold = 6.0;
    private const int MaxLogLines = 500;
    private static readonly IBrush ModifiedBrush = new SolidColorBrush(Color.Parse("#FF8800"));
    private static readonly IBrush DropIndicatorBrush = new SolidColorBrush(Color.Parse("#0078D4"));
    private static readonly IBrush DropHighlightBrush = new SolidColorBrush(Color.Parse("#200078D4"));
    private static readonly IBrush DropHighlightBorder = new SolidColorBrush(Color.Parse("#0078D4"));
    private static readonly IBrush GhostBackground = new SolidColorBrush(Color.Parse("#E8F0FE"));
    private static readonly IBrush GhostBorder = new SolidColorBrush(Color.Parse("#0078D4"));
    private record DropTarget(ArchiveEntry Entry, TreeViewItem Item, DropPosition Position, ArchiveEntry? Parent);
    private enum DropPosition { Before, After, Inside }
    public MainWindow()
    {
        InitializeComponent();
        if (!Design.IsDesignMode)
        {
            InitializeControls();
            WireEvents();
            _manager.DocumentChanged += RefreshTree;
            UpdateEditMenu();
        }
    }
    private void InitializeControls()
    {
        _undoMenu = this.FindControl<MenuItem>("UndoMenu")!;
        _redoMenu = this.FindControl<MenuItem>("RedoMenu")!;
        _sortMenu = this.FindControl<MenuItem>("SortMenu")!;
        _sortRecursiveMenu = this.FindControl<MenuItem>("SortRecursiveMenu")!;
        _treeView = this.FindControl<TreeView>("TreeView")!;
        _statusLabel = this.FindControl<TextBlock>("StatusLabel")!;
        _modifiedLabel = this.FindControl<TextBlock>("ModifiedLabel")!;
        _progressBar = this.FindControl<ProgressBar>("ProgressBar")!;
        _logTextBox = this.FindControl<TextBox>("LogTextBox")!;
        _dragOverlay = this.FindControl<Canvas>("DragOverlay")!;
    }
    private void WireEvents()
    {
        this.FindControl<MenuItem>("NewMenu")!.Click += async (_, _) => await NewArchiveAsync();
        this.FindControl<MenuItem>("OpenMenu")!.Click += async (_, _) => await OpenArchiveAsync();
        this.FindControl<MenuItem>("SaveMenu")!.Click += async (_, _) => await SaveAsync();
        this.FindControl<MenuItem>("SaveAsMenu")!.Click += async (_, _) => await SaveAsAsync();
        this.FindControl<MenuItem>("ExitMenu")!.Click += (_, _) => Close();
        _undoMenu.Click += (_, _) => { if (_manager.CanUndo) { _ = _manager.Undo(); Log("Undo"); } };
        _redoMenu.Click += (_, _) => { if (_manager.CanRedo) { _ = _manager.Redo(); Log("Redo"); } };
        _sortMenu.Click += async (_, _) => { await _manager.SortAlphabeticallyAsync(); Log("Sorted"); };
        _sortRecursiveMenu.Click += async (_, _) => { await _manager.SortAlphabeticallyRecursiveAsync(); Log("Sorted Recursively"); };
        this.FindControl<MenuItem>("InfoMenu")!.Click += async (_, _) => await DialogService.ShowInfoAsync(this);
        _treeView.SelectionChanged += (_, _) => UpdateEditMenu();
        _treeView.AddHandler(PointerPressedEvent, OnPointerPressed, RoutingStrategies.Tunnel);
        _treeView.AddHandler(PointerMovedEvent, OnPointerMoved, RoutingStrategies.Tunnel);
        _treeView.AddHandler(PointerReleasedEvent, OnPointerReleased, RoutingStrategies.Tunnel);
    }
    private void OnPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (!e.GetCurrentPoint(_treeView).Properties.IsLeftButtonPressed)
        {
            return;
        }
        TreeViewItem? item = (e.Source as Visual)?.FindAncestorOfType<TreeViewItem>();
        if (item is null || item == _rootItem)
        {
            return;
        }
        if (item.DataContext is not ArchiveEntry entry)
        {
            return;
        }
        if (_manager.Current.FileType != ArchiveType.DSARC)
        {
            return;
        }
        if (item.Parent is TreeViewItem { DataContext: ArchiveEntry { NestedType: ArchiveType.MSND } })
        {
            return;
        }
        _dragStartPos = e.GetPosition(this);
        _isDragPending = true;
        _dragSourceItem = item;
        _dragSourceEntry = entry;
    }
    private void OnPointerMoved(object? sender, PointerEventArgs e)
    {
        Point pos = e.GetPosition(this);
        if (_isDragPending && !_isDragging)
        {
            Point delta = pos - _dragStartPos;
            if (Math.Abs(delta.X) > DragThreshold || Math.Abs(delta.Y) > DragThreshold)
            {
                StartDrag();
            }
            return;
        }
        if (_isDragging)
        {
            UpdateDrag(pos, e.GetPosition(_treeView));
        }
    }
    private void OnPointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        if (e.InitialPressMouseButton == MouseButton.Right && !_isDragging)
        {
            TreeViewItem? item = (e.Source as Visual)?.FindAncestorOfType<TreeViewItem>();
            if (item is not null)
            {
                _treeView.SelectedItem = item;
                ShowContextMenu(item);
                e.Handled = true;
            }
        }
        if (_isDragging) { FinishDrag(); e.Handled = true; }
        CancelDrag();
    }
    private void StartDrag()
    {
        if (_dragSourceEntry is null)
        {
            return;
        }
        _isDragging = true;
        _isDragPending = false;
        _dragGhost = new Border
        {
            Background = GhostBackground,
            BorderBrush = GhostBorder,
            BorderThickness = new Thickness(1.5),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(10, 6),
            Opacity = 0.95,
            Child = new TextBlock { Text = _dragSourceEntry.Name, FontWeight = FontWeight.Medium }
        };
        _dragOverlay.Children.Add(_dragGhost);
        _dropIndicator = new Rectangle { Height = 2, Fill = DropIndicatorBrush, IsVisible = false };
        _dragOverlay.Children.Add(_dropIndicator);
        _dropHighlight = new Border { Background = DropHighlightBrush, BorderBrush = DropHighlightBorder, BorderThickness = new Thickness(2), CornerRadius = new CornerRadius(3), IsVisible = false };
        _dragOverlay.Children.Insert(0, _dropHighlight);
        _ = _dragSourceItem?.Opacity = 0.4;
    }
    private void UpdateDrag(Point windowPos, Point treePos)
    {
        if (_dragGhost is not null)
        {
            Canvas.SetLeft(_dragGhost, windowPos.X + 12);
            Canvas.SetTop(_dragGhost, windowPos.Y + 12);
        }
        DropTarget? target = FindDropTarget(treePos);
        _currentDropTarget = target;
        _ = _dropIndicator?.IsVisible = false;
        _ = _dropHighlight?.IsVisible = false;
        if (target is null || target.Entry == _dragSourceEntry)
        {
            return;
        }
        Rect? itemBounds = GetItemBoundsInWindow(target.Item);
        if (!itemBounds.HasValue)
        {
            return;
        }
        if (target.Position == DropPosition.Inside)
        {
            if (_dropHighlight is not null)
            {
                _dropHighlight.IsVisible = true;
                _dropHighlight.Width = itemBounds.Value.Width;
                _dropHighlight.Height = itemBounds.Value.Height;
                Canvas.SetLeft(_dropHighlight, itemBounds.Value.X);
                Canvas.SetTop(_dropHighlight, itemBounds.Value.Y);
            }
        }
        else if (_dropIndicator is not null)
        {
            _dropIndicator.IsVisible = true;
            _dropIndicator.Width = itemBounds.Value.Width;
            _dropIndicator.Height = 2;
            double indicatorY = target.Position == DropPosition.Before ? itemBounds.Value.Y - 1 : itemBounds.Value.Bottom - 1;
            Canvas.SetLeft(_dropIndicator, itemBounds.Value.X);
            Canvas.SetTop(_dropIndicator, indicatorY);
        }
    }
    private Rect? GetItemBoundsInWindow(TreeViewItem item)
    {
        try
        {
            Point? topLeft = item.TranslatePoint(new Point(0, 0), this);
            return topLeft is null ? null : new Rect(topLeft.Value, item.Bounds.Size);
        }
        catch { return null; }
    }
    private DropTarget? FindDropTarget(Point treePos)
    {
        if (_rootItem is null)
        {
            return null;
        }
        foreach ((TreeViewItem item, ArchiveEntry? parent, bool isInsideMsnd) in GetVisibleTreeItemsWithContext(_rootItem))
        {
            if (item.DataContext is not ArchiveEntry entry)
            {
                continue;
            }
            if (entry == _dragSourceEntry)
            {
                continue;
            }
            if (isInsideMsnd)
            {
                continue;
            }
            Point? itemPos = item.TranslatePoint(new Point(0, 0), _treeView);
            if (itemPos is null)
            {
                continue;
            }
            Rect bounds = new(itemPos.Value, item.Bounds.Size);
            if (!bounds.Contains(treePos))
            {
                continue;
            }
            double relY = treePos.Y - bounds.Y;
            if (entry.NestedType == ArchiveType.DSARC)
            {
                bool isExpanded = item.IsExpanded;
                bool hasChildren = entry.Children.Count > 0;
                if (isExpanded && hasChildren)
                {
                    return relY < bounds.Height / 2
                        ? new DropTarget(entry, item, DropPosition.Before, parent)
                        : new DropTarget(entry, item, DropPosition.After, parent);
                }
                double quarter = bounds.Height / 4;
                return relY < quarter
                    ? new DropTarget(entry, item, DropPosition.Before, parent)
                    : relY > bounds.Height - quarter
                    ? new DropTarget(entry, item, DropPosition.After, parent)
                    : new DropTarget(entry, item, DropPosition.Inside, parent);
            }
            return new DropTarget(entry, item, relY < bounds.Height / 2 ? DropPosition.Before : DropPosition.After, parent);
        }
        return null;
    }
    private static IEnumerable<(TreeViewItem Item, ArchiveEntry? Parent, bool IsInsideMsnd)> GetVisibleTreeItemsWithContext(TreeViewItem parent, ArchiveEntry? parentEntry = null, bool parentIsMsnd = false)
    {
        if (parent.ItemsSource is not IEnumerable<TreeViewItem> children)
        {
            yield break;
        }
        foreach (TreeViewItem child in children)
        {
            ArchiveEntry? childEntry = child.DataContext as ArchiveEntry;
            bool isMsnd = childEntry?.NestedType == ArchiveType.MSND;
            if (child.IsExpanded)
            {
                foreach ((TreeViewItem Item, ArchiveEntry? Parent, bool IsInsideMsnd) grandchild in GetVisibleTreeItemsWithContext(child, childEntry, isMsnd || parentIsMsnd))
                {
                    yield return grandchild;
                }
            }
            yield return (child, parentEntry, parentIsMsnd);
        }
    }
    private async void FinishDrag()
    {
        if (_dragSourceEntry is null || _currentDropTarget is null)
        {
            return;
        }
        if (_currentDropTarget.Entry == _dragSourceEntry)
        {
            return;
        }
        DropTarget target = _currentDropTarget;
        if (target.Position == DropPosition.Inside && target.Entry.NestedType == ArchiveType.DSARC)
        {
            await _manager.MoveEntryToContainerAsync(_dragSourceEntry, target.Entry);
            Log($"Moved '{_dragSourceEntry.Name}' into '{target.Entry.Name}'");
        }
        else
        {
            ArchiveEntry? targetParent = target.Parent;
            long? destParentId = targetParent?.Id;
            List<ArchiveEntry> targetSiblings = GetTargetSiblings(targetParent);
            int targetIdx = targetSiblings.FindIndex(c => c.Id == target.Entry.Id);
            if (targetIdx < 0)
            {
                return;
            }
            int newIdx = target.Position == DropPosition.After ? targetIdx + 1 : targetIdx;
            ArchiveEntry? sourceParent = EntryFinder.FindParent(_manager.Current.RootEntry, _dragSourceEntry.Id);
            bool sameParent = (sourceParent?.Id ?? 0) == (targetParent?.Id ?? 0);
            if (sameParent)
            {
                int sourceIdx = targetSiblings.FindIndex(c => c.Id == _dragSourceEntry.Id);
                if (sourceIdx >= 0 && sourceIdx < newIdx)
                {
                    newIdx--;
                }
            }
            await _manager.MoveEntryAsync(_dragSourceEntry, destParentId, newIdx);
            Log($"Moved '{_dragSourceEntry.Name}'");
        }
    }
    private List<ArchiveEntry> GetTargetSiblings(ArchiveEntry? targetParent)
    {
        if (targetParent is null)
        {
            return _manager.Current.RootEntry.Children;
        }
        ArchiveEntry? parentInDoc = EntryFinder.FindById(_manager.Current.RootEntry, targetParent.Id);
        return parentInDoc?.Children ?? _manager.Current.RootEntry.Children;
    }
    private void CancelDrag()
    {
        _isDragPending = false;
        _isDragging = false;
        _currentDropTarget = null;
        _ = _dragSourceItem?.Opacity = 1.0;
        _dragOverlay.Children.Clear();
        _dragGhost = null;
        _dropIndicator = null;
        _dropHighlight = null;
        _dragSourceItem = null;
        _dragSourceEntry = null;
    }
    private void ShowContextMenu(TreeViewItem node)
    {
        AvaloniaList<Control> items = [];
        bool isRoot = node == _rootItem;
        ArchiveEntry? entry = node.DataContext as ArchiveEntry;
        ArchiveType docType = _manager.Current.FileType;
        if (isRoot)
        {
            items.Add(MenuItem("Import Folder", ImportFolderAsync));
            items.Add(MenuItem("Extract All", () => ExtractAllAsync(false)));
            if (docType == ArchiveType.DSARC)
            {
                items.Add(MenuItem("Extract All (Nested)", () => ExtractAllAsync(true)));
                items.Add(new Separator());
                items.Add(MenuItem("Add Blank File", () => AddBlankFileAsync(null)));
                items.Add(MenuItem("Add Blank DSARC", () => AddBlankContainerAsync(null, ArchiveType.DSARC)));
                items.Add(MenuItem("Add Blank DSEQ", () => AddBlankContainerAsync(null, ArchiveType.MSND)));
            }
        }
        else if (entry is not null)
        {
            bool isInMsnd = node.Parent is TreeViewItem { DataContext: ArchiveEntry { NestedType: ArchiveType.MSND } };
            if (entry.IsNested && entry.Children.Count > 0)
            {
                items.Add(MenuItem("Import Folder", () => ImportToNestedAsync(entry)));
                items.Add(MenuItem("Extract All", () => ExtractNestedAsync(entry)));
            }
            bool isTopLevel = node.Parent == _rootItem;
            if (isTopLevel || docType == ArchiveType.MSND)
            {
                items.Add(MenuItem("Extract", () => ExtractSingleAsync(entry)));
                items.Add(MenuItem("Replace", () => RenameEntryAsync(entry)));
                items.Add(MenuItem("Rename", () => RenameEntryAsync(entry)));
            }
            else if (node.Parent is TreeViewItem { DataContext: ArchiveEntry parent })
            {
                items.Add(MenuItem("Extract File", () => ExtractSingleAsync(entry)));
                if (parent.NestedType == ArchiveType.MSND)
                {
                    items.Add(MenuItem("Replace File", () => ReplaceChunkAsync(parent, entry)));
                }
                else
                {
                    items.Add(MenuItem("Replace", () => RenameEntryAsync(entry)));
                }
                if (!isInMsnd)
                {
                    items.Add(MenuItem("Rename", () => RenameEntryAsync(entry)));
                }
            }
            if (!isInMsnd)
            {
                if (items.Count > 0)
                {
                    items.Add(new Separator());
                }
                items.Add(MenuItem("Delete", () => DeleteEntryAsync(entry)));
            }
            if (entry.NestedType == ArchiveType.DSARC)
            {
                items.Add(new Separator());
                if (_manager.CanSortContainer(entry))
                {
                    items.Add(MenuItem("Sort Contents", () => SortContainerAsync(entry)));
                    items.Add(MenuItem("Sort Contents (Recursive)", () => SortRecursiveAsync(entry)));
                }
                items.Add(MenuItem("Add Blank File", () => AddBlankFileAsync(entry)));
                items.Add(MenuItem("Add Blank DSARC", () => AddBlankContainerAsync(entry, ArchiveType.DSARC)));
                items.Add(MenuItem("Add Blank DSEQ", () => AddBlankContainerAsync(entry, ArchiveType.MSND)));
            }
        }
        if (items.Count == 0)
        {
            return;
        }
        ContextMenu menu = new() { ItemsSource = items, PlacementTarget = node };
        menu.Open(node);
    }
    private async Task SortContainerAsync(ArchiveEntry container)
    {
        await _manager.SortAlphabeticallyAsync(container);
        Log($"Sorted contents of '{container.Name}'");
    }
    private async Task SortRecursiveAsync(ArchiveEntry? container)
    {
        await _manager.SortAlphabeticallyRecursiveAsync(container);
        Log($"Sorted contents of '{container?.Name ?? "Root"}' (Recursive)");
    }
    private static MenuItem MenuItem(string header, Func<Task> action)
    {
        MenuItem item = new() { Header = header };
        item.Click += async (_, _) => await action();
        return item;
    }
    private static FilePickerFileType ArchiveFileFilter => new("Disgaea Archives") { Patterns = ["*.dat", "*.msnd"] };
    private async Task<string?> PickOpenFileAsync(string title, bool archives = true)
    {
        IReadOnlyList<IStorageFile> result = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = title,
            AllowMultiple = false,
            FileTypeFilter = archives ? [ArchiveFileFilter, FilePickerFileTypes.All] : [FilePickerFileTypes.All]
        });
        return result.Count > 0 ? result[0].TryGetLocalPath() : null;
    }
    private async Task<string?> PickSaveFileAsync(string title, string? name = null)
    {
        IStorageFile? result = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = title,
            SuggestedFileName = name,
            FileTypeChoices = [ArchiveFileFilter, FilePickerFileTypes.All]
        });
        return result?.TryGetLocalPath();
    }
    private async Task<string?> PickFolderAsync(string title)
    {
        IReadOnlyList<IStorageFolder> result = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions { Title = title });
        return result.Count > 0 ? result[0].TryGetLocalPath() : null;
    }
    private async Task NewArchiveAsync()
    {
        ArchiveTypeChoice? choice = await DialogService.ShowNewFileDialogAsync(this);
        if (choice is null)
        {
            return;
        }
        ArchiveType type = choice == ArchiveTypeChoice.DSARC ? ArchiveType.DSARC : ArchiveType.MSND;
        if (type == ArchiveType.DSARC)
        {
            await DialogService.ShowMessageAsync(this, "New DSARC",
                "A new empty DSARC has been created.\n\n" +
                "Note: You must add at least one file before saving. " +
                "Use the right-click context menu to add files or nested archives.");
        }
        _ = _manager.CreateNewEmpty(type);
        Log($"Created new {(type == ArchiveType.DSARC ? "DSARC" : "DSEQ")}");
    }
    private async Task OpenArchiveAsync()
    {
        string? path = await PickOpenFileAsync("Open Archive");
        if (path is null)
        {
            return;
        }
        try
        {
            _expansionState.Clear();
            await RunAsync(ct => _manager.LoadArchiveAsync(path, ct), $"Opened: {System.IO.Path.GetFileName(path)}");
        }
        catch (Exception ex) { await DialogService.ShowErrorAsync(this, ex.Message); }
    }
    private async Task SaveAsync()
    {
        if (_manager.Current.FilePath is null)
        {
            await SaveAsAsync();
            return;
        }
        if (!await ValidateBeforeSaveAsync())
        {
            return;
        }
        await DoSaveAsync(_manager.Current.FilePath);
    }
    private async Task SaveAsAsync()
    {
        if (!await ValidateBeforeSaveAsync())
        {
            return;
        }
        string name = _manager.Current.FilePath is not null
            ? System.IO.Path.GetFileName(_manager.Current.FilePath)
            : _manager.Current.FileType == ArchiveType.MSND ? "archive.msnd" : "archive.dat";
        string? path = await PickSaveFileAsync("Save Archive As", name);
        if (path is not null)
        {
            await DoSaveAsync(path);
        }
    }
    private async Task<bool> ValidateBeforeSaveAsync()
    {
        if (_manager.IsRootEmpty())
        {
            string archiveType = _manager.Current.FileType == ArchiveType.DSARC ? "DSARC" : "DSEQ";
            await DialogService.ShowErrorAsync(this, $"Cannot save an empty {archiveType} archive.\n\nPlease add content before saving.");
            return false;
        }
        if (_manager.HasBlankFiles())
        {
            string message = "The archive contains blank files or empty containers.\n\n" +
                             "Yes: Remove all blank files/empty folders and save\n" +
                             "No: Save without removing blank files\n" +
                             "Cancel: Do not save";
            ConfirmResult result = await DialogService.ShowConfirmWithCancelAsync(this, "Blank Files Detected", message);
            if (result == ConfirmResult.Yes)
            {
                await _manager.RemoveAllEmptyEntriesAsync();
                Log("Removed blank files and empty archives.");
            }
            else if (result == ConfirmResult.Cancel)
            {
                return false;
            }
        }
        return true;
    }
    private async Task DoSaveAsync(string path)
    {
        try
        {
            Progress<double> progress = new(v =>
            {
                Dispatcher.UIThread.Post(() => _progressBar.Value = v * 100);
            });
            await RunAsync(ct => _manager.SaveArchiveAsync(path, progress, ct), "Saved");
        }
        catch (Exception ex) { await DialogService.ShowErrorAsync(this, ex.Message); }
    }
    private async Task ImportFolderAsync()
    {
        string? folder = await PickFolderAsync("Select Folder");
        if (folder is null)
        {
            return;
        }
        try
        {
            _expansionState.Clear();
            await RunAsync(ct => _manager.ImportFolderAsync(folder, ct), "Imported");
        }
        catch (Exception ex) { await DialogService.ShowErrorAsync(this, ex.Message); }
    }
    private async Task ImportToNestedAsync(ArchiveEntry entry)
    {
        string? folder = await PickFolderAsync("Select Folder");
        if (folder is null)
        {
            return;
        }
        try { await RunAsync(ct => _manager.ImportToNestedAsync(entry, folder, ct), $"Imported to {entry.Name}"); }
        catch (Exception ex) { await DialogService.ShowErrorAsync(this, ex.Message); }
    }
    private async Task ExtractAllAsync(bool nested)
    {
        string? folder = await PickFolderAsync("Select Output Folder");
        if (folder is null)
        {
            return;
        }
        try
        {
            Progress<(int C, int T)> p = new(x => _progressBar.Value = x.C * 100.0 / Math.Max(1, x.T));
            await RunAsync(ct => _manager.ExtractAllAsync(folder, nested, p, ct), "Extracted");
        }
        catch (Exception ex) { await DialogService.ShowErrorAsync(this, ex.Message); }
    }
    private async Task ExtractNestedAsync(ArchiveEntry entry)
    {
        string? folder = await PickFolderAsync("Select Output Folder");
        if (folder is null)
        {
            return;
        }
        try
        {
            Progress<(int C, int T)> p = new(x => _progressBar.Value = x.C * 100.0 / Math.Max(1, x.T));
            await RunAsync(ct => _manager.ExtractNestedAsync(entry, folder, p, ct), $"Extracted {entry.Name}");
        }
        catch (Exception ex) { await DialogService.ShowErrorAsync(this, ex.Message); }
    }
    private async Task ExtractSingleAsync(ArchiveEntry entry)
    {
        string? folder = await PickFolderAsync("Select Output Folder");
        if (folder is null)
        {
            return;
        }
        try { await RunAsync(ct => _manager.ExtractSingleAsync(entry, folder, ct), $"Extracted {entry.Name}"); }
        catch (Exception ex) { await DialogService.ShowErrorAsync(this, ex.Message); }
    }
    private async Task ReplaceEntryAsync(ArchiveEntry entry)
    {
        string? path = await PickOpenFileAsync("Select File", false);
        if (path is null)
        {
            return;
        }
        try
        {
            byte[] data = await File.ReadAllBytesAsync(path);
            _ = await _manager.ReplaceEntryAsync(entry, data);
            Log($"Replaced {entry.Name}");
        }
        catch (Exception ex) { await DialogService.ShowErrorAsync(this, ex.Message); }
    }
    private async Task ReplaceChunkAsync(ArchiveEntry parent, ArchiveEntry chunk)
    {
        string? path = await PickOpenFileAsync("Select File", false);
        if (path is null)
        {
            return;
        }
        try
        {
            byte[] data = await File.ReadAllBytesAsync(path);
            _ = await _manager.ReplaceChunkAsync(parent, chunk, data);
            Log($"Replaced {chunk.Name}");
        }
        catch (Exception ex) { await DialogService.ShowErrorAsync(this, ex.Message); }
    }
    private async Task RenameEntryAsync(ArchiveEntry entry)
    {
        string? name = await DialogService.ShowInputDialogAsync(this, "Rename", "Enter new name:", entry.Name);
        if (string.IsNullOrEmpty(name) || name == entry.Name)
        {
            return;
        }
        if (_manager.HasDuplicateName(entry, name))
        {
            bool proceed = await DialogService.ShowConfirmAsync(this, "Duplicate Name",
                $"A file named '{name}' already exists at this level. Do you want to continue anyway?");
            if (!proceed)
            {
                return;
            }
        }
        await _manager.RenameEntryAsync(entry, name);
        Log($"Renamed to {name}");
    }
    private async Task AddBlankFileAsync(ArchiveEntry? parent)
    {
        string name = GetUniqueName(parent, "new_file", "");
        await _manager.AddBlankEntryAsync(parent, name, false, null);
        Log("Added blank file");
    }
    private async Task AddBlankContainerAsync(ArchiveEntry? parent, ArchiveType type)
    {
        (string baseName, string ext) = type == ArchiveType.DSARC ? ("new_archive", ".dat") : ("new_sound", ".msnd");
        string name = GetUniqueName(parent, baseName, ext);
        await _manager.AddBlankEntryAsync(parent, name, true, type);
        Log($"Added blank {type}");
    }
    private string GetUniqueName(ArchiveEntry? parent, string baseName, string ext)
    {
        HashSet<string> names = (parent?.Children ?? _manager.Current.RootEntry.Children)
            .Select(c => c.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        string candidate = $"{baseName}{ext}";
        if (!names.Contains(candidate))
        {
            return candidate;
        }
        for (int i = 1; ; i++)
        {
            candidate = $"{baseName}_{i}{ext}";
            if (!names.Contains(candidate))
            {
                return candidate;
            }
        }
    }
    private async Task DeleteEntryAsync(ArchiveEntry entry)
    {
        string msg = entry.Children.Count > 0
            ? $"Delete '{entry.Name}' and {entry.Children.Count} items inside?"
            : $"Delete '{entry.Name}'?";
        if (!await DialogService.ShowConfirmAsync(this, "Delete", msg))
        {
            return;
        }
        await _manager.DeleteEntryAsync(entry);
        Log($"Deleted {entry.Name}");
    }
    private void SaveExpansionState()
    {
        if (_rootItem is null)
        {
            return;
        }
        SaveExpansionStateRecursive(_rootItem);
    }
    private void SaveExpansionStateRecursive(TreeViewItem item)
    {
        if (item.DataContext is ArchiveEntry entry)
        {
            _expansionState[entry.Id] = item.IsExpanded;
        }
        if (item.ItemsSource is IEnumerable<TreeViewItem> children)
        {
            foreach (TreeViewItem child in children)
            {
                SaveExpansionStateRecursive(child);
            }
        }
    }
    private void RefreshTree()
    {
        Dispatcher.UIThread.Post(() =>
        {
            SaveExpansionState();
            ArchiveDocument doc = _manager.Current;
            if (!doc.HasContent && doc.FilePath is null)
            {
                _treeView.ItemsSource = null;
                _rootItem = null;
                UpdateModifiedIndicator();
                UpdateEditMenu();
                return;
            }
            string rootName = doc.FilePath is not null
                ? System.IO.Path.GetFileName(doc.FilePath)
                : doc.FileType == ArchiveType.MSND ? "New DSEQ" : "New DSARC";
            if (doc.IsModified)
            {
                rootName += " *";
            }
            _rootItem = new TreeViewItem
            {
                Header = CreateHeader(rootName, doc.IsModified),
                IsExpanded = true,
                ItemsSource = doc.RootEntry.Children.Select((e, i) => { e.ImportOrder = i; return CreateNode(e); }).ToList()
            };
            _treeView.ItemsSource = new[] { _rootItem };
            UpdateModifiedIndicator();
            UpdateEditMenu();
        });
    }
    private TreeViewItem CreateNode(ArchiveEntry entry)
    {
        string name = entry.IsModified ? $"{entry.Name} *" : entry.Name;
        bool shouldExpand = !_expansionState.TryGetValue(entry.Id, out bool savedState) || savedState;
        TreeViewItem node = new()
        {
            Header = CreateHeader(name, entry.IsModified),
            DataContext = entry,
            IsExpanded = shouldExpand
        };
        if (entry.Children.Count > 0)
        {
            node.ItemsSource = entry.Children.Select((c, i) => { c.ImportOrder = i; return CreateNode(c); }).ToList();
        }
        return node;
    }
    private static TextBlock CreateHeader(string text, bool modified)
    {
        TextBlock header = new() { Text = text };
        if (modified)
        {
            header.Foreground = ModifiedBrush;
        }
        return header;
    }
    private void UpdateModifiedIndicator()
    {
        _modifiedLabel.Text = _manager.Current.IsModified ? "● Modified" : "";
    }
    private void UpdateEditMenu()
    {
        _undoMenu.IsEnabled = _manager.CanUndo;
        _redoMenu.IsEnabled = _manager.CanRedo;
        _sortMenu.IsEnabled = _manager.Current.FileType == ArchiveType.DSARC && _manager.Current.RootEntry.Children.Count > 1;
        _sortRecursiveMenu.IsEnabled = _manager.Current.FileType == ArchiveType.DSARC && _manager.HasNestedArchives();
    }
    private async Task RunAsync(Func<CancellationToken, Task> action, string msg)
    {
        _cts?.Cancel();
        _cts = new CancellationTokenSource();
        try
        {
            await action(_cts.Token);
            _statusLabel.Text = msg;
            Log(msg);
        }
        catch (OperationCanceledException) { Log("Cancelled"); }
        finally { _cts = null; _progressBar.Value = 0; }
    }
    private void Log(string msg)
    {
        Dispatcher.UIThread.Post(() =>
        {
            _logTextBox.Text += $"[{DateTime.Now:HH:mm:ss}] {msg}\n";
            string text = _logTextBox.Text ?? string.Empty;
            int lineCount = text.Count(c => c == '\n');
            if (lineCount > MaxLogLines)
            {
                int removeLength = text.IndexOf('\n') + 1;
                _logTextBox.Text = text[removeLength..];
            }
            _logTextBox.CaretIndex = _logTextBox.Text?.Length ?? 0;
        });
    }
}