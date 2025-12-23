using Avalonia;
using Avalonia.Collections;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using System.Collections.ObjectModel;
using System.Globalization;
namespace Disgaea_DS_Manager;

public partial class MainWindow : Window
{
    private string? _archivePath;
    private Collection<Entry> _entries = [];
    private string? _srcFolder;
    private ArchiveType? _filetype;
    private TreeViewItem? _rootItem;
    private bool _archiveOpenedFromDisk;
    private readonly IArchiveService _archiveService;
    private CancellationTokenSource? _cts;
    [Obsolete] public MainWindow() : this(new ArchiveService()) { }
    [Obsolete]
    public MainWindow(IArchiveService archiveService)
    {
        _archiveService = archiveService ?? throw new ArgumentNullException(nameof(archiveService));
        InitializeComponent();
        if (!Design.IsDesignMode)
        {
            WireUpEvents();
        }
    }
    [Obsolete]
    private void WireUpEvents()
    {
        AttachMenuItem("NewMenu", _ => NewArchive());
        AttachMenuItem("OpenMenu", async _ => await OpenArchiveAsync().ConfigureAwait(false));
        AttachMenuItem("SaveMenu", async _ => await SaveArchiveAsync().ConfigureAwait(false));
        AttachMenuItem("SaveAsMenu", async _ => await SaveAsAsync().ConfigureAwait(false));
        AttachMenuItem("ExitMenu", _ => Close());
        if (this.FindControl<TreeView>("TreeView") is { } tree)
        {
            tree.SelectionChanged += TreeView_SelectionChanged;
            tree.AddHandler(InputElement.PointerReleasedEvent, TreeView_PointerReleased, RoutingStrategies.Tunnel);
        }
    }
    private void AttachMenuItem(string name, Action<RoutedEventArgs> handler)
    {
        if (this.FindControl<MenuItem>(name) is { } menuItem)
        {
            menuItem.Click += (s, e) => handler(e);
        }
    }
    private void AttachMenuItem(string name, Func<RoutedEventArgs, Task> handler)
    {
        if (this.FindControl<MenuItem>(name) is { } menuItem)
        {
            menuItem.Click += async (s, e) => await handler(e);
        }
    }
    #region UI Helpers
    [Obsolete]
    private async Task<string?> SelectFolderAsync(string? title = null)
    {
        OpenFolderDialog dlg = new() { Title = title ?? "Select Folder" };
        try { return await dlg.ShowAsync(this); } catch { return null; }
    }
    [Obsolete]
    private async Task<string?> SelectFileAsync(string filter = "All Files (*.*)|*.*")
    {
        OpenFileDialog dlg = new()
        {
            AllowMultiple = false,
            Filters = filter.Contains("All Files")
                ? [new FileDialogFilter { Name = "All Files", Extensions = ["*"] }]
                : [
                    new FileDialogFilter { Name = "Disgaea DS Archives", Extensions = ["dat", "msnd"] },
                    new FileDialogFilter { Name = "All Files", Extensions = ["*"] }
                ]
        };
        try
        {
            string[]? res = await dlg.ShowAsync(this);
            return res?.FirstOrDefault();
        }
        catch { return null; }
    }
    [Obsolete]
    private async Task<string?> SelectSaveFileAsync(string filter = "All Files (*.*)|*.*")
    {
        SaveFileDialog dlg = new();
        try { return await dlg.ShowAsync(this); } catch { return null; }
    }
    private async Task ShowMessageDialogAsync(string text, string title = "")
    {
        await Dispatcher.UIThread.InvokeAsync(async () =>
        {
            Window dialog = new()
            {
                Title = title,
                Width = 480,
                Height = 160,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                SizeToContent = SizeToContent.Manual
            };
            StackPanel panel = new()
            {
                Margin = new Thickness(20),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Stretch
            };
            TextBlock textBlock = new()
            {
                Text = text,
                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Stretch
            };
            StackPanel buttonPanel = new()
            {
                Orientation = Avalonia.Layout.Orientation.Horizontal,
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Bottom,
                Margin = new Thickness(0, 20, 0, 0)
            };
            Button okButton = new()
            {
                Content = "OK",
                Width = 80,
                IsDefault = true,
                IsCancel = true
            };
            okButton.Click += (s, e) => dialog.Close();
            buttonPanel.Children.Add(okButton);
            panel.Children.Add(textBlock);
            panel.Children.Add(buttonPanel);
            dialog.Content = panel;
            await dialog.ShowDialog(this);
            AppendLog($"{(string.IsNullOrEmpty(title) ? "" : title + ": ")}{text}");
        });
    }
    private Task ShowErrorAsync(string msg, string caption = "Error")
    {
        return ShowMessageDialogAsync(msg, caption);
    }

    private Task ShowWarningAsync(string msg, string caption = "Warning")
    {
        return ShowMessageDialogAsync(msg, caption);
    }

    private void AppendLog(string msg)
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (this.FindControl<TextBox>("LogTextBox") is { } log)
            {
                log.Text += $"{msg}\r\n";
                log.CaretIndex = log.Text?.Length ?? 0;
            }
        });
    }
    private void UpdateProgress(int val, int total)
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (this.FindControl<ProgressBar>("ProgressBar") is not { } pb)
            {
                return;
            }

            try
            {
                int t = Math.Max(1, total);
                pb.Maximum = t;
                pb.Value = Math.Clamp(val, 0, t);
            }
            catch (ArgumentOutOfRangeException) { }
        });
    }
    private void SetStatus(string text)
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (this.FindControl<TextBlock>("StatusLabel") is { } lbl)
            {
                lbl.Text = text;
            }
        });
    }
    #endregion
    #region Archive Operations
    [Obsolete]
    private async void NewArchive()
    {
        try
        {
            string? file = await SelectSaveFileAsync().ConfigureAwait(false);
            if (string.IsNullOrEmpty(file))
            {
                return;
            }

            _archivePath = file;
            _entries = [];
            _srcFolder = null;
            _filetype = _archivePath.EndsWith(".msnd", StringComparison.OrdinalIgnoreCase) ? ArchiveType.MSND : ArchiveType.DSARC;
            _archiveOpenedFromDisk = false;
            RefreshTree();
            SetStatus("New archive created");
            AppendLog($"Created new {_filetype} archive: {Path.GetFileName(_archivePath)}");
        }
        catch (Exception ex)
        {
            await ShowErrorAsync($"Failed to create new archive: {ex.Message}");
        }
    }
    [Obsolete]
    private async Task OpenArchiveAsync()
    {
        try
        {
            string? file = await SelectFileAsync("Disgaea DS Archives (*.dat, *.msnd)|*.dat;*.msnd").ConfigureAwait(false);
            if (file is null)
            {
                return;
            }

            _archivePath = file;
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            _entries = await _archiveService.LoadArchiveAsync(_archivePath, _cts.Token).ConfigureAwait(false);
            _filetype = Detector.FromFile(_archivePath);
            _archiveOpenedFromDisk = true;
            RefreshTree();
            SetStatus($"Opened {Path.GetFileName(_archivePath)}");
            AppendLog($"Opened {Path.GetFileName(_archivePath)} as {_filetype.ToString().ToUpper(CultureInfo.InvariantCulture)}");
        }
        catch (OperationCanceledException) { AppendLog("Open archive cancelled."); }
        catch (Exception ex) { await ShowErrorAsync($"Failed to open archive: {ex.Message}"); }
        finally { _cts = null; }
    }
    [Obsolete]
    private async Task SaveArchiveAsync()
    {
        try
        {
            if (_archivePath is null)
            {
                await SaveAsAsync().ConfigureAwait(false);
                return;
            }
            if (_entries?.Count == 0)
            {
                await ShowWarningAsync("No files in archive to save.");
                return;
            }
            if (_srcFolder is null && !_archiveOpenedFromDisk && HasEmbeddedMsnd())
            {
                await ShowWarningAsync("Cannot save archive with embedded MSND replacements without a source folder.");
                return;
            }
            _srcFolder ??= await SelectFolderAsync("Select temporary folder for archive contents");
            if (_srcFolder is null)
            {
                return;
            }

            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
            string? originalArchivePath = _archiveOpenedFromDisk ? _archivePath : null;
            await _archiveService.SaveArchiveAsync(_archivePath, _filetype!.Value, _entries.ToList(),
                _srcFolder, progress, _cts.Token, originalArchivePath).ConfigureAwait(false);
            SetStatus(_filetype == ArchiveType.MSND ? "MSND saved" : "DSARC saved");
            AppendLog($"{_filetype.ToString().ToUpper(CultureInfo.InvariantCulture)} saved.");
        }
        catch (OperationCanceledException) { AppendLog("Save cancelled."); }
        catch (Exception ex) { await ShowErrorAsync($"Failed to save archive: {ex.Message}"); }
        finally { _cts = null; UpdateProgress(0, 1); }
    }
    private bool HasEmbeddedMsnd()
    {
        return _entries?.Any(e => e.IsMsnd && e.Children.Count > 0) == true;
    }

    [Obsolete]
    private async Task SaveAsAsync()
    {
        string? file = await SelectSaveFileAsync().ConfigureAwait(false);
        if (file is null)
        {
            return;
        }

        _archivePath = file;
        await SaveArchiveAsync().ConfigureAwait(false);
    }
    private void RefreshTree()
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (this.FindControl<TreeView>("TreeView") is not { } tree)
            {
                return;
            }

            string rootText = _archivePath != null ? Path.GetFileName(_archivePath) : "[New Archive]";
            _rootItem = new TreeViewItem { Header = rootText, DataContext = null, IsExpanded = true };
            AvaloniaList<TreeViewItem> rootChildren = [];
            if (_filetype.HasValue)
            {
                rootChildren.AddRange(_filetype.Value == ArchiveType.DSARC
                    ? CreateDsarcTreeItems()
                    : CreateMsndTreeItems());
            }
            _ = _rootItem.SetValue(ItemsControl.ItemsSourceProperty, rootChildren);
            _ = tree.SetValue(ItemsControl.ItemsSourceProperty, new AvaloniaList<TreeViewItem> { _rootItem });
            UpdateContextMenuVisibility(tree.SelectedItem);
        });
    }
    private IEnumerable<TreeViewItem> CreateDsarcTreeItems()
    {
        foreach (Entry e in _entries)
        {
            if (e.IsMsnd && e.Children.Count > 0)
            {
                TreeViewItem msNode = new() { Header = e.Path.Name, DataContext = e, IsExpanded = true };
                AvaloniaList<TreeViewItem> childList = [.. e.Children.Select(c => new TreeViewItem { Header = c.Path.Name, DataContext = c })];
                _ = msNode.SetValue(ItemsControl.ItemsSourceProperty, childList);
                yield return msNode;
            }
            else
            {
                yield return new TreeViewItem { Header = e.Path.Name, DataContext = e };
            }
        }
    }
    private IEnumerable<TreeViewItem> CreateMsndTreeItems()
    {
        return _entries.Select(e => new TreeViewItem { Header = e.Path.Name, DataContext = e });
    }
    #endregion
    #region Tree & Context Menu
    private void TreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (this.FindControl<TreeView>("TreeView") is { } tree)
        {
            UpdateContextMenuVisibility(tree.SelectedItem);
        }
    }
    [Obsolete]
    private void TreeView_PointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        if (e.InitialPressMouseButton != MouseButton.Right)
        {
            return;
        }

        if (this.FindControl<TreeView>("TreeView") is { } tree && tree.SelectedItem is TreeViewItem treeViewItem)
        {
            ShowContextMenu(treeViewItem);
            e.Handled = true;
        }
    }
    private void UpdateContextMenuVisibility(object? selectedItem) { }
    [Obsolete]
    private void ShowContextMenu(TreeViewItem node)
    {
        ContextMenu menu = new();
        AvaloniaList<MenuItem> items = [];
        bool isRoot = node == _rootItem;
        if (isRoot)
        {
            items.Add(CreateMenuItem("Import Folder", ImportFolderAsync));
            items.Add(CreateMenuItem("Extract All", ExtractAllAsync));
            if (_filetype == ArchiveType.DSARC)
            {
                items.Add(CreateMenuItem("Extract All (Nested)", ExtractAllNestedRootAsync));
            }
        }
        else
        {
            bool nodeHasChildren = node.Items?.Cast<object>().Any() == true;
            if (nodeHasChildren && node.DataContext is Entry { IsMsnd: true } entry)
            {
                items.Add(CreateMenuItem("Import Folder", async () => await ImportFolderToNodeAsync(entry)));
                items.Add(CreateMenuItem("Extract All", async () => await ExtractAllFromNodeAsync(entry)));
                items.Add(CreateMenuItem("Extract All Nested", async () => await ExtractAllNestedFromNodeAsync(entry)));
            }
            if (_filetype == ArchiveType.DSARC)
            {
                if (node.Parent == _rootItem)
                {
                    items.Add(CreateMenuItem("Extract", async () => await ExtractItemAsync(node)));
                    items.Add(CreateMenuItem("Replace", async () => await ReplaceItemAsync(node)));
                }
                else if (node.Parent != null)
                {
                    items.Add(CreateMenuItem("Extract (chunk)", async () => await ExtractChunkItemAsync(node)));
                    items.Add(CreateMenuItem("Replace (chunk)", async () => await ReplaceChunkItemAsync(node)));
                }
            }
            else if (_filetype == ArchiveType.MSND)
            {
                items.Add(CreateMenuItem("Extract", async () => await ExtractItemAsync(node)));
                items.Add(CreateMenuItem("Replace", async () => await ReplaceItemAsync(node)));
            }
        }
        _ = menu.SetValue(ItemsControl.ItemsSourceProperty, items);
        menu.PlacementTarget = node;
        menu.Open(node);
    }
    private MenuItem CreateMenuItem(string header, Func<Task> action)
    {
        MenuItem menuItem = new() { Header = header };
        menuItem.Click += async (s, e) => await action();
        return menuItem;
    }
    #endregion
    #region Import/Export/Replace Methods (Complete Implementation)
    [Obsolete]
    private async Task ImportFolderAsync()
    {
        string? folder = await SelectFolderAsync("Select folder to import").ConfigureAwait(false);
        if (folder is null)
        {
            return;
        }

        try
        {
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            ImportResult result = await _archiveService.InspectFolderForImportAsync(folder, _cts.Token).ConfigureAwait(false);
            if (result?.Entries == null || result.Entries.Count == 0)
            {
                await ShowWarningAsync("No files found to import");
                return;
            }
            _entries = result.Entries;
            _filetype = result.FileType;
            _srcFolder = result.SourceFolder;
            RefreshTree();
            SetStatus("Folder imported");
            AppendLog("Folder imported (nested-aware).");
        }
        catch (OperationCanceledException) { AppendLog("Import folder cancelled."); }
        catch (Exception ex) { await ShowErrorAsync(ex.Message); }
        finally { _cts = null; }
    }
    [Obsolete]
    private async Task ExtractAllAsync()
    {
        if (_archivePath is null || _filetype is null)
        {
            await ShowWarningAsync("Open archive first");
            return;
        }
        string? destFolder = await SelectFolderAsync("Choose output folder for Extract All").ConfigureAwait(false);
        if (destFolder is null)
        {
            return;
        }

        try
        {
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
            (string outBase, List<string> mapperLines) = await _archiveService.ExtractAllAsync(_archivePath, _filetype.Value, destFolder, progress, _cts.Token).ConfigureAwait(false);
            AppendLog($"Starting Extract All -> {outBase}");
            SetStatus("Extract complete");
            AppendLog("Extract All complete.");
        }
        catch (OperationCanceledException) { AppendLog("Extract All cancelled."); }
        catch (Exception ex) { await ShowErrorAsync(ex.Message); }
        finally { _cts = null; UpdateProgress(0, 1); }
    }
    [Obsolete]
    private async Task ExtractItemAsync(TreeViewItem node)
    {
        if (_rootItem is null) { await ShowWarningAsync("Invalid selection"); return; }
        Entry? entry = GetEntryFromNode(node);
        if (entry is null) { await ShowWarningAsync("Invalid selection"); return; }
        string? dest = await SelectFolderAsync("Select folder to extract to").ConfigureAwait(false);
        if (dest is null)
        {
            return;
        }

        try
        {
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            await _archiveService.ExtractItemAsync(_archivePath!, _filetype!.Value, entry, dest, _cts.Token).ConfigureAwait(false);
            SetStatus($"Extracted {entry.Path.Name}");
            AppendLog($"Extracted {entry.Path.Name} ({entry.Size} bytes)");
        }
        catch (OperationCanceledException) { AppendLog("Extract item cancelled."); }
        catch (Exception ex) { await ShowErrorAsync(ex.Message); }
        finally { _cts = null; }
    }
    [Obsolete]
    private async Task ReplaceItemAsync(TreeViewItem node)
    {
        if (_rootItem is null) { await ShowWarningAsync("Invalid selection"); return; }
        Entry? entry = GetEntryFromNode(node);
        if (entry is null) { await ShowWarningAsync("Invalid selection"); return; }
        string? replacement = await SelectFileAsync().ConfigureAwait(false);
        if (replacement is null)
        {
            return;
        }

        try
        {
            if (_filetype == ArchiveType.MSND && !Msnd.MSND_ORDER.Contains(Path.GetExtension(replacement).ToLowerInvariant()))
            {
                await ShowWarningAsync("Replacement must be sseq/sbnk/swar");
                return;
            }
            if (_filetype == ArchiveType.MSND)
            {
                if (string.IsNullOrEmpty(_srcFolder))
                {
                    _srcFolder = Path.GetDirectoryName(replacement);
                    if (string.IsNullOrEmpty(_srcFolder))
                    {
                        await ShowErrorAsync("Cannot determine source folder");
                        return;
                    }
                }
                string targetPath = Path.Combine(_srcFolder, Path.GetFileName(replacement));
                if (!string.Equals(replacement, targetPath, StringComparison.OrdinalIgnoreCase))
                {
                    File.Copy(replacement, targetPath, true);
                }

                entry.Path = new FileInfo(Path.GetFileName(replacement));
            }
            else
            {
                _srcFolder ??= Path.GetDirectoryName(replacement);
                await _archiveService.CopyFileToFolderAsync(replacement, _srcFolder!, CancellationToken.None).ConfigureAwait(false);
                entry.Path = new FileInfo(Path.GetFileName(replacement));
            }
            RefreshTree();
            SetStatus("File replaced");
            AppendLog($"Replaced file with: {Path.GetFileName(replacement)}");
        }
        catch (Exception ex) { await ShowErrorAsync(ex.Message); }
    }
    [Obsolete]
    private async Task ExtractChunkItemAsync(TreeViewItem node)
    {
        if (node.Parent is not TreeViewItem parent || parent.DataContext is not Entry parentEntry)
        { await ShowWarningAsync("Invalid parent"); return; }
        if (node.DataContext is not Entry chunkEntry)
        { await ShowWarningAsync("Invalid chunk"); return; }
        string? dest = await SelectFolderAsync("Select folder to extract chunk to").ConfigureAwait(false);
        if (dest is null)
        {
            return;
        }

        try
        {
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            await _archiveService.ExtractChunkItemAsync(_archivePath!, parentEntry, chunkEntry, dest, _cts.Token).ConfigureAwait(false);
            SetStatus($"Extracted {chunkEntry.Path.Name}");
            AppendLog($"Extracted {chunkEntry.Path.Name} ({chunkEntry.Size} bytes)");
        }
        catch (OperationCanceledException) { AppendLog("Extract chunk cancelled."); }
        catch (Exception ex) { await ShowErrorAsync(ex.Message); }
        finally { _cts = null; }
    }
    [Obsolete]
    private async Task<byte[]> ReplaceChunkItemAsync(TreeViewItem node)
    {
        if (node.Parent is not TreeViewItem parent || parent.DataContext is not Entry parentEntry)
        { await ShowWarningAsync("Invalid parent"); return []; }
        if (node.DataContext is not Entry chunkEntry)
        { await ShowWarningAsync("Invalid chunk"); return []; }
        string? replacement = await SelectFileAsync().ConfigureAwait(false);
        if (replacement is null)
        {
            return [];
        }

        try
        {
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            if (string.IsNullOrEmpty(_srcFolder))
            {
                _srcFolder = Path.GetDirectoryName(replacement);
                if (string.IsNullOrEmpty(_srcFolder))
                {
                    await ShowErrorAsync("Cannot determine source folder for replacement");
                    return [];
                }
            }
            string msndTargetPath = Path.Combine(_srcFolder, parentEntry.Path.Name);
            if (!File.Exists(msndTargetPath))
            {
                await _archiveService.ExtractItemAsync(_archivePath!, ArchiveType.DSARC, parentEntry, _srcFolder, _cts.Token);
                AppendLog($"Extracted base MSND file to {msndTargetPath} for modification");
            }
            byte[] rebuilt = await _archiveService.ReplaceChunkItemAsync(
                _archivePath!, parentEntry, chunkEntry, replacement, _srcFolder, _cts.Token).ConfigureAwait(false);
            parentEntry.Children.Clear();
            foreach (Entry child in Msnd.Parse(rebuilt, Path.GetFileNameWithoutExtension(parentEntry.Path.Name)))
            {
                parentEntry.Children.Add(child);
            }

            RefreshTree();
            SetStatus("File replaced - use Save");
            AppendLog($"Rebuilt embedded MSND {parentEntry.Path.Name} after replacing {Path.GetExtension(chunkEntry.Path.Name)}");
            return rebuilt;
        }
        catch (OperationCanceledException) { AppendLog("Replace chunk cancelled."); return []; }
        catch (Exception ex) { await ShowErrorAsync(ex.Message); return []; }
        finally { _cts = null; }
    }
    [Obsolete]
    private async Task ExtractAllFromNodeAsync(Entry entry)
    {
        if (!entry.IsMsnd)
        {
            await ShowWarningAsync("Selection is not an embedded archive");
            return;
        }
        string? dest = await SelectFolderAsync("Select root folder for extraction").ConfigureAwait(false);
        if (dest is null)
        {
            return;
        }

        try
        {
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            string outDir = Path.Combine(dest, Path.GetFileNameWithoutExtension(entry.Path.Name));
            _ = Directory.CreateDirectory(outDir);
            byte[] msndBuf = await _archiveService.ReadRangeAsync(_archivePath!, entry.Offset, entry.Size, _cts.Token).ConfigureAwait(false);
            Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
            await _archiveService.NestedExtractBufferAsync(msndBuf, outDir, Path.GetFileNameWithoutExtension(entry.Path.Name), progress, _cts.Token).ConfigureAwait(false);
            if (entry.Children.Count == 0)
            {
                foreach (Entry child in Msnd.Parse(msndBuf, Path.GetFileNameWithoutExtension(entry.Path.Name)))
                {
                    entry.Children.Add(child);
                }
            }
            RefreshTree();
            SetStatus("Extract complete");
            AppendLog("Extract complete.");
        }
        catch (OperationCanceledException) { AppendLog("Nested extract cancelled."); }
        catch (Exception ex) { await ShowErrorAsync(ex.Message); }
        finally { _cts = null; UpdateProgress(0, 1); }
    }
    [Obsolete]
    private async Task ImportFolderToNodeAsync(Entry entry)
    {
        if (!entry.IsMsnd)
        {
            await ShowWarningAsync("Selection not an embedded archive");
            return;
        }
        string? folder = await SelectFolderAsync("Select folder containing msnd parts").ConfigureAwait(false);
        if (folder is null)
        {
            return;
        }

        try
        {
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            byte[] rebuilt = await _archiveService.BuildMsndFromFolderAsync(folder, _cts.Token).ConfigureAwait(false);
            _srcFolder ??= folder;
            string outPath = Path.Combine(_srcFolder, entry.Path.Name);
            await _archiveService.WriteFileAsync(outPath, rebuilt, _cts.Token).ConfigureAwait(false);
            entry.Children.Clear();
            foreach (Entry child in Msnd.Parse(rebuilt, Path.GetFileNameWithoutExtension(entry.Path.Name)))
            {
                entry.Children.Add(child);
            }

            RefreshTree();
            SetStatus("Imported and staged");
            AppendLog($"Imported and rebuilt embedded archive {entry.Path.Name} (staged to {outPath}).");
        }
        catch (OperationCanceledException) { AppendLog("Import to node cancelled."); }
        catch (Exception ex) { await ShowErrorAsync(ex.Message); }
        finally { _cts = null; }
    }
    [Obsolete]
    private async Task ExtractAllNestedRootAsync()
    {
        if (_archivePath is null || _filetype != ArchiveType.DSARC)
        {
            await ShowWarningAsync("Open DSARC first");
            return;
        }
        string? baseFolder = await SelectFolderAsync("Select base folder for nested extract").ConfigureAwait(false);
        if (baseFolder is null)
        {
            return;
        }

        string expected = Path.GetFileNameWithoutExtension(_archivePath);
        string? outdir = ResolveSelectedFolderForExpected(baseFolder, expected, out bool found, out int matches);
        if (!found)
        {
            string candidate = Path.Combine(baseFolder, expected);
            if (File.Exists(candidate))
            {
                await ShowErrorAsync($"File exists cannot create folder: {expected}");
                return;
            }
            _ = Directory.CreateDirectory(candidate);
            AppendLog($"Created folder '{candidate}' for nested extract.");
            outdir = candidate;
            matches = 1;
        }
        if (matches > 1)
        {
            AppendLog($"Multiple candidate folders matched '{expected}'; using '{outdir}'.");
        }

        try
        {
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            byte[] buf = await _archiveService.ReadFileAsync(_archivePath, _cts.Token).ConfigureAwait(false);
            Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
            await _archiveService.NestedExtractBufferAsync(buf, outdir, expected, progress, _cts.Token).ConfigureAwait(false);
            SetStatus("Nested extract complete");
            AppendLog("Nested extract complete.");
        }
        catch (OperationCanceledException) { AppendLog("Nested root extract cancelled."); }
        catch (Exception ex) { await ShowErrorAsync(ex.Message); }
        finally { _cts = null; UpdateProgress(0, 1); }
    }
    [Obsolete]
    private async Task ExtractAllNestedFromNodeAsync(Entry entry)
    {
        string? baseFolder = await SelectFolderAsync("Select base folder for nested extract from node").ConfigureAwait(false);
        if (baseFolder is null)
        {
            return;
        }

        string expected = Path.GetFileNameWithoutExtension(entry.Path.Name);
        string? outdir = ResolveSelectedFolderForExpected(baseFolder, expected, out bool found, out int matches);
        if (!found)
        {
            string candidate = Path.Combine(baseFolder, expected);
            if (File.Exists(candidate))
            {
                await ShowErrorAsync($"File exists cannot create folder: {expected}");
                return;
            }
            _ = Directory.CreateDirectory(candidate);
            AppendLog($"Created folder '{candidate}' for nested extract.");
            outdir = candidate;
            matches = 1;
        }
        if (matches > 1)
        {
            AppendLog($"Multiple candidate folders matched '{expected}'; using '{outdir}'.");
        }

        try
        {
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            byte[] buf = await _archiveService.ReadRangeAsync(_archivePath!, entry.Offset, entry.Size, _cts.Token).ConfigureAwait(false);
            Progress<(int current, int total)> progress = new(t => UpdateProgress(t.current, t.total));
            await _archiveService.NestedExtractBufferAsync(buf, outdir, expected, progress, _cts.Token).ConfigureAwait(false);
            SetStatus("Nested extract complete");
            AppendLog("Nested extract complete.");
        }
        catch (OperationCanceledException) { AppendLog("Nested extract from node cancelled."); }
        catch (Exception ex) { await ShowErrorAsync(ex.Message); }
        finally { _cts = null; UpdateProgress(0, 1); }
    }
    private Entry? GetEntryFromNode(TreeViewItem node)
    {
        if (_rootItem?.Items is null)
        {
            return null;
        }

        List<TreeViewItem> list = _rootItem.Items.Cast<TreeViewItem>().ToList();
        int idx = list.IndexOf(node);
        return idx >= 0 && idx < _entries.Count ? _entries[idx] : null;
    }
    private string? ResolveSelectedFolderForExpected(string selectedPath, string expectedFolderName, out bool found, out int matches)
    {
        found = false;
        matches = 0;
        if (string.IsNullOrEmpty(selectedPath))
        {
            return null;
        }

        try
        {
            if (string.Equals(Path.GetFileName(selectedPath), expectedFolderName, StringComparison.OrdinalIgnoreCase))
            {
                found = true; matches = 1; return selectedPath;
            }
            string direct = Path.Combine(selectedPath, expectedFolderName);
            if (Directory.Exists(direct))
            {
                found = true; matches = 1; return direct;
            }
            string[] topMatches = Directory.GetDirectories(selectedPath, expectedFolderName, SearchOption.TopDirectoryOnly);
            if (topMatches.Length > 0)
            {
                found = true; matches = topMatches.Length; return topMatches[0];
            }
            string[] recMatches = Directory.GetDirectories(selectedPath, expectedFolderName, SearchOption.AllDirectories);
            if (recMatches.Length > 0)
            {
                found = true; matches = recMatches.Length; return recMatches[0];
            }
        }
        catch (Exception ex) { AppendLog($"ResolveSelectedFolderForExpected error: {ex.Message}"); }
        return null;
    }
    #endregion
}