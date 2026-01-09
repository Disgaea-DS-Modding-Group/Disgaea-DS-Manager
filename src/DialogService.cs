using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using System.Reflection;
namespace Disgaea_DS_Manager.Services;

public enum ConfirmResult { Yes, No, Cancel }
public static class DialogService
{
    private const int DefaultDialogWidth = 400;
    private const int DefaultMargin = 20;
    public static async Task ShowErrorAsync(Window? owner, string message)
    {
        await ShowMessageAsync(owner, "Error", message).ConfigureAwait(true);
    }
    public static async Task ShowWarningAsync(Window? owner, string message)
    {
        await ShowMessageAsync(owner, "Warning", message).ConfigureAwait(true);
    }
    public static async Task ShowMessageAsync(Window? owner, string title, string message)
    {
        if (owner is null)
        {
            return;
        }
        Window dialog = CreateBaseDialog(title, DefaultDialogWidth, 150);
        StackPanel stack = new() { Margin = new Thickness(DefaultMargin) };
        stack.Children.Add(new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, DefaultMargin) });
        Button button = new() { Content = "OK", Width = 80, HorizontalAlignment = HorizontalAlignment.Right };
        button.Click += (_, _) => dialog.Close();
        stack.Children.Add(button);
        dialog.Content = stack;
        await dialog.ShowDialog(owner).ConfigureAwait(true);
    }
    public static async Task ShowInfoAsync(Window? owner)
    {
        if (owner is null)
        {
            return;
        }
        Assembly assembly = Assembly.GetExecutingAssembly();
        Version? version = assembly.GetName().Version;
        string versionStr = version is not null ? $"{version.Major}.{version.Minor}.{version.Build}" : "Unknown";
        AssemblyCompanyAttribute? companyAttr = assembly.GetCustomAttribute<AssemblyCompanyAttribute>();
        string message = $"Disgaea DS Manager\nVersion: {versionStr}\n\nEdit archives used in Disgaea DS.\n\n";
        if (companyAttr is not null)
        {
            message += $"By: {companyAttr.Company}\n";
        }
        await ShowMessageAsync(owner, "About", message).ConfigureAwait(true);
    }
    public static async Task<ArchiveTypeChoice?> ShowNewFileDialogAsync(Window? owner)
    {
        if (owner is null)
        {
            return null;
        }
        TaskCompletionSource<ArchiveTypeChoice?> result = new();
        Window dialog = CreateBaseDialog("Create New Archive", 320, 180);
        RadioButton dsarcRadio = new() { Content = "DSARC (Container Archive)", IsChecked = true, Margin = new Thickness(0, 5) };
        RadioButton dseqRadio = new() { Content = "DSEQ / MSND (Sound Archive)", Margin = new Thickness(0, 5) };
        StackPanel buttonPanel = CreateButtonPanel();
        Button createButton = new() { Content = "Create", Width = 80, Margin = new Thickness(0, 0, 10, 0) };
        Button cancelButton = new() { Content = "Cancel", Width = 80 };
        createButton.Click += (_, _) =>
        {
            _ = result.TrySetResult(dsarcRadio.IsChecked == true ? ArchiveTypeChoice.DSARC : ArchiveTypeChoice.DSEQ);
            dialog.Close();
        };
        cancelButton.Click += (_, _) => { _ = result.TrySetResult(null); dialog.Close(); };
        buttonPanel.Children.Add(createButton);
        buttonPanel.Children.Add(cancelButton);
        StackPanel stack = new() { Margin = new Thickness(DefaultMargin) };
        stack.Children.Add(new TextBlock { Text = "Choose archive type:", FontWeight = FontWeight.SemiBold, Margin = new Thickness(0, 0, 0, 10) });
        stack.Children.Add(dsarcRadio);
        stack.Children.Add(dseqRadio);
        stack.Children.Add(buttonPanel);
        dialog.Content = stack;
        dialog.Closed += (_, _) => result.TrySetResult(null);
        await dialog.ShowDialog(owner).ConfigureAwait(true);
        return await result.Task.ConfigureAwait(false);
    }
    public static async Task<bool> ShowConfirmAsync(Window? owner, string title, string message)
    {
        if (owner is null)
        {
            return false;
        }
        TaskCompletionSource<bool> result = new();
        Window dialog = CreateBaseDialog(title, DefaultDialogWidth, 150);
        StackPanel stack = new() { Margin = new Thickness(DefaultMargin) };
        stack.Children.Add(new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, DefaultMargin) });
        StackPanel buttonPanel = CreateButtonPanel();
        Button yesButton = new() { Content = "Yes", Width = 80, Margin = new Thickness(0, 0, 10, 0) };
        Button noButton = new() { Content = "No", Width = 80 };
        yesButton.Click += (_, _) => { _ = result.TrySetResult(true); dialog.Close(); };
        noButton.Click += (_, _) => { _ = result.TrySetResult(false); dialog.Close(); };
        buttonPanel.Children.Add(yesButton);
        buttonPanel.Children.Add(noButton);
        stack.Children.Add(buttonPanel);
        dialog.Content = stack;
        dialog.Closed += (_, _) => result.TrySetResult(false);
        await dialog.ShowDialog(owner).ConfigureAwait(true);
        return await result.Task.ConfigureAwait(false);
    }
    public static async Task<ConfirmResult> ShowConfirmWithCancelAsync(Window? owner, string title, string message)
    {
        if (owner is null)
        {
            return ConfirmResult.Cancel;
        }
        TaskCompletionSource<ConfirmResult> result = new();
        int initialHeight = 180;
        Window dialog = CreateBaseDialog(title, DefaultDialogWidth, initialHeight);
        StackPanel stack = new() { Margin = new Thickness(DefaultMargin) };
        TextBlock messageBlock = new() { Text = message, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, DefaultMargin) };
        stack.Children.Add(messageBlock);
        StackPanel buttonPanel = CreateButtonPanel();
        Button yesButton = new() { Content = "Yes", Width = 80, Margin = new Thickness(0, 0, 10, 0) };
        Button noButton = new() { Content = "No", Width = 80, Margin = new Thickness(0, 0, 10, 0) };
        Button cancelButton = new() { Content = "Cancel", Width = 80 };
        yesButton.Click += (_, _) => { _ = result.TrySetResult(ConfirmResult.Yes); dialog.Close(); };
        noButton.Click += (_, _) => { _ = result.TrySetResult(ConfirmResult.No); dialog.Close(); };
        cancelButton.Click += (_, _) => { _ = result.TrySetResult(ConfirmResult.Cancel); dialog.Close(); };
        buttonPanel.Children.Add(yesButton);
        buttonPanel.Children.Add(noButton);
        buttonPanel.Children.Add(cancelButton);
        stack.Children.Add(buttonPanel);
        dialog.Content = stack;
        dialog.Closed += (_, _) => result.TrySetResult(ConfirmResult.Cancel);
        await dialog.ShowDialog(owner).ConfigureAwait(true);
        return await result.Task.ConfigureAwait(false);
    }
    public static async Task<string?> ShowInputDialogAsync(Window? owner, string title, string prompt, string defaultValue = "")
    {
        if (owner is null)
        {
            return null;
        }
        TaskCompletionSource<string?> result = new();
        Window dialog = CreateBaseDialog(title, DefaultDialogWidth, 180);
        StackPanel stack = new() { Margin = new Thickness(DefaultMargin) };
        stack.Children.Add(new TextBlock { Text = prompt, Margin = new Thickness(0, 0, 0, 10) });
        TextBox textBox = new() { Text = defaultValue, Margin = new Thickness(0, 0, 0, DefaultMargin) };
        stack.Children.Add(textBox);
        StackPanel buttonPanel = CreateButtonPanel();
        Button okButton = new() { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 10, 0) };
        Button cancelButton = new() { Content = "Cancel", Width = 80 };
        okButton.Click += (_, _) => { _ = result.TrySetResult(textBox.Text); dialog.Close(); };
        cancelButton.Click += (_, _) => { _ = result.TrySetResult(null); dialog.Close(); };
        buttonPanel.Children.Add(okButton);
        buttonPanel.Children.Add(cancelButton);
        stack.Children.Add(buttonPanel);
        dialog.Content = stack;
        dialog.Closed += (_, _) => result.TrySetResult(null);
        await dialog.ShowDialog(owner).ConfigureAwait(true);
        return await result.Task.ConfigureAwait(false);
    }
    private static Window CreateBaseDialog(string title, int width, int height, int? maxHeight = null)
    {
        Window window = new()
        {
            Title = title,
            Width = width,
            Height = height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SizeToContent = SizeToContent.Height
        };
        return window;
    }
    private static StackPanel CreateButtonPanel()
    {
        return new()
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, DefaultMargin, 0, 0)
        };
    }
}
public enum ArchiveTypeChoice { DSARC, DSEQ }