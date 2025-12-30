using Avalonia;
namespace Disgaea_DS_Manager;

public static class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        _ = BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }
    public static AppBuilder BuildAvaloniaApp()
    {
        return AppBuilder.Configure<App>().UsePlatformDetect().WithInterFont().LogToTrace();
    }
}