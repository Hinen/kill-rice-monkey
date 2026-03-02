using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using KillRiceMonkey.App.ViewModels;

namespace KillRiceMonkey.App;

public partial class MainWindow : Window
{
    private const int HotkeyId = 8100;
    private const uint ModNone = 0x0000;
    private const uint VkF8 = 0x77;
    private const int WmHotkey = 0x0312;

    private readonly MainWindowViewModel _viewModel;
    private HwndSource? _hwndSource;

    public MainWindow(MainWindowViewModel viewModel)
    {
        _viewModel = viewModel;
        InitializeComponent();
        DataContext = _viewModel;
        SourceInitialized += OnSourceInitialized;
        Closed += OnClosed;
    }

    private void OnSourceInitialized(object? sender, EventArgs e)
    {
        var helper = new WindowInteropHelper(this);
        _hwndSource = HwndSource.FromHwnd(helper.Handle);
        _hwndSource?.AddHook(WndProc);
        RegisterHotKey(helper.Handle, HotkeyId, ModNone, VkF8);
    }

    private void OnClosed(object? sender, EventArgs e)
    {
        var helper = new WindowInteropHelper(this);
        UnregisterHotKey(helper.Handle, HotkeyId);
        _hwndSource?.RemoveHook(WndProc);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WmHotkey && wParam.ToInt32() == HotkeyId)
        {
            _ = _viewModel.HandleHotkeyAsync();
            handled = true;
        }

        return IntPtr.Zero;
    }

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}
