using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using KillRiceMonkey.App.ViewModels;

namespace KillRiceMonkey.App;

public partial class MainWindow : Window
{
    private const int StartHotkeyId = 8100;
    private const int CancelHotkeyId = 8101;
    private const uint ModNone = 0x0000;
    private const uint VkF8 = 0x77;
    private const uint VkF9 = 0x78;
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
        RegisterHotKey(helper.Handle, StartHotkeyId, ModNone, VkF8);
        RegisterHotKey(helper.Handle, CancelHotkeyId, ModNone, VkF9);
    }

    private void OnClosed(object? sender, EventArgs e)
    {
        var helper = new WindowInteropHelper(this);
        UnregisterHotKey(helper.Handle, StartHotkeyId);
        UnregisterHotKey(helper.Handle, CancelHotkeyId);
        _hwndSource?.RemoveHook(WndProc);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WmHotkey)
        {
            var id = wParam.ToInt32();
            if (id == StartHotkeyId)
            {
                _ = _viewModel.HandleHotkeyAsync();
                handled = true;
            }
            else if (id == CancelHotkeyId)
            {
                _viewModel.CancelAutomation();
                handled = true;
            }
        }

        return IntPtr.Zero;
    }

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}
