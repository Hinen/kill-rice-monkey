using System.Windows;
using TicketingAutoPurchase.App.ViewModels;

namespace TicketingAutoPurchase.App;

public partial class MainWindow : Window
{
    public MainWindow(MainWindowViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
    }
}
