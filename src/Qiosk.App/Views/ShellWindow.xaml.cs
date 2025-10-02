using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Qiosk.App.ViewModels;

namespace Qiosk.App.Views;

public partial class ShellWindow : Window
{
    private ShellViewModel? ViewModel => DataContext as ShellViewModel;

    public ShellWindow()
    {
        InitializeComponent();
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        if (ViewModel is { } vm)
        {
            await vm.InitializeAsync();
            Topmost = vm.IsLockMode;
            vm.PropertyChanged += ViewModelOnPropertyChanged;
        }
    }

    private void OnClosing(object? sender, CancelEventArgs e)
    {
        if (ViewModel is { } vm)
        {
            vm.PropertyChanged -= ViewModelOnPropertyChanged;
            vm.Dispose();
        }
    }

    private void AdminPasswordBox_OnPasswordChanged(object sender, RoutedEventArgs e)
    {
        if (ViewModel is not { } vm)
        {
            return;
        }

        if (sender is PasswordBox passwordBox && vm.AdminPasswordInput != passwordBox.Password)
        {
            vm.AdminPasswordInput = passwordBox.Password;
        }
    }

    private void ViewModelOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (ViewModel is not { } vm)
        {
            return;
        }

        if (e.PropertyName == nameof(ShellViewModel.AdminPasswordInput))
        {
            if (AdminPasswordBox.Password != vm.AdminPasswordInput)
            {
                AdminPasswordBox.Password = vm.AdminPasswordInput;
            }
        }
        else if (e.PropertyName == nameof(ShellViewModel.IsAdminVisible))
        {
            AdminPasswordBox.Password = string.Empty;
        }
        else if (e.PropertyName == nameof(ShellViewModel.IsLockMode))
        {
            Topmost = vm.IsLockMode;
        }
    }
}


