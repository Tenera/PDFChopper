using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using PdfChopper.Services;
using PdfChopper.ViewModels;
using PdfChopper.Views;
using SukiUI.Toasts;

namespace PdfChopper;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var toastManager = new SukiToastManager();
            var pdfService = new PdfService();
            var dialogService = new DialogService(toastManager);
            desktop.MainWindow = new MainWindow
            {
                DataContext = new MainWindowViewModel(toastManager, pdfService, dialogService),
            };
        }

        base.OnFrameworkInitializationCompleted();
    }
}
