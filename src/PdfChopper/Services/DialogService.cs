using System;
using Avalonia.Controls.Notifications;
using SukiUI.Toasts;

namespace PdfChopper.Services;

public class DialogService : IDialogService
{
    private readonly ISukiToastManager _toastManager;

    public DialogService(ISukiToastManager toastManager)
    {
        _toastManager = toastManager;
    }

    public void ShowSuccess(string title, string message)
    {
        _toastManager.CreateToast()
            .WithTitle(title)
            .WithContent(message)
            .OfType(NotificationType.Success)
            .Dismiss().After(TimeSpan.FromSeconds(3))
            .Dismiss().ByClicking()
            .Queue();
    }

    public void ShowError(string title, string message)
    {
        _toastManager.CreateToast()
            .WithTitle(title)
            .WithContent(message)
            .OfType(NotificationType.Error)
            .Dismiss().After(TimeSpan.FromSeconds(5))
            .Dismiss().ByClicking()
            .Queue();
    }

    public void ShowWarning(string title, string message)
    {
        _toastManager.CreateToast()
            .WithTitle(title)
            .WithContent(message)
            .OfType(NotificationType.Warning)
            .Dismiss().After(TimeSpan.FromSeconds(4))
            .Dismiss().ByClicking()
            .Queue();
    }
}
