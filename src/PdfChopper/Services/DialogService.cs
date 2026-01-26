using System.Threading.Tasks;
using Avalonia.Controls;
using MsBox.Avalonia;
using MsBox.Avalonia.Dto;
using MsBox.Avalonia.Enums;

namespace PdfChopper.Services;

public static class DialogService
{
    public static async Task ShowMessage(string title, string message)
    {
        var msg = MessageBoxManager.GetMessageBoxStandard(new MessageBoxStandardParams
        {
            ButtonDefinitions = ButtonEnum.Ok,
            CanResize = false,
            ContentTitle = title,
            ContentMessage = message,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Icon = Icon.None
        });

        await msg.ShowAsync();
    }
}
