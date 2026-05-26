using System.Collections.Generic;
using PdfChopper.Services;

namespace PdfChopper.Tests.Stubs;

public class StubDialogService : IDialogService
{
    public List<(string Type, string Title, string Message)> Calls { get; } = [];

    public void ShowSuccess(string title, string message) =>
        Calls.Add(("Success", title, message));

    public void ShowError(string title, string message) =>
        Calls.Add(("Error", title, message));

    public void ShowWarning(string title, string message) =>
        Calls.Add(("Warning", title, message));
}
