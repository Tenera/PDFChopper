using System.Threading.Tasks;

namespace PdfChopper.Services;

public interface IDialogService
{
    Task ShowMessage(string title, string message);
}
