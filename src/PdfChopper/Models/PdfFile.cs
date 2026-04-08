using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using PdfChopper.Services;

namespace PdfChopper.Models;

public class PdfFile : INotifyPropertyChanged
{
    private int _endPage;
    private int _startPage;

    public PdfFile(string filePath, int pageCount)
    {
        FileName = Path.GetFileName(filePath);
        FilePath = filePath;
        PageCount = pageCount;
        _startPage = 1;
        _endPage = pageCount;
    }

    public static async Task<PdfFile> CreateAsync(string filePath)
    {
        var pageCount = await Task.Run(() => PdfService.GetPageCount(filePath));
        return new PdfFile(filePath, pageCount);
    }

    public string FilePath { get; }

    public string FileName { get; }

    public int PageCount { get; }

    public int StartPage
    {
        get => _startPage;
        set
        {
            if (_startPage == value) return;
            if (value > PageCount || value <= 0 || value > _endPage) return;

            _startPage = value;
            OnPropertyChanged();
        }
    }

    public int EndPage
    {
        get => _endPage;
        set
        {
            if (_endPage == value) return;
            if (value > PageCount || value <= 0 || value < _startPage) return;

            _endPage = value;
            OnPropertyChanged();
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null!)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}