namespace PdfChopper.Models;

public class PdfFileExtract : PageRangeModel
{
    private string _filePath;

    public PdfFileExtract(int pageCount, string filePath) : base(pageCount)
    {
        _filePath = filePath;
    }

    public string FilePath
    {
        get => _filePath;
        set
        {
            if (value == _filePath) return;
            _filePath = value;
            OnPropertyChanged();
        }
    }
}
