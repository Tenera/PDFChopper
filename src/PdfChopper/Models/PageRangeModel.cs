using CommunityToolkit.Mvvm.ComponentModel;

namespace PdfChopper.Models;

public abstract class PageRangeModel : ObservableObject
{
    private int _startPage;
    private int _endPage;

    protected PageRangeModel(int pageCount)
    {
        PageCount = pageCount;
        _startPage = 1;
        _endPage = pageCount;
    }

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
}
