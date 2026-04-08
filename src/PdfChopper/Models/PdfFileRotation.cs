using CommunityToolkit.Mvvm.ComponentModel;

namespace PdfChopper.Models;

public class PdfFileRotation : ObservableObject
{
    private readonly PdfFile _parent;
    private int _startPage;
    private int _endPage;

    public PdfFileRotation(PdfFile parent)
    {
        _parent = parent;
        _startPage = 1;
        _endPage = _parent.PageCount;
    }

    public int StartPage
    {
        get => _startPage;
        set
        {
            if (_startPage == value) return;
            if (value > _parent.PageCount || value <= 0 || value > _endPage) return;

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
            if (value > _parent.PageCount || value <= 0 || value < _startPage) return;

            _endPage = value;
            OnPropertyChanged();
        }
    }

    public int Rotate
    {
        get;
        set
        {
            var normalized = ((value % 4) + 4) % 4;
            if (field == normalized) return;
            field = normalized;
            OnPropertyChanged();
        }
    }
}
