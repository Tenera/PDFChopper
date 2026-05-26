namespace PdfChopper.Models;

public class PdfFileRotation : PageRangeModel
{
    public PdfFileRotation(int pageCount) : base(pageCount) { }

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
