using System.IO;

namespace PdfChopper.Models;

public class PdfFile : PageRangeModel
{
    public PdfFile(string filePath, int pageCount) : base(pageCount)
    {
        FileName = Path.GetFileName(filePath);
        FilePath = filePath;
    }

    public string FilePath { get; }

    public string FileName { get; }
}
