using System.Collections.Generic;
using PdfChopper.Models;

namespace PdfChopper.Services;

public interface IPdfService
{
    int GetPageCount(string filePath);
    void Merge(IReadOnlyList<PdfFile> files, string outputPath);
    void Split(string inputPath, IReadOnlyList<PdfFileExtract> extracts);
    void Rotate(string inputPath, IReadOnlyList<PdfFileRotation> parts, string outputPath);
    void Interleave(IReadOnlyList<PdfFile> files, string outputPath);
    void Reorder(string inputPath, string outputPath);
}
