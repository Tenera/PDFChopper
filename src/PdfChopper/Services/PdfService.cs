using System.Collections.Generic;
using System.Threading.Tasks;
using PdfChopper.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PdfChopper.Services;

public static class PdfService
{
    public static int GetPageCount(string filePath)
    {
        using var document = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
        var count = document.PageCount;
        document.Close();
        return count;
    }

    public static async Task MergeAsync(IReadOnlyList<PdfFile> files, string outputPath)
    {
        using var outputDocument = new PdfDocument();
        foreach (var file in files)
        {
            using var inputDocument = PdfReader.Open(file.FilePath, PdfDocumentOpenMode.Import);
            for (var j = file.StartPage; j <= file.EndPage; j++)
            {
                outputDocument.AddPage(inputDocument.Pages[j - 1]);
            }
            inputDocument.Close();
        }
        await outputDocument.SaveAsync(outputPath);
        outputDocument.Close();
    }

    public static async Task SplitAsync(string inputPath, IReadOnlyList<PdfFileExtract> extracts)
    {
        using var inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
        foreach (var extract in extracts)
        {
            using var outputDocument = new PdfDocument();
            for (var j = extract.StartPage; j <= extract.EndPage; j++)
            {
                outputDocument.AddPage(inputDocument.Pages[j - 1]);
            }
            await outputDocument.SaveAsync(extract.FilePath);
            outputDocument.Close();
        }
        inputDocument.Close();
    }

    public static async Task RotateAsync(string inputPath, IReadOnlyList<PdfFileRotation> parts, string outputPath)
    {
        using var inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
        using var outputDocument = new PdfDocument();
        foreach (var inputDocumentPage in inputDocument.Pages)
        {
            outputDocument.AddPage(inputDocumentPage);
        }
        foreach (var part in parts)
        {
            for (var j = part.StartPage; j <= part.EndPage; j++)
            {
                var rotation = (part.Rotate % 4) * 90;
                if (rotation == 0) continue;

                var page = outputDocument.Pages[j - 1];
                page.Rotate = (page.Rotate + rotation);
            }
        }
        await outputDocument.SaveAsync(outputPath);
        outputDocument.Close();
        inputDocument.Close();
    }

    public static async Task InterleaveAsync(IReadOnlyList<PdfFile> files, string outputPath)
    {
        var openDocs = new List<PdfDocument>(files.Count);
        try
        {
            var pageQueues = new List<Queue<PdfPage>>();
            foreach (var interleaveFile in files)
            {
                var q = new Queue<PdfPage>();
                var inputDocument = PdfReader.Open(interleaveFile.FilePath, PdfDocumentOpenMode.Import);
                openDocs.Add(inputDocument);
                for (var j = interleaveFile.StartPage; j <= interleaveFile.EndPage; j++)
                {
                    q.Enqueue(inputDocument.Pages[j - 1]);
                }
                pageQueues.Add(q);
            }

            using var outputDocument = new PdfDocument();
            var pagesAdded = true;
            while (pagesAdded)
            {
                pagesAdded = false;
                foreach (var pageQueue in pageQueues)
                {
                    if (pageQueue.Count > 0)
                    {
                        outputDocument.Pages.Add(pageQueue.Dequeue());
                        pagesAdded = true;
                    }
                }
            }

            await outputDocument.SaveAsync(outputPath);
            outputDocument.Close();
        }
        finally
        {
            foreach (var pdfDocument in openDocs)
            {
                pdfDocument.Close();
                pdfDocument.Dispose();
            }
        }
    }

    public static async Task ReorderAsync(string inputPath, string outputPath)
    {
        using var inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
        using var outputDocument = new PdfDocument();
        var isEven = inputDocument.PageCount % 2 == 0;
        var middle = isEven ? inputDocument.PageCount / 2 : (inputDocument.PageCount / 2) + 1;
        for (var i = 0; i < middle; i++)
        {
            outputDocument.Pages.Add(inputDocument.Pages[i]);
            if (i < middle - 1 || i == middle - 1 && isEven)
            {
                outputDocument.Pages.Add(inputDocument.Pages[inputDocument.PageCount - i - 1]);
            }
        }
        await outputDocument.SaveAsync(outputPath);
        outputDocument.Close();
        inputDocument.Close();
    }
}
