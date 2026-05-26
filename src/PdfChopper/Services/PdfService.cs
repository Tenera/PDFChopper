using System.Collections.Generic;
using PdfChopper.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PdfChopper.Services;

public class PdfService : IPdfService
{
    public int GetPageCount(string filePath)
    {
        using var document = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
        return document.PageCount;
    }

    public void Merge(IReadOnlyList<PdfFile> files, string outputPath)
    {
        var openDocs = new List<PdfDocument>(files.Count);
        try
        {
            using var outputDocument = new PdfDocument();
            foreach (var file in files)
            {
                var inputDocument = PdfReader.Open(file.FilePath, PdfDocumentOpenMode.Import);
                openDocs.Add(inputDocument);
                for (var j = file.StartPage; j <= file.EndPage; j++)
                    outputDocument.AddPage(inputDocument.Pages[j - 1]);
            }
            outputDocument.Save(outputPath);
        }
        finally
        {
            foreach (var doc in openDocs)
                doc.Dispose();
        }
    }

    public void Split(string inputPath, IReadOnlyList<PdfFileExtract> extracts)
    {
        using var inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
        foreach (var extract in extracts)
        {
            using var outputDocument = new PdfDocument();
            for (var j = extract.StartPage; j <= extract.EndPage; j++)
                outputDocument.AddPage(inputDocument.Pages[j - 1]);
            outputDocument.Save(extract.FilePath);
        }
    }

    public void Rotate(string inputPath, IReadOnlyList<PdfFileRotation> parts, string outputPath)
    {
        using var inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
        using var outputDocument = new PdfDocument();
        foreach (var inputDocumentPage in inputDocument.Pages)
            outputDocument.AddPage(inputDocumentPage);
        foreach (var part in parts)
        {
            var rotation = part.Rotate * 90;
            if (rotation == 0) continue;

            for (var j = part.StartPage; j <= part.EndPage; j++)
            {
                var page = outputDocument.Pages[j - 1];
                page.Rotate = (page.Rotate + rotation) % 360;
            }
        }
        outputDocument.Save(outputPath);
    }

    public void Interleave(IReadOnlyList<PdfFile> files, string outputPath)
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
                    q.Enqueue(inputDocument.Pages[j - 1]);
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
            outputDocument.Save(outputPath);
        }
        finally
        {
            foreach (var doc in openDocs)
                doc.Dispose();
        }
    }

    public void Reorder(string inputPath, string outputPath)
    {
        using var inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
        using var outputDocument = new PdfDocument();
        var isEven = inputDocument.PageCount % 2 == 0;
        var middle = isEven ? inputDocument.PageCount / 2 : (inputDocument.PageCount / 2) + 1;
        for (var i = 0; i < middle; i++)
        {
            outputDocument.Pages.Add(inputDocument.Pages[i]);
            if (i < middle - 1 || i == middle - 1 && isEven)
                outputDocument.Pages.Add(inputDocument.Pages[inputDocument.PageCount - i - 1]);
        }
        outputDocument.Save(outputPath);
    }
}
