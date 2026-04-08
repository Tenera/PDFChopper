using System;
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
        PdfDocument? document = null;
        try
        {
            document = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
            return document.PageCount;
        }
        finally
        {
            document?.Close();
            document?.Dispose();
        }
    }

    public static async Task MergeAsync(IReadOnlyList<PdfFile> files, string outputPath)
    {
        var openDocs = new List<PdfDocument>(files.Count);
        PdfDocument? outputDocument = null;
        try
        {
            outputDocument = new PdfDocument();
            foreach (var file in files)
            {
                var inputDocument = PdfReader.Open(file.FilePath, PdfDocumentOpenMode.Import);
                openDocs.Add(inputDocument);
                for (var j = file.StartPage; j <= file.EndPage; j++)
                {
                    outputDocument.AddPage(inputDocument.Pages[j - 1]);
                }
            }
            await outputDocument.SaveAsync(outputPath);
        }
        finally
        {
            outputDocument?.Close();
            outputDocument?.Dispose();
            foreach (var doc in openDocs)
            {
                doc.Close();
                doc.Dispose();
            }
        }
    }

    public static async Task SplitAsync(string inputPath, IReadOnlyList<PdfFileExtract> extracts)
    {
        PdfDocument? inputDocument = null;
        try
        {
            inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
            foreach (var extract in extracts)
            {
                PdfDocument? outputDocument = null;
                try
                {
                    outputDocument = new PdfDocument();
                    for (var j = extract.StartPage; j <= extract.EndPage; j++)
                    {
                        outputDocument.AddPage(inputDocument.Pages[j - 1]);
                    }
                    await outputDocument.SaveAsync(extract.FilePath);
                }
                finally
                {
                    outputDocument?.Close();
                    outputDocument?.Dispose();
                }
            }
        }
        finally
        {
            inputDocument?.Close();
            inputDocument?.Dispose();
        }
    }

    public static async Task RotateAsync(string inputPath, IReadOnlyList<PdfFileRotation> parts, string outputPath)
    {
        PdfDocument? inputDocument = null;
        PdfDocument? outputDocument = null;
        try
        {
            inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
            outputDocument = new PdfDocument();
            foreach (var inputDocumentPage in inputDocument.Pages)
            {
                outputDocument.AddPage(inputDocumentPage);
            }
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
            await outputDocument.SaveAsync(outputPath);
        }
        finally
        {
            outputDocument?.Close();
            outputDocument?.Dispose();
            inputDocument?.Close();
            inputDocument?.Dispose();
        }
    }

    public static async Task InterleaveAsync(IReadOnlyList<PdfFile> files, string outputPath)
    {
        var openDocs = new List<PdfDocument>(files.Count);
        PdfDocument? outputDocument = null;
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

            outputDocument = new PdfDocument();
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
        }
        finally
        {
            outputDocument?.Close();
            outputDocument?.Dispose();
            foreach (var doc in openDocs)
            {
                doc.Close();
                doc.Dispose();
            }
        }
    }

    public static async Task ReorderAsync(string inputPath, string outputPath)
    {
        PdfDocument? inputDocument = null;
        PdfDocument? outputDocument = null;
        try
        {
            inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
            outputDocument = new PdfDocument();
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
        }
        finally
        {
            outputDocument?.Close();
            outputDocument?.Dispose();
            inputDocument?.Close();
            inputDocument?.Dispose();
        }
    }
}
