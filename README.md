# PDFChopper

A desktop application to split, merge, rotate, interleave, and reorder PDF documents. Built with Avalonia UI and PDFsharp.

## Features

### Merge

Combine multiple PDF files into a single document. You can reorder files and select specific page ranges from each file before merging.

### Split

Extract one or more page ranges from a PDF into separate output files. Select a source file, then define as many extracts as needed, each with its own page range and output path.

### Rotate

Rotate pages within a PDF by 90, 180, or 270 degrees. Define multiple page ranges with different rotation amounts and save to a new file.

### Interleave

Merge pages from multiple PDFs in round-robin order. The first page of each file is taken, then the second page, and so on. Useful for combining separately scanned odd and even pages.

### Reorder

Fix page order for recto-verso scanned documents. If you scanned odd pages first (1, 3, 5, ...) with an auto feeder, then flipped the stack and scanned even pages in reverse (10, 8, 6, ...), this feature reassembles them in the correct order.

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download)

## Build and Run

```bash
dotnet build src/PdfChopper/PdfChopper.csproj
dotnet run --project src/PdfChopper/PdfChopper.csproj
```

## Publish

The project is configured for AOT compilation and trimming:

```bash
dotnet publish src/PdfChopper/PdfChopper.csproj -c Release
```

## Tech Stack

- [Avalonia UI](https://avaloniaui.net/) 11.3 - Cross-platform .NET UI framework
- [PDFsharp](https://docs.pdfsharp.net/) 6.2 - PDF manipulation library
- [CommunityToolkit.Mvvm](https://learn.microsoft.com/en-us/dotnet/communitytoolkit/mvvm/) - MVVM source generators and helpers

## License

[MIT](LICENSE)
