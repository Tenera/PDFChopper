using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using PdfChopper.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using PdfChopper.Services;


namespace PdfChopper.ViewModels;

public partial class MainWindowViewModel : ObservableObject
{
    private static readonly List<FileDialogFilter> PdfFileDialogFilters = [new() { Name = "PDF Files", Extensions = ["pdf"] }];

    // Helper properties and methods
    private static Window? MainWindow => Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop ? desktop.MainWindow : null;

    public MainWindowViewModel()
    {
        // Subscribe to collection changes to update command CanExecute state
        FilesToMerge.CollectionChanged += (_, __) =>
        {
            MergeCommand.NotifyCanExecuteChanged();
            UpCommand.NotifyCanExecuteChanged();
            DownCommand.NotifyCanExecuteChanged();
            DeleteCommand.NotifyCanExecuteChanged();
            ClearCommand.NotifyCanExecuteChanged();
        };

        FileExtracts.CollectionChanged += (_, __) =>
        {
            SplitCommand.NotifyCanExecuteChanged();
            ClearExtractsCommand.NotifyCanExecuteChanged();
            DeleteExtractCommand.NotifyCanExecuteChanged();
            AddExtractCommand.NotifyCanExecuteChanged();
        };

        InterleaveFiles.CollectionChanged += (_, __) =>
        {
            InterleaveCommand.NotifyCanExecuteChanged();
            ClearInterleaveFilesCommand.NotifyCanExecuteChanged();
            DeleteInterleaveFileCommand.NotifyCanExecuteChanged();
            AddInterleaveFileCommand.NotifyCanExecuteChanged();
        };
    }

    #region Merge
    public ObservableCollection<PdfFile> FilesToMerge { get; } = [];

    [ObservableProperty]
    private PdfFile? _selectedPdfFile;

    partial void OnSelectedPdfFileChanged(PdfFile? value)
    {
        UpCommand.NotifyCanExecuteChanged();
        DownCommand.NotifyCanExecuteChanged();
        DeleteCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand]
    public async Task Add()
    {
        var dialog = new OpenFileDialog
        {
            AllowMultiple = true,
            Filters = PdfFileDialogFilters,
            Title = "Select files to merge"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (result is not null)
        {
            AddFiles(result);
        }
    }

    private void AddFiles(string[] files)
    {
        if (files.Length == 0) return;

        foreach (var file in files.Where(x => !string.IsNullOrWhiteSpace(x)))
        {
            var fileInfo = new FileInfo(file);
            if (fileInfo.Exists
                && string.Equals(fileInfo.Extension, ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                FilesToMerge.Add(new PdfFile(file));
            }
        }

        MergeCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand(CanExecute = nameof(CanMerge))]
    public async Task Merge()
    {
        var dialog = new SaveFileDialog
        {
            DefaultExtension = ".pdf",
            Filters = PdfFileDialogFilters,
            Title = "Save merged file to",
            ShowOverwritePrompt = true,
            InitialFileName = "Merged.pdf"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (result is not null)
        {
            await CreateMergedFile(result);
        }
    }

    private async Task CreateMergedFile(string fileName)
    {
        try
        {
            using var outputDocument = new PdfDocument();
            foreach (var file in FilesToMerge)
            {
                using var inputDocument = PdfReader.Open(file.FilePath, PdfDocumentOpenMode.Import);
                for (var j = file.StartPage; j <= file.EndPage; j++)
                {
                    outputDocument.AddPage(inputDocument.Pages[j - 1]);
                }
                inputDocument.Close();
            }
            await outputDocument.SaveAsync(fileName);
            outputDocument.Close();
            await DialogService.ShowMessage("Merge successful", "Files merged successfully");
        }
        catch (Exception ex)
        {
            await DialogService.ShowMessage("Error occurred", ex.Message);
        }
    }

    public bool CanMerge => FilesToMerge is { Count: > 1 };

    [RelayCommand(CanExecute = nameof(CanUp))]
    public void Up()
    {
        if (SelectedPdfFile is null) return;

        var selectedFile = SelectedPdfFile;
        var index = FilesToMerge.IndexOf(selectedFile);
        FilesToMerge.RemoveAt(index);
        FilesToMerge.Insert(index - 1, selectedFile);
        SelectedPdfFile = selectedFile;
    }

    public bool CanUp => SelectedPdfFile is not null
                         && FilesToMerge.Any()
                         && FilesToMerge.IndexOf(SelectedPdfFile) > 0;

    [RelayCommand(CanExecute = nameof(CanDown))]
    public void Down()
    {
        if (SelectedPdfFile is null) return;

        var selectedFile = SelectedPdfFile;
        var index = FilesToMerge.IndexOf(selectedFile);
        FilesToMerge.RemoveAt(index);
        FilesToMerge.Insert(index + 1, selectedFile);
        SelectedPdfFile = selectedFile;
    }

    public bool CanDown => SelectedPdfFile != null
                           && FilesToMerge.Any()
                           && FilesToMerge.IndexOf(SelectedPdfFile) < FilesToMerge.Count - 1;

    [RelayCommand(CanExecute = nameof(CanDelete))]
    public void Delete()
    {
        if (SelectedPdfFile == null) return;

        FilesToMerge.Remove(SelectedPdfFile);

        ClearCommand.NotifyCanExecuteChanged();
        MergeCommand.NotifyCanExecuteChanged();
    }

    public bool CanDelete => SelectedPdfFile is not null;

    [RelayCommand(CanExecute = nameof(CanClear))]
    public void Clear()
    {
        FilesToMerge.Clear();
        MergeCommand.NotifyCanExecuteChanged();
        ClearCommand.NotifyCanExecuteChanged();
    }

    public bool CanClear => FilesToMerge.Any();

    #endregion

    #region Split

    public ObservableCollection<PdfFileExtract> FileExtracts { get; } = [];

    [ObservableProperty]
    private PdfFile? _fileToSplit;

    partial void OnFileToSplitChanged(PdfFile? value)
    {
        SplitCommand.NotifyCanExecuteChanged();
        AddExtractCommand.NotifyCanExecuteChanged();
    }

    [ObservableProperty]
    private PdfFileExtract? _selectedExtract;

    partial void OnSelectedExtractChanged(PdfFileExtract? value)
    {
        DeleteExtractCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand]
    public async Task SelectSplitFile()
    {
        var dialog = new OpenFileDialog
        {
            AllowMultiple = false,
            Filters = PdfFileDialogFilters,
            Title = "Select file to split"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (result is { Length: > 0 })
        {
            await SetSplitFile(result[0]);
        }
    }

    private async Task SetSplitFile(string filepath)
    {
        try
        {
            FileToSplit = new PdfFile(filepath);
        }
        catch (Exception)
        {
            await DialogService.ShowMessage( "Invalid file", "Invalid file specified. Please select a valid PDF-file");
        }
    }

    [RelayCommand(CanExecute = nameof(CanSplit))]
    public async Task Split()
    {
        try
        {
            if (FileToSplit is null || !FileExtracts.Any()) return;

            using var inputDocument = PdfReader.Open(FileToSplit.FilePath, PdfDocumentOpenMode.Import);
            foreach (var extract in FileExtracts)
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
            await DialogService.ShowMessage("Split successful", "File split successfully");
        }
        catch (Exception ex)
        {
            await DialogService.ShowMessage("Error occurred", ex.Message);
        }
    }

    public bool CanSplit => FileToSplit != null && FileExtracts.Any();

    [RelayCommand(CanExecute = nameof(CanAddExtract))]
    public async Task AddExtract()
    {
        var dialog = new SaveFileDialog
        {
            DefaultExtension = ".pdf",
            Filters = PdfFileDialogFilters,
            Title = "Save extract to",
            ShowOverwritePrompt = true,
            InitialFileName = "Extract.pdf"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (string.IsNullOrWhiteSpace(result)) return;

        if (FileExtracts.Any(x => x.FilePath.Equals(result, StringComparison.OrdinalIgnoreCase)))
        {
            await DialogService.ShowMessage("Duplicate extract", "A file with the same path is already in the list of extracts.");
            return;
        }

        FileExtracts.Add(new PdfFileExtract(FileToSplit!, result));
        SplitCommand.NotifyCanExecuteChanged();
        
    }

    public bool CanAddExtract => FileToSplit != null;

    [RelayCommand(CanExecute = nameof(CanDeleteExtract))]
    public void DeleteExtract()
    {
        if (SelectedExtract == null) return;

        FileExtracts.Remove(SelectedExtract);

        SplitCommand.NotifyCanExecuteChanged();
        ClearExtractsCommand.NotifyCanExecuteChanged();
    }

    public bool CanDeleteExtract => SelectedExtract != null;

    [RelayCommand(CanExecute = nameof(CanClearExtracts))]
    public void ClearExtracts()
    {
        FileExtracts.Clear();
        SplitCommand.NotifyCanExecuteChanged();
        ClearExtractsCommand.NotifyCanExecuteChanged();
    }

    public bool CanClearExtracts => FileExtracts.Any();

    #endregion

    #region Interleave

    public ObservableCollection<PdfFile> InterleaveFiles { get; } = [];

    [ObservableProperty]
    private PdfFile? _selectedInterleaveFile;

    partial void OnSelectedInterleaveFileChanged(PdfFile? value)
    {
        DeleteInterleaveFileCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand(CanExecute = nameof(CanInterleave))]
    public async Task Interleave()
    {
        var dialog = new SaveFileDialog
        {
            DefaultExtension = ".pdf",
            Filters = PdfFileDialogFilters,
            Title = "Save interleaved file to",
            ShowOverwritePrompt = true,
            InitialFileName = "Interleaved.pdf"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (string.IsNullOrWhiteSpace(result)) return;

        await CreateInterleavedFile(result);
    }

    private async Task CreateInterleavedFile(string filePath)
    {
        var openDocs = new List<PdfDocument>(InterleaveFiles.Count);
        try
        {
            if (InterleaveFiles.Count <= 1) return;

            // Read all files and enqueue their pages
            var pageQueues = new List<Queue<PdfPage>>();
            foreach (var interleaveFile in InterleaveFiles)
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

            // Dequeue pages in round-robin fashion and add to output document
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

            await outputDocument.SaveAsync(filePath);
            outputDocument.Close();

            await DialogService.ShowMessage("Interleave successful", "Files interleaved successfully");
        }
        catch (Exception ex)
        {
            await DialogService.ShowMessage("Error occurred", ex.Message);
        }
        finally
        {
            // Close and dispose all opened documents
            foreach (var pdfDocument in openDocs)
            {
                pdfDocument.Close();
                pdfDocument.Dispose();
            }
        }
    }

    public bool CanInterleave => InterleaveFiles.Count > 1;

    [RelayCommand(CanExecute = nameof(CanAddInterleaveFile))]
    public async Task AddInterleaveFile()
    {
        var dialog = new OpenFileDialog
        {
            AllowMultiple = false,
            Filters = PdfFileDialogFilters,
            Title = "Add file(s) to interleave"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (result is { Length: > 0 })
        {
            foreach (var file in result)
            {
                InterleaveFiles.Add(new PdfFile(file));
            }
            InterleaveCommand.NotifyCanExecuteChanged();
        }
    }

    public bool CanAddInterleaveFile => true;

    [RelayCommand(CanExecute = nameof(CanDeleteInterleaveFile))]
    public void DeleteInterleaveFile()
    {
        if (SelectedInterleaveFile == null) return;

        InterleaveFiles.Remove(SelectedInterleaveFile);

        InterleaveCommand.NotifyCanExecuteChanged();
        ClearInterleaveFilesCommand.NotifyCanExecuteChanged();
    }

    public bool CanDeleteInterleaveFile => SelectedInterleaveFile != null;

    [RelayCommand(CanExecute = nameof(CanClearInterleaveFiles))]
    public void ClearInterleaveFiles()
    {
        InterleaveFiles.Clear();
        InterleaveCommand.NotifyCanExecuteChanged();
        ClearInterleaveFilesCommand.NotifyCanExecuteChanged();
    }

    public bool CanClearInterleaveFiles => FilesToMerge.Any();

    #endregion

    #region Reorder

    [ObservableProperty]
    private PdfFile? _fileToReorder;

    partial void OnFileToReorderChanged(PdfFile? value)
    {
        OnPropertyChanged(nameof(CanReorder));
        ReorderCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand]
    public async Task SelectReorderFile()
    {
        var dialog = new OpenFileDialog
        {
            AllowMultiple = false,
            Filters = PdfFileDialogFilters,
            Title = "Select file to reorder"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (result is { Length: > 0 })
        {
            await SetReorderFile(result[0]);
        }
    }

    private async Task SetReorderFile(string filepath)
    {
        try
        {
            FileToReorder = new PdfFile(filepath);
        }
        catch (Exception)
        {
            await DialogService.ShowMessage("Invalid file", "Invalid file specified. Please select a valid PDF-file");
        }
    }

    [RelayCommand(CanExecute = nameof(CanReorder))]   
    public async Task Reorder()
    {
        try
        {
            if (FileToReorder == null) return;
            var outputPath = FileToReorder.FilePath.Replace(".pdf", "_2.pdf");

            using var inputDocument = PdfReader.Open(FileToReorder.FilePath, PdfDocumentOpenMode.Import);
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

            await DialogService.ShowMessage("Reordered successful", $"File reordered successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            await DialogService.ShowMessage("Error occurred", ex.Message);
        }
    }

    public bool CanReorder => FileToReorder != null;

    #endregion
}