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


namespace PdfChopper.ViewModels;

public partial class MainWindowViewModel : ObservableObject
{
    #region Merge
    public ObservableCollection<PdfFile> FilesToMerge { get; } = [];

    [ObservableProperty]
    private PdfFile? _selectedPdfFile;

    [RelayCommand]
    public async Task Add()
    {
        var dialog = new OpenFileDialog
        {
            AllowMultiple = true,
            Filters = [new FileDialogFilter { Name = "PDF Files", Extensions = ["pdf"] }],
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

        foreach (var file in files)
        {
            var fileInfo = new FileInfo(file);
            if (fileInfo.Exists
                && fileInfo.Extension.ToLowerInvariant() == ".pdf"
                && FilesToMerge.All(x => x.FilePath != file))
            {
                FilesToMerge.Add(new PdfFile(file));
            }
        }

        OnPropertyChanged(nameof(CanMerge));
        OnPropertyChanged(nameof(FilesToMerge));
    }

    [RelayCommand(CanExecute = nameof(CanMerge))]
    public async Task Merge()
    {
        var dialog = new SaveFileDialog
        {
            DefaultExtension = ".pdf",
            Filters = [new FileDialogFilter { Name = "PDF Files", Extensions = ["pdf"] }],
            Title = "Save merged file to"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (result != null)
        {
            await CreateMergedFile(result);
        }
    }

    private async Task CreateMergedFile(string dialogFileName)
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
            await outputDocument.SaveAsync(dialogFileName);
            outputDocument.Close();
            Console.WriteLine("Files merged successfully");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    public bool CanMerge => FilesToMerge != null && FilesToMerge.Count > 1;

    [RelayCommand(CanExecute = nameof(CanUp))]
    public void Up()
    {
        var selectedFile = SelectedPdfFile;
        var index = FilesToMerge.IndexOf(selectedFile);
        FilesToMerge.RemoveAt(index);
        FilesToMerge.Insert(index - 1, selectedFile);
        SelectedPdfFile = selectedFile;
    }

    public bool CanUp => SelectedPdfFile != null
                         && FilesToMerge != null
                         && FilesToMerge.Any()
                         && FilesToMerge.IndexOf(SelectedPdfFile) > 0;

    [RelayCommand(CanExecute = nameof(CanDown))]
    public void Down()
    {
        var selectedFile = SelectedPdfFile;
        var index = FilesToMerge.IndexOf(selectedFile);
        FilesToMerge.RemoveAt(index);
        FilesToMerge.Insert(index + 1, selectedFile);
        SelectedPdfFile = selectedFile;
    }

    public bool CanDown => SelectedPdfFile != null
                           && FilesToMerge != null
                           && FilesToMerge.Any()
                           && FilesToMerge.IndexOf(SelectedPdfFile) < FilesToMerge.Count - 1;

    #endregion

    #region Split

    public ObservableCollection<PdfFileExtract> FileExtracts { get; } = new ObservableCollection<PdfFileExtract>();

    [ObservableProperty]
    private PdfFile? _fileToSplit;

    public async Task SelectSplitFile()
    {
        var dialog = new OpenFileDialog
        {
            AllowMultiple = false,
            Filters = [new FileDialogFilter { Name = "PDF Files", Extensions = ["pdf"] }],
            Title = "Select file to split"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (result != null && result.Length > 0)
        {
            SetSplitFile(result[0]);
        }
    }

    private void SetSplitFile(string filepath)
    {
        try
        {
            FileToSplit = new PdfFile(filepath);
        }
        catch (Exception)
        {
            Console.WriteLine("Invalid file specified. Please select a valid PDF-file");
        }
    }

    public async Task Split()
    {
        try
        {
            if (FileToSplit == null || !FileExtracts.Any()) return;

            using (var inputDocument = PdfReader.Open(FileToSplit.FilePath, PdfDocumentOpenMode.Import))
            {
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
            }
            Console.WriteLine("File split successfully");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    public bool CanSplit => FileToSplit != null && FileExtracts.Any();

    public async Task AddExtract()
    {
        var dialog = new SaveFileDialog
        {
            DefaultExtension = ".pdf",
            Filters = [new FileDialogFilter { Name = "PDF Files", Extensions = ["pdf"] }],
            Title = "Save extract to"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (result != null)
        {
            FileExtracts.Add(new PdfFileExtract(FileToSplit!, result));
            OnPropertyChanged(nameof(CanSplit));
        }
    }

    public bool CanAddExtract => FileToSplit != null;

    [RelayCommand(CanExecute = nameof(CanDeleteExtract))]
    public void DeleteExtract()
    {
        if (SelectedExtract == null) return;

        FileExtracts.Remove(SelectedExtract);

        OnPropertyChanged(nameof(CanSplit));
    }

    public bool CanDeleteExtract => SelectedExtract != null;

    [ObservableProperty]
    private PdfFileExtract? _selectedExtract;

    #endregion

    #region Interleave

    public ObservableCollection<PdfFile> InterleaveFiles { get; } = new ObservableCollection<PdfFile>();

    public async Task Interleave()
    {
        var dialog = new SaveFileDialog
        {
            DefaultExtension = ".pdf",
            Filters = [new FileDialogFilter { Name = "PDF Files", Extensions = ["pdf"] }],
            Title = "Save interleaved file to"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (result != null)
        {
            await CreateInterleavedFile(result);
        }
    }

    public async Task CreateInterleavedFile(string filePath)
    {
        try
        {
            if (InterleaveFiles.Count <= 1) return;

            var pageQueues = new List<Queue<PdfPage>>();
            foreach (var interleaveFile in InterleaveFiles)
            {
                var q = new Queue<PdfPage>();
                using var inputDocument = PdfReader.Open(interleaveFile.FilePath, PdfDocumentOpenMode.Import);
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

            await outputDocument.SaveAsync(filePath);
            outputDocument.Close();

            Console.WriteLine("Files interleaved successfully");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    public bool CanInterleave => InterleaveFiles?.Count > 1;

    public async Task AddInterleaveFile()
    {
        var dialog = new OpenFileDialog
        {
            AllowMultiple = false,
            Filters = [new FileDialogFilter { Name = "PDF Files", Extensions = ["pdf"] }],
            Title = "Add file to interleave"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (result != null && result.Length > 0)
        {
            InterleaveFiles.Add(new PdfFile(result[0]));
            OnPropertyChanged(nameof(CanInterleave));
        }
    }

    public bool CanAddInterleaveFile => true;

    [RelayCommand(CanExecute = nameof(CanDeleteInterleaveFile))]
    public void DeleteInterleaveFile()
    {
        if (SelectedInterleaveFile == null) return;

        InterleaveFiles.Remove(SelectedInterleaveFile);

        OnPropertyChanged(nameof(CanInterleave));
    }

    public bool CanDeleteInterleaveFile => SelectedInterleaveFile != null;

    [ObservableProperty]
    private PdfFile? _selectedInterleaveFile;

    #endregion

    #region Reorder

    [ObservableProperty]
    private PdfFile? _fileToReorder;

    partial void OnFileToReorderChanged(PdfFile? value)
    {
        OnPropertyChanged(nameof(CanReorder));
    }

    public async Task SelectReorderFile()
    {
        var dialog = new OpenFileDialog
        {
            AllowMultiple = false,
            Filters = [new FileDialogFilter { Name = "PDF Files", Extensions = ["pdf"] }],
            Title = "Select file to reorder"
        };
        var result = await dialog.ShowAsync(MainWindow);
        if (result != null && result.Length > 0)
        {
            SetReorderFile(result[0]);
        }
    }

    private void SetReorderFile(string filepath)
    {
        try
        {
            FileToReorder = new PdfFile(filepath);
        }
        catch (Exception)
        {
            Console.WriteLine("Invalid file specified. Please select a valid PDF-file");
        }
    }

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

            Console.WriteLine($"File reordered successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    public bool CanReorder => FileToReorder != null;

    #endregion

    // Helper properties and methods
    private Window? MainWindow => Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop ? desktop.MainWindow : null;
}