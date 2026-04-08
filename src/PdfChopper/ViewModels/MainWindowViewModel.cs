using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using PdfChopper.Models;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
using PdfChopper.Services;


namespace PdfChopper.ViewModels;

public partial class MainWindowViewModel : ObservableObject
{
    private static readonly FilePickerFileType PdfFileType = new("PDF Files") { Patterns = ["*.pdf"] };

    private static IStorageProvider? StorageProvider =>
        Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop
            ? desktop.MainWindow?.StorageProvider
            : null;

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
        };

        InterleaveFiles.CollectionChanged += (_, __) =>
        {
            InterleaveCommand.NotifyCanExecuteChanged();
            ClearInterleaveFilesCommand.NotifyCanExecuteChanged();
            DeleteInterleaveFileCommand.NotifyCanExecuteChanged();
        };

        FileParts.CollectionChanged += (_, __) =>
        {
            RotateCommand.NotifyCanExecuteChanged();
            ClearPartsCommand.NotifyCanExecuteChanged();
            DeletePartCommand.NotifyCanExecuteChanged();
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
        if (StorageProvider is not { CanOpen: true } provider) return;

        var files = await provider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select files to merge",
            AllowMultiple = true,
            FileTypeFilter = [PdfFileType]
        });

        if (files.Count > 0)
        {
            await AddFiles(files.Select(f => f.TryGetLocalPath()).Where(p => p is not null).ToArray()!);
        }
    }

    private async Task AddFiles(string[] files)
    {
        if (files.Length == 0) return;

        foreach (var file in files.Where(x => !string.IsNullOrWhiteSpace(x)))
        {
            var fileInfo = new FileInfo(file);
            if (fileInfo.Exists
                && string.Equals(fileInfo.Extension, ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    FilesToMerge.Add(await PdfFile.CreateAsync(file));
                }
                catch (Exception ex)
                {
                    await DialogService.ShowMessage("Could not open file",
                        $"Skipping '{fileInfo.Name}': {ex.Message}");
                }
            }
        }
    }

    [RelayCommand(CanExecute = nameof(CanMerge))]
    public async Task Merge()
    {
        if (StorageProvider is not { CanSave: true } provider) return;

        var file = await provider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Save merged file to",
            DefaultExtension = "pdf",
            FileTypeChoices = [PdfFileType],
            ShowOverwritePrompt = true,
            SuggestedFileName = "Merged.pdf"
        });

        var filePath = file?.TryGetLocalPath();
        if (filePath is not null)
        {
            await CreateMergedFile(filePath);
        }
    }

    private async Task CreateMergedFile(string fileName)
    {
        try
        {
            await PdfService.MergeAsync(FilesToMerge, fileName);
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

    public bool CanDown => SelectedPdfFile is not null
                           && FilesToMerge.Any()
                           && FilesToMerge.IndexOf(SelectedPdfFile) < FilesToMerge.Count - 1;

    [RelayCommand(CanExecute = nameof(CanDelete))]
    public void Delete()
    {
        if (SelectedPdfFile is null) return;
        FilesToMerge.Remove(SelectedPdfFile);
    }

    public bool CanDelete => SelectedPdfFile is not null;

    [RelayCommand(CanExecute = nameof(CanClear))]
    public void Clear()
    {
        FilesToMerge.Clear();
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
        if (StorageProvider is not { CanOpen: true } provider) return;

        var files = await provider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select file to split",
            AllowMultiple = false,
            FileTypeFilter = [PdfFileType]
        });

        if (files.Count > 0)
        {
            var filePath = files[0].TryGetLocalPath();
            if (filePath is not null)
                await SetSplitFile(filePath);
        }
    }

    private async Task SetSplitFile(string filepath)
    {
        try
        {
            FileToSplit = await PdfFile.CreateAsync(filepath);
            ClearExtracts();
        }
        catch (Exception ex)
        {
            await DialogService.ShowMessage("Invalid file", $"Could not open the selected file: {ex.Message}");
        }
    }

    [RelayCommand(CanExecute = nameof(CanSplit))]
    public async Task Split()
    {
        try
        {
            if (FileToSplit is null || !FileExtracts.Any()) return;

            await PdfService.SplitAsync(FileToSplit.FilePath, FileExtracts);
            await DialogService.ShowMessage("Split successful", "File split successfully");
        }
        catch (Exception ex)
        {
            await DialogService.ShowMessage("Error occurred", ex.Message);
        }
    }

    public bool CanSplit => FileToSplit is not null && FileExtracts.Any();

    [RelayCommand(CanExecute = nameof(CanAddExtract))]
    public async Task AddExtract()
    {
        if (StorageProvider is not { CanSave: true } provider) return;

        var file = await provider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Save extract to",
            DefaultExtension = "pdf",
            FileTypeChoices = [PdfFileType],
            ShowOverwritePrompt = true,
            SuggestedFileName = "Extract.pdf"
        });

        var result = file?.TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(result)) return;

        if (FileExtracts.Any(x => x.FilePath.Equals(result, StringComparison.OrdinalIgnoreCase)))
        {
            await DialogService.ShowMessage("Duplicate extract", "A file with the same path is already in the list of extracts.");
            return;
        }

        FileExtracts.Add(new PdfFileExtract(FileToSplit!, result));
    }

    public bool CanAddExtract => FileToSplit is not null;

    [RelayCommand(CanExecute = nameof(CanDeleteExtract))]
    public void DeleteExtract()
    {
        if (SelectedExtract is null) return;
        FileExtracts.Remove(SelectedExtract);
    }

    public bool CanDeleteExtract => SelectedExtract is not null;

    [RelayCommand(CanExecute = nameof(CanClearExtracts))]
    public void ClearExtracts()
    {
        FileExtracts.Clear();
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
        if (StorageProvider is not { CanSave: true } provider) return;

        var file = await provider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Save interleaved file to",
            DefaultExtension = "pdf",
            FileTypeChoices = [PdfFileType],
            ShowOverwritePrompt = true,
            SuggestedFileName = "Interleaved.pdf"
        });

        var filePath = file?.TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(filePath)) return;

        await CreateInterleavedFile(filePath);
    }

    private async Task CreateInterleavedFile(string filePath)
    {
        try
        {
            if (InterleaveFiles.Count <= 1) return;

            await PdfService.InterleaveAsync(InterleaveFiles, filePath);
            await DialogService.ShowMessage("Interleave successful", "Files interleaved successfully");
        }
        catch (Exception ex)
        {
            await DialogService.ShowMessage("Error occurred", ex.Message);
        }
    }

    public bool CanInterleave => InterleaveFiles.Count > 1;

    [RelayCommand]
    public async Task AddInterleaveFile()
    {
        if (StorageProvider is not { CanOpen: true } provider) return;

        var files = await provider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Add files to interleave",
            AllowMultiple = true,
            FileTypeFilter = [PdfFileType]
        });

        if (files.Count > 0)
        {
            foreach (var file in files)
            {
                var filePath = file.TryGetLocalPath();
                if (filePath is null) continue;
                try
                {
                    InterleaveFiles.Add(await PdfFile.CreateAsync(filePath));
                }
                catch (Exception ex)
                {
                    await DialogService.ShowMessage("Could not open file",
                        $"Skipping '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }
        }
    }

    [RelayCommand(CanExecute = nameof(CanDeleteInterleaveFile))]
    public void DeleteInterleaveFile()
    {
        if (SelectedInterleaveFile is null) return;
        InterleaveFiles.Remove(SelectedInterleaveFile);
    }

    public bool CanDeleteInterleaveFile => SelectedInterleaveFile is not null;

    [RelayCommand(CanExecute = nameof(CanClearInterleaveFiles))]
    public void ClearInterleaveFiles()
    {
        InterleaveFiles.Clear();
    }

    public bool CanClearInterleaveFiles => InterleaveFiles.Any();

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
        if (StorageProvider is not { CanOpen: true } provider) return;

        var files = await provider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select file to reorder",
            AllowMultiple = false,
            FileTypeFilter = [PdfFileType]
        });

        if (files.Count > 0)
        {
            var filePath = files[0].TryGetLocalPath();
            if (filePath is not null)
                await SetReorderFile(filePath);
        }
    }

    private async Task SetReorderFile(string filepath)
    {
        try
        {
            FileToReorder = await PdfFile.CreateAsync(filepath);
        }
        catch (Exception ex)
        {
            await DialogService.ShowMessage("Invalid file", $"Could not open the selected file: {ex.Message}");
        }
    }

    [RelayCommand(CanExecute = nameof(CanReorder))]
    public async Task Reorder()
    {
        try
        {
            if (FileToReorder is null) return;
            var outputPath = FileToReorder.FilePath.Replace(".pdf", "_2.pdf", StringComparison.InvariantCultureIgnoreCase);

            await PdfService.ReorderAsync(FileToReorder.FilePath, outputPath);
            await DialogService.ShowMessage("Reordered successful", $"File reordered successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            await DialogService.ShowMessage("Error occurred", ex.Message);
        }
    }

    public bool CanReorder => FileToReorder is not null;

    #endregion

    #region Rotate

    public ObservableCollection<PdfFileRotation> FileParts { get; } = [];

    [ObservableProperty]
    private PdfFile? _fileToRotate;

    partial void OnFileToRotateChanged(PdfFile? value)
    {
        RotateCommand.NotifyCanExecuteChanged();
        AddPartCommand.NotifyCanExecuteChanged();
    }

    [ObservableProperty]
    private PdfFileRotation? _selectedPart;

    partial void OnSelectedPartChanged(PdfFileRotation? value)
    {
        DeletePartCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand]
    public async Task SelectRotateFile()
    {
        if (StorageProvider is not { CanOpen: true } provider) return;

        var files = await provider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select file to rotate pages in",
            AllowMultiple = false,
            FileTypeFilter = [PdfFileType]
        });

        if (files.Count > 0)
        {
            var filePath = files[0].TryGetLocalPath();
            if (filePath is not null)
                await SetRotateFile(filePath);
        }
    }

    private async Task SetRotateFile(string filepath)
    {
        try
        {
            FileToRotate = await PdfFile.CreateAsync(filepath);
            ClearParts();
        }
        catch (Exception ex)
        {
            await DialogService.ShowMessage("Invalid file", $"Could not open the selected file: {ex.Message}");
        }
    }

    [RelayCommand(CanExecute = nameof(CanRotate))]
    public async Task Rotate()
    {
        if (StorageProvider is not { CanSave: true } provider) return;

        var file = await provider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Save rotated file to",
            DefaultExtension = "pdf",
            FileTypeChoices = [PdfFileType],
            ShowOverwritePrompt = true,
            SuggestedFileName = "Rotated.pdf"
        });

        var filePath = file?.TryGetLocalPath();
        if (filePath is not null)
        {
            await CreateRotatedFile(filePath);
        }
    }

    private async Task CreateRotatedFile(string fileName)
    {
        try
        {
            if (FileToRotate is null || !FileParts.Any()) return;

            await PdfService.RotateAsync(FileToRotate.FilePath, FileParts, fileName);
            await DialogService.ShowMessage("Rotate successful", "File pages rotated successfully");
        }
        catch (Exception ex)
        {
            await DialogService.ShowMessage("Error occurred", ex.Message);
        }
    }

    public bool CanRotate => FileToRotate is not null && FileParts.Any();

    [RelayCommand(CanExecute = nameof(CanAddPart))]
    public void AddPart()
    {
        FileParts.Add(new PdfFileRotation(FileToRotate!));
    }

    public bool CanAddPart => FileToRotate is not null;

    [RelayCommand(CanExecute = nameof(CanDeletePart))]
    public void DeletePart()
    {
        if (SelectedPart is null) return;
        FileParts.Remove(SelectedPart);
    }

    public bool CanDeletePart => SelectedPart is not null;

    [RelayCommand(CanExecute = nameof(CanClearParts))]
    public void ClearParts()
    {
        FileParts.Clear();
    }

    public bool CanClearParts => FileParts.Any();

    #endregion
}
