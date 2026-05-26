using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Headless.XUnit;
using FluentAssertions;
using PdfChopper.Models;
using PdfChopper.Tests.Stubs;
using PdfChopper.ViewModels;
using SukiUI.Toasts;
using Xunit;

namespace PdfChopper.Tests.ViewModels;

public class MainWindowViewModelTests
{
    private static MainWindowViewModel CreateViewModel() =>
        new(new SukiToastManager(), new StubPdfService(), new StubDialogService());

    private static (MainWindowViewModel Vm, StubDialogService Dialog, StubPdfService PdfService) CreateViewModelWithStubs()
    {
        var dialog = new StubDialogService();
        var pdfService = new StubPdfService();
        var vm = new MainWindowViewModel(new SukiToastManager(), pdfService, dialog);
        return (vm, dialog, pdfService);
    }

    #region Merge

    [AvaloniaFact]
    public void MergeCommand_DisabledWhenEmpty()
    {
        var vm = CreateViewModel();

        vm.MergeCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void MergeCommand_DisabledWithOneFile()
    {
        var vm = CreateViewModel();
        vm.FilesToMerge.Add(new PdfFile("a.pdf", 5));

        vm.MergeCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void MergeCommand_EnabledWithTwoFiles()
    {
        var vm = CreateViewModel();
        vm.FilesToMerge.Add(new PdfFile("a.pdf", 5));
        vm.FilesToMerge.Add(new PdfFile("b.pdf", 3));

        vm.MergeCommand.CanExecute(null).Should().BeTrue();
    }

    [AvaloniaFact]
    public void DeleteCommand_DisabledWhenNoSelection()
    {
        var vm = CreateViewModel();
        vm.FilesToMerge.Add(new PdfFile("a.pdf", 5));

        vm.DeleteCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void DeleteCommand_EnabledWhenSelected()
    {
        var vm = CreateViewModel();
        var file = new PdfFile("a.pdf", 5);
        vm.FilesToMerge.Add(file);
        vm.SelectedPdfFile = file;

        vm.DeleteCommand.CanExecute(null).Should().BeTrue();
    }

    [AvaloniaFact]
    public void Delete_RemovesSelectedFile()
    {
        var vm = CreateViewModel();
        var file = new PdfFile("a.pdf", 5);
        vm.FilesToMerge.Add(file);
        vm.SelectedPdfFile = file;

        vm.Delete();

        vm.FilesToMerge.Should().BeEmpty();
    }

    [AvaloniaFact]
    public void ClearCommand_DisabledWhenEmpty()
    {
        var vm = CreateViewModel();

        vm.ClearCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void ClearCommand_EnabledWhenHasItems()
    {
        var vm = CreateViewModel();
        vm.FilesToMerge.Add(new PdfFile("a.pdf", 5));

        vm.ClearCommand.CanExecute(null).Should().BeTrue();
    }

    [AvaloniaFact]
    public void Clear_RemovesAllFiles()
    {
        var vm = CreateViewModel();
        vm.FilesToMerge.Add(new PdfFile("a.pdf", 5));
        vm.FilesToMerge.Add(new PdfFile("b.pdf", 3));

        vm.Clear();

        vm.FilesToMerge.Should().BeEmpty();
    }

    [AvaloniaFact]
    public void UpCommand_DisabledWhenFirstItemSelected()
    {
        var vm = CreateViewModel();
        var file1 = new PdfFile("a.pdf", 5);
        vm.FilesToMerge.Add(file1);
        vm.FilesToMerge.Add(new PdfFile("b.pdf", 3));
        vm.SelectedPdfFile = file1;

        vm.UpCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void UpCommand_EnabledWhenNonFirstItemSelected()
    {
        var vm = CreateViewModel();
        vm.FilesToMerge.Add(new PdfFile("a.pdf", 5));
        var file2 = new PdfFile("b.pdf", 3);
        vm.FilesToMerge.Add(file2);
        vm.SelectedPdfFile = file2;

        vm.UpCommand.CanExecute(null).Should().BeTrue();
    }

    [AvaloniaFact]
    public void Up_MovesSelectedItemUp()
    {
        var vm = CreateViewModel();
        var file1 = new PdfFile("a.pdf", 5);
        var file2 = new PdfFile("b.pdf", 3);
        vm.FilesToMerge.Add(file1);
        vm.FilesToMerge.Add(file2);
        vm.SelectedPdfFile = file2;

        vm.Up();

        vm.FilesToMerge[0].Should().Be(file2);
        vm.FilesToMerge[1].Should().Be(file1);
        vm.SelectedPdfFile.Should().Be(file2);
    }

    [AvaloniaFact]
    public void DownCommand_DisabledWhenLastItemSelected()
    {
        var vm = CreateViewModel();
        vm.FilesToMerge.Add(new PdfFile("a.pdf", 5));
        var file2 = new PdfFile("b.pdf", 3);
        vm.FilesToMerge.Add(file2);
        vm.SelectedPdfFile = file2;

        vm.DownCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void DownCommand_EnabledWhenNonLastItemSelected()
    {
        var vm = CreateViewModel();
        var file1 = new PdfFile("a.pdf", 5);
        vm.FilesToMerge.Add(file1);
        vm.FilesToMerge.Add(new PdfFile("b.pdf", 3));
        vm.SelectedPdfFile = file1;

        vm.DownCommand.CanExecute(null).Should().BeTrue();
    }

    [AvaloniaFact]
    public void Down_MovesSelectedItemDown()
    {
        var vm = CreateViewModel();
        var file1 = new PdfFile("a.pdf", 5);
        var file2 = new PdfFile("b.pdf", 3);
        vm.FilesToMerge.Add(file1);
        vm.FilesToMerge.Add(file2);
        vm.SelectedPdfFile = file1;

        vm.Down();

        vm.FilesToMerge[0].Should().Be(file2);
        vm.FilesToMerge[1].Should().Be(file1);
        vm.SelectedPdfFile.Should().Be(file1);
    }

    [AvaloniaFact]
    public void MergeCommand_DisabledAfterClear()
    {
        var vm = CreateViewModel();
        vm.FilesToMerge.Add(new PdfFile("a.pdf", 5));
        vm.FilesToMerge.Add(new PdfFile("b.pdf", 3));
        vm.MergeCommand.CanExecute(null).Should().BeTrue();

        vm.Clear();

        vm.MergeCommand.CanExecute(null).Should().BeFalse();
    }

    #endregion

    #region Split

    [AvaloniaFact]
    public void SplitCommand_DisabledWhenNoFile()
    {
        var vm = CreateViewModel();

        vm.SplitCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void SplitCommand_DisabledWhenNoExtracts()
    {
        var vm = CreateViewModel();
        vm.FileToSplit = new PdfFile("test.pdf", 5);

        vm.SplitCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void SplitCommand_EnabledWhenFileAndExtracts()
    {
        var vm = CreateViewModel();
        vm.FileToSplit = new PdfFile("test.pdf", 5);
        vm.FileExtracts.Add(new PdfFileExtract(5, "out.pdf"));

        vm.SplitCommand.CanExecute(null).Should().BeTrue();
    }

    [AvaloniaFact]
    public void AddExtractCommand_DisabledWhenNoFile()
    {
        var vm = CreateViewModel();

        vm.AddExtractCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void AddExtractCommand_EnabledWhenFileSelected()
    {
        var vm = CreateViewModel();
        vm.FileToSplit = new PdfFile("test.pdf", 5);

        vm.AddExtractCommand.CanExecute(null).Should().BeTrue();
    }

    [AvaloniaFact]
    public void DeleteExtractCommand_DisabledWhenNoSelection()
    {
        var vm = CreateViewModel();
        vm.FileExtracts.Add(new PdfFileExtract(5, "out.pdf"));

        vm.DeleteExtractCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void ClearExtracts_RemovesAll()
    {
        var vm = CreateViewModel();
        vm.FileExtracts.Add(new PdfFileExtract(5, "a.pdf"));
        vm.FileExtracts.Add(new PdfFileExtract(5, "b.pdf"));

        vm.ClearExtracts();

        vm.FileExtracts.Should().BeEmpty();
    }

    #endregion

    #region Rotate

    [AvaloniaFact]
    public void RotateCommand_DisabledWhenNoFile()
    {
        var vm = CreateViewModel();

        vm.RotateCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void RotateCommand_DisabledWhenNoParts()
    {
        var vm = CreateViewModel();
        vm.FileToRotate = new PdfFile("test.pdf", 5);

        vm.RotateCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void RotateCommand_EnabledWhenFileAndParts()
    {
        var vm = CreateViewModel();
        vm.FileToRotate = new PdfFile("test.pdf", 5);
        vm.FileParts.Add(new PdfFileRotation(5));

        vm.RotateCommand.CanExecute(null).Should().BeTrue();
    }

    [AvaloniaFact]
    public void AddPartCommand_DisabledWhenNoFile()
    {
        var vm = CreateViewModel();

        vm.AddPartCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void AddPartCommand_EnabledWhenFileSelected()
    {
        var vm = CreateViewModel();
        vm.FileToRotate = new PdfFile("test.pdf", 5);

        vm.AddPartCommand.CanExecute(null).Should().BeTrue();
    }

    [AvaloniaFact]
    public void AddPart_CreatesRotationWithCorrectPageCount()
    {
        var vm = CreateViewModel();
        vm.FileToRotate = new PdfFile("test.pdf", 7);

        vm.AddPart();

        vm.FileParts.Should().HaveCount(1);
        vm.FileParts[0].PageCount.Should().Be(7);
        vm.FileParts[0].StartPage.Should().Be(1);
        vm.FileParts[0].EndPage.Should().Be(7);
        vm.FileParts[0].Rotate.Should().Be(0);
    }

    [AvaloniaFact]
    public void ClearParts_RemovesAll()
    {
        var vm = CreateViewModel();
        vm.FileToRotate = new PdfFile("test.pdf", 5);
        vm.AddPart();
        vm.AddPart();

        vm.ClearParts();

        vm.FileParts.Should().BeEmpty();
    }

    #endregion

    #region Interleave

    [AvaloniaFact]
    public void InterleaveCommand_DisabledWithOneFile()
    {
        var vm = CreateViewModel();
        vm.InterleaveFiles.Add(new PdfFile("a.pdf", 5));

        vm.InterleaveCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void InterleaveCommand_EnabledWithTwoFiles()
    {
        var vm = CreateViewModel();
        vm.InterleaveFiles.Add(new PdfFile("a.pdf", 5));
        vm.InterleaveFiles.Add(new PdfFile("b.pdf", 3));

        vm.InterleaveCommand.CanExecute(null).Should().BeTrue();
    }

    [AvaloniaFact]
    public void DeleteInterleaveFileCommand_DisabledWhenNoSelection()
    {
        var vm = CreateViewModel();
        vm.InterleaveFiles.Add(new PdfFile("a.pdf", 5));

        vm.DeleteInterleaveFileCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void ClearInterleaveFiles_RemovesAll()
    {
        var vm = CreateViewModel();
        vm.InterleaveFiles.Add(new PdfFile("a.pdf", 5));
        vm.InterleaveFiles.Add(new PdfFile("b.pdf", 3));

        vm.ClearInterleaveFiles();

        vm.InterleaveFiles.Should().BeEmpty();
    }

    #endregion

    #region Reorder

    [AvaloniaFact]
    public void ReorderCommand_DisabledWhenNoFile()
    {
        var vm = CreateViewModel();

        vm.ReorderCommand.CanExecute(null).Should().BeFalse();
    }

    [AvaloniaFact]
    public void ReorderCommand_EnabledWhenFileSelected()
    {
        var vm = CreateViewModel();
        vm.FileToReorder = new PdfFile("test.pdf", 5);

        vm.ReorderCommand.CanExecute(null).Should().BeTrue();
    }

    #endregion

    #region Theme

    [AvaloniaFact]
    public void IsDarkTheme_DefaultsFalse()
    {
        var vm = CreateViewModel();

        vm.IsDarkTheme.Should().BeFalse();
    }

    #endregion

    #region Dialog - Split

    [AvaloniaFact]
    public async Task Split_ShowsSuccessDialog()
    {
        var (vm, dialog, _) = CreateViewModelWithStubs();
        vm.FileToSplit = new PdfFile("test.pdf", 5);
        vm.FileExtracts.Add(new PdfFileExtract(5, "out.pdf"));

        await vm.Split();

        dialog.Calls.Should().ContainSingle()
            .Which.Should().Be(("Success", "Split successful", "File split successfully"));
    }

    [AvaloniaFact]
    public async Task Split_ShowsErrorDialogOnFailure()
    {
        var (vm, dialog, pdfService) = CreateViewModelWithStubs();
        pdfService.ExceptionToThrow = new InvalidOperationException("corrupt file");
        vm.FileToSplit = new PdfFile("test.pdf", 5);
        vm.FileExtracts.Add(new PdfFileExtract(5, "out.pdf"));

        await vm.Split();

        dialog.Calls.Should().ContainSingle()
            .Which.Should().Be(("Error", "Error occurred", "corrupt file"));
    }

    #endregion

    #region Dialog - Merge

    [AvaloniaFact]
    public async Task CreateMergedFile_ShowsSuccessDialog()
    {
        var (vm, dialog, _) = CreateViewModelWithStubs();
        vm.FilesToMerge.Add(new PdfFile("a.pdf", 5));
        vm.FilesToMerge.Add(new PdfFile("b.pdf", 3));

        await vm.CreateMergedFile("output.pdf");

        dialog.Calls.Should().ContainSingle()
            .Which.Should().Be(("Success", "Merge successful", "Files merged successfully"));
    }

    [AvaloniaFact]
    public async Task CreateMergedFile_ShowsErrorDialogOnFailure()
    {
        var (vm, dialog, pdfService) = CreateViewModelWithStubs();
        pdfService.ExceptionToThrow = new InvalidOperationException("write error");
        vm.FilesToMerge.Add(new PdfFile("a.pdf", 5));
        vm.FilesToMerge.Add(new PdfFile("b.pdf", 3));

        await vm.CreateMergedFile("output.pdf");

        dialog.Calls.Should().ContainSingle()
            .Which.Should().Be(("Error", "Error occurred", "write error"));
    }

    #endregion

    #region Dialog - Rotate

    [AvaloniaFact]
    public async Task CreateRotatedFile_ShowsSuccessDialog()
    {
        var (vm, dialog, _) = CreateViewModelWithStubs();
        vm.FileToRotate = new PdfFile("test.pdf", 5);
        vm.FileParts.Add(new PdfFileRotation(5));

        await vm.CreateRotatedFile("output.pdf");

        dialog.Calls.Should().ContainSingle()
            .Which.Should().Be(("Success", "Rotate successful", "File pages rotated successfully"));
    }

    [AvaloniaFact]
    public async Task CreateRotatedFile_ShowsErrorDialogOnFailure()
    {
        var (vm, dialog, pdfService) = CreateViewModelWithStubs();
        pdfService.ExceptionToThrow = new InvalidOperationException("rotation failed");
        vm.FileToRotate = new PdfFile("test.pdf", 5);
        vm.FileParts.Add(new PdfFileRotation(5));

        await vm.CreateRotatedFile("output.pdf");

        dialog.Calls.Should().ContainSingle()
            .Which.Should().Be(("Error", "Error occurred", "rotation failed"));
    }

    #endregion

    #region Dialog - Interleave

    [AvaloniaFact]
    public async Task CreateInterleavedFile_ShowsSuccessDialog()
    {
        var (vm, dialog, _) = CreateViewModelWithStubs();
        vm.InterleaveFiles.Add(new PdfFile("a.pdf", 5));
        vm.InterleaveFiles.Add(new PdfFile("b.pdf", 3));

        await vm.CreateInterleavedFile("output.pdf");

        dialog.Calls.Should().ContainSingle()
            .Which.Should().Be(("Success", "Interleave successful", "Files interleaved successfully"));
    }

    [AvaloniaFact]
    public async Task CreateInterleavedFile_ShowsErrorDialogOnFailure()
    {
        var (vm, dialog, pdfService) = CreateViewModelWithStubs();
        pdfService.ExceptionToThrow = new InvalidOperationException("interleave failed");
        vm.InterleaveFiles.Add(new PdfFile("a.pdf", 5));
        vm.InterleaveFiles.Add(new PdfFile("b.pdf", 3));

        await vm.CreateInterleavedFile("output.pdf");

        dialog.Calls.Should().ContainSingle()
            .Which.Should().Be(("Error", "Error occurred", "interleave failed"));
    }

    #endregion

    #region Dialog - SetFile errors

    [AvaloniaFact]
    public async Task SetSplitFile_ShowsErrorForInvalidFile()
    {
        var (vm, dialog, pdfService) = CreateViewModelWithStubs();
        pdfService.ExceptionToThrow = new InvalidOperationException("bad pdf");

        await vm.SetSplitFile("bad.pdf");

        vm.FileToSplit.Should().BeNull();
        dialog.Calls.Should().ContainSingle()
            .Which.Type.Should().Be("Error");
    }

    [AvaloniaFact]
    public async Task SetSplitFile_SetsFileOnSuccess()
    {
        var (vm, dialog, _) = CreateViewModelWithStubs();

        await vm.SetSplitFile("good.pdf");

        vm.FileToSplit.Should().NotBeNull();
        vm.FileToSplit!.FilePath.Should().Be("good.pdf");
        vm.FileToSplit.PageCount.Should().Be(5);
        dialog.Calls.Should().BeEmpty();
    }

    [AvaloniaFact]
    public async Task SetRotateFile_ShowsErrorForInvalidFile()
    {
        var (vm, dialog, pdfService) = CreateViewModelWithStubs();
        pdfService.ExceptionToThrow = new InvalidOperationException("bad pdf");

        await vm.SetRotateFile("bad.pdf");

        vm.FileToRotate.Should().BeNull();
        dialog.Calls.Should().ContainSingle()
            .Which.Type.Should().Be("Error");
    }

    [AvaloniaFact]
    public async Task SetRotateFile_SetsFileOnSuccess()
    {
        var (vm, dialog, _) = CreateViewModelWithStubs();

        await vm.SetRotateFile("good.pdf");

        vm.FileToRotate.Should().NotBeNull();
        vm.FileToRotate!.FilePath.Should().Be("good.pdf");
        dialog.Calls.Should().BeEmpty();
    }

    [AvaloniaFact]
    public async Task SetReorderFile_ShowsErrorForInvalidFile()
    {
        var (vm, dialog, pdfService) = CreateViewModelWithStubs();
        pdfService.ExceptionToThrow = new InvalidOperationException("bad pdf");

        await vm.SetReorderFile("bad.pdf");

        vm.FileToReorder.Should().BeNull();
        dialog.Calls.Should().ContainSingle()
            .Which.Type.Should().Be("Error");
    }

    [AvaloniaFact]
    public async Task SetReorderFile_SetsFileOnSuccess()
    {
        var (vm, dialog, _) = CreateViewModelWithStubs();

        await vm.SetReorderFile("good.pdf");

        vm.FileToReorder.Should().NotBeNull();
        vm.FileToReorder!.FilePath.Should().Be("good.pdf");
        dialog.Calls.Should().BeEmpty();
    }

    #endregion
}
