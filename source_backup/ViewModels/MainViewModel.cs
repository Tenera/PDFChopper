using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using Caliburn.Micro;
using MergeSplitPdf.Model;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace MergeSplitPdf.ViewModels
{
    public class MainViewModel : Screen
    {
        #region Merge
        public ObservableCollection<PdfFile> FilesToMerge { get; } = new ObservableCollection<PdfFile>();

        private PdfFile _selectedPdfFile;

        public PdfFile SelectedPdfFile
        {
            get => _selectedPdfFile;
            set
            {
                if (Equals(value, _selectedPdfFile)) return;
                _selectedPdfFile = value;
                NotifyOfPropertyChange();
                NotifyOfPropertyChange(() => CanDown);
                NotifyOfPropertyChange(() => CanUp);
            }
        }

        public void DropMergeFile(DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

            // Note that you can have more than one file.
            AddFiles((string[])e.Data.GetData(DataFormats.FileDrop));
        }

        public void Add()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                CheckFileExists = true,
                Multiselect = true,
                DefaultExt = ".pdf",
                Filter = "PDF Files (*.pdf)|*.pdf",
                Title = "Select files to merge"
            };
            if (dialog.ShowDialog() == true)
            {
                AddFiles(dialog.FileNames);
            }
        }

        private void AddFiles(string[] files)
        {
            if (files == null || files.Length == 0) return;

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
            NotifyOfPropertyChange(() => CanMerge);
        }

        public void Merge()
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                DefaultExt = ".pdf",
                Filter = "PDF Files (*.pdf)|*.pdf",
                Title = "Save merged file to"
            };
            if (dialog.ShowDialog() == true)
            {
                CreateMergedFile(dialog.FileName);
            }
        }

        private void CreateMergedFile(string dialogFileName)
        {
            try
            {
                using (var outputDocument = new PdfDocument())
                {
                    foreach (var file in FilesToMerge)
                    {
                        using (var inputDocument = PdfReader.Open(file.FilePath, PdfDocumentOpenMode.Import))
                        {
                            for (var j = file.StartPage; j <= file.EndPage; j++)
                            {
                                outputDocument.AddPage(inputDocument.Pages[j - 1]);
                            }
                            inputDocument.Close();
                        }
                    }
                    outputDocument.Save(dialogFileName);
                    outputDocument.Close();
                    MessageBox.Show("Files merged successfully", "Merge successful", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error occurred", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        public bool CanMerge => FilesToMerge != null && FilesToMerge.Count > 1;

        public void Up()
        {
            var selectedFile = SelectedPdfFile;
            var index = FilesToMerge.IndexOf(selectedFile);
            FilesToMerge.RemoveAt(index);
            FilesToMerge.Insert(index-1, selectedFile);
            SelectedPdfFile = selectedFile;
        }

        public bool CanUp => SelectedPdfFile != null 
                             && FilesToMerge != null 
                             && FilesToMerge.Any() 
                             && FilesToMerge.IndexOf(SelectedPdfFile) > 0;
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

        private PdfFile _fileToSplit;

        public PdfFile FileToSplit
        {
            get => _fileToSplit;
            set
            {
                if (Equals(value, _fileToSplit)) return;
                _fileToSplit = value;
                FileExtracts.Clear();
                NotifyOfPropertyChange();
                NotifyOfPropertyChange(nameof(CanAddExtract));
                NotifyOfPropertyChange(nameof(CanSplit));
            }
        }

        public void SelectSplitFile()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                CheckFileExists = true,
                Multiselect = false,
                DefaultExt = ".pdf",
                Filter = "PDF Files (*.pdf)|*.pdf",
                Title = "Select file to split"
            };
            if (dialog.ShowDialog() == true)
            {
                SetSplitFile(dialog.FileName);               
            }
        }

        public void DropSplitFile(DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

            var files = (string[]) e.Data.GetData(DataFormats.FileDrop);
            if (files == null || files.Length == 0) return;
            if (FileToSplit?.FilePath == files[0]) return;

            SetSplitFile(files[0]);
        }

        private void SetSplitFile(string filepath)
        {
            try
            {
                FileToSplit = new PdfFile(filepath);
            }
            catch (Exception)
            {
                MessageBox.Show("Invalid file specified. Please select a valid PDF-file", "Invalid file", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        public void Split()
        {
            try
            {
                if (FileToSplit == null || !FileExtracts.Any()) return;

                using (var inputDocument = PdfReader.Open(FileToSplit.FilePath, PdfDocumentOpenMode.Import))
                {
                    foreach (var extract in FileExtracts)
                    {
                        using (var outputDocument = new PdfDocument())
                        {
                            for (var j = extract.StartPage; j <= extract.EndPage; j++)
                            {
                                outputDocument.AddPage(inputDocument.Pages[j - 1]);
                            }
                            outputDocument.Save(extract.FilePath);
                            outputDocument.Close();
                        }
                    }
                    inputDocument.Close();
                }
                MessageBox.Show("File split successfully", "Split successful", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error occurred", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        public bool CanSplit => FileToSplit != null && FileExtracts.Any();

        public void AddExtract()
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                DefaultExt = ".pdf",
                Filter = "PDF Files (*.pdf)|*.pdf",
                Title = "Save extract to"
            };
            if (dialog.ShowDialog() == true)
            {
                FileExtracts.Add(new PdfFileExtract(FileToSplit, dialog.FileName));
                NotifyOfPropertyChange(nameof(CanSplit));
            }
        }

        public bool CanAddExtract => FileToSplit != null;

        public void DeleteExtract()
        {
            if (SelectedExtract == null) return;

            FileExtracts.Remove(SelectedExtract);

            NotifyOfPropertyChange(nameof(CanSplit));
        }

        public bool CanDeleteExtract => SelectedExtract != null;

        private PdfFileExtract _selectedExtract;

        public PdfFileExtract SelectedExtract
        {
            get => _selectedExtract;
            set
            {
                if (Equals(value, _selectedExtract)) return;
                _selectedExtract = value;
                NotifyOfPropertyChange();
                NotifyOfPropertyChange(nameof(CanDeleteExtract));
            }
        }
        #endregion

        #region Interleave

        public ObservableCollection<PdfFile> InterleaveFiles { get; } = new ObservableCollection<PdfFile>();

        public void Interleave()
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                DefaultExt = ".pdf",
                Filter = "PDF Files (*.pdf)|*.pdf",
                Title = "Save interleaved file to"
            };
            if (dialog.ShowDialog() == true)
            {
                CreateInterleavedFile(dialog.FileName);
            }
        }
        public void CreateInterleavedFile(string filePath)
        {
            try
            {
                if (InterleaveFiles.Count <= 1) return;

                var pageQueues = new List<Queue<PdfPage>>();
                foreach (var interleaveFile in InterleaveFiles)
                {
                    var q = new Queue<PdfPage>();
                    using (var inputDocument = PdfReader.Open(interleaveFile.FilePath, PdfDocumentOpenMode.Import))
                    {
                        for (var j = interleaveFile.StartPage; j <= interleaveFile.EndPage; j++)
                        {
                            q.Enqueue(inputDocument.Pages[j - 1]);
                        }
                    }
                    pageQueues.Add(q);
                }

                using (var outputDocument = new PdfDocument())
                {
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

                    outputDocument.Save(filePath);
                    outputDocument.Close();
                }
                
                MessageBox.Show("Files interleaved successfully", "Interleave successful", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error occurred", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        public bool CanInterleave => InterleaveFiles?.Count > 1;

        public void AddInterleaveFile()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".pdf",
                Filter = "PDF Files (*.pdf)|*.pdf",
                Title = "Add file to interleave"
            };
            if (dialog.ShowDialog() == true)
            {
                InterleaveFiles.Add(new PdfFile(dialog.FileName));
                NotifyOfPropertyChange(nameof(CanInterleave));
            }
        }

        public bool CanAddInterleaveFile => true;

        public void DeleteInterleaveFile()
        {
            if (SelectedInterleaveFile == null) return;

            InterleaveFiles.Remove(SelectedInterleaveFile);

            NotifyOfPropertyChange(nameof(CanInterleave));
        }

        public bool CanDeleteInterleaveFile => SelectedInterleaveFile != null;

        private PdfFile _selectedInterleaveFile;

        public PdfFile SelectedInterleaveFile
        {
            get => _selectedInterleaveFile;
            set
            {
                if (Equals(value, _selectedInterleaveFile)) return;
                _selectedInterleaveFile = value;
                NotifyOfPropertyChange();
                NotifyOfPropertyChange(nameof(CanDeleteInterleaveFile));
            }
        }
        #endregion

        #region Reorder

        private PdfFile _fileToReorder;

        public PdfFile FileToReorder
        {
            get => _fileToReorder;
            set
            {
                if (Equals(value, _fileToReorder)) return;
                _fileToReorder = value;
                NotifyOfPropertyChange();
                NotifyOfPropertyChange(nameof(CanReorder));
            }
        }

        public void SelectReorderFile()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                CheckFileExists = true,
                Multiselect = false,
                DefaultExt = ".pdf",
                Filter = "PDF Files (*.pdf)|*.pdf",
                Title = "Select file to reorder"
            };
            if (dialog.ShowDialog() == true)
            {
                SetReorderFile(dialog.FileName);
            }
        }

        public void DropReorderFile(DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files == null || files.Length == 0) return;
            if (FileToReorder?.FilePath == files[0]) return;

            SetReorderFile(files[0]);
        }

        private void SetReorderFile(string filepath)
        {
            try
            {
                FileToReorder = new PdfFile(filepath);
            }
            catch (Exception)
            {
                MessageBox.Show("Invalid file specified. Please select a valid PDF-file", "Invalid file", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        public void Reorder()
        {
            try
            {
                if (FileToReorder == null) return;
                var outputPath = FileToReorder.FilePath.Replace(".pdf", "_2.pdf");

                using (var inputDocument = PdfReader.Open(FileToReorder.FilePath, PdfDocumentOpenMode.Import))
                using (var outputDocument = new PdfDocument())
                {
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
                    outputDocument.Save(outputPath);
                    outputDocument.Close();
                }
                
                MessageBox.Show($"File reordered successfully to {outputPath}", "Reordered successful", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error occurred", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        public bool CanReorder => FileToReorder != null;

        #endregion
    }
}
