using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using MergeSplitPdf.Properties;
using PdfSharp.Pdf.IO;

namespace MergeSplitPdf.Model
{
    public class PdfFile : INotifyPropertyChanged
    {
        private int _endPage;
        private int _startPage;

        public PdfFile(string filePath)
        {
            var file = new FileInfo(filePath);
            using (var inputDocument = PdfReader.Open(filePath, PdfDocumentOpenMode.Import))
            {
                PageCount = inputDocument.PageCount;
                FileName = file.Name;
                FilePath = filePath;
                _startPage = 1;
                _endPage = PageCount;
            }
        }

        public string FilePath { get; }

        public string FileName { get; }

        public int PageCount { get; } 

        public int StartPage
        {
            get => _startPage;
            set
            {
                if (_startPage == value) return;
                if (value > PageCount || value <= 0 || value > _endPage) return;

                _startPage = value;
                OnPropertyChanged();
            }
        }

        public int EndPage
        {
            get => _endPage;
            set
            {
                if (_endPage == value) return;
                if (value > PageCount || value <= 0 || value < _startPage) return;

                _endPage = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
