using System.ComponentModel;
using System.Runtime.CompilerServices;
using MergeSplitPdf.Properties;

namespace MergeSplitPdf.Model
{
    public class PdfFileExtract : INotifyPropertyChanged
    {
        private readonly PdfFile _parent;
        private int _startPage;
        private int _endPage;
        private string _filePath;

        public PdfFileExtract(PdfFile parent, string filePath)
        {
            _parent = parent;
            FilePath = filePath;
            _startPage = 1;
            _endPage = _parent.PageCount;
        }

        public string FilePath
        {
            get => _filePath;
            set
            {
                if (value == _filePath) return;
                _filePath = value;
                OnPropertyChanged();
            }
        }

        public int StartPage
        {
            get => _startPage;
            set
            {
                if (_startPage == value) return;
                if (value > _parent.PageCount || value <= 0 || value > _endPage) return;

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
                if (value > _parent.PageCount || value <= 0 || value < _startPage) return;

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