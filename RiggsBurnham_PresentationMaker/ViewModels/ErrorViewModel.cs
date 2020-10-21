using System.ComponentModel;
using System.Runtime.CompilerServices;
using Prism.Commands;

namespace RiggsBurnham_PresentationMaker.ViewModels
{
    public class TooManyPicturesErrorViewModel : INotifyPropertyChanged
    {
        private PresentationMakerViewModel _parent;
        private string _errorTitle = "";
        private string _errorDescription = "";
        public TooManyPicturesErrorViewModel(PresentationMakerViewModel parent, string errorTitle, string errorDescription)
        {
            _parent = parent;
            ErrorTitle = errorTitle;
            ErrorDescription = errorDescription;
            CloseTooManyPicturesErrorWindowCommand = new DelegateCommand(CloseTooManyPicturesErrorWindow);
        }

        #region property changed
        public event PropertyChangedEventHandler PropertyChanged;
        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion

        public string ErrorTitle
        {
            get => _errorTitle;
            set
            {
                _errorTitle = value;
                NotifyPropertyChanged();
            } 
        }

        public string ErrorDescription
        {
            get => _errorDescription;
            set
            {
                _errorDescription = value;
                NotifyPropertyChanged();
            }
        }

        public DelegateCommand CloseTooManyPicturesErrorWindowCommand { get; set; }

        private void CloseTooManyPicturesErrorWindow()
        {
            if (_parent.TooManyPicturesError != null)
            {
                _parent.TooManyPicturesError.Hide();
                _parent.TooManyPicturesError = null;
            }
            if (_parent.FailedToLoadPictureError != null)
            {
                _parent.FailedToLoadPictureError.Hide();
                _parent.FailedToLoadPictureError = null;
            }
        }
    }
}
