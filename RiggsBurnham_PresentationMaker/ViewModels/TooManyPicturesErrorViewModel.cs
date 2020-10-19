using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Prism.Commands;

namespace RiggsBurnham_PresentationMaker.ViewModels
{
    public class TooManyPicturesErrorViewModel
    {
        private PresentationMakerViewModel _parent;
        public TooManyPicturesErrorViewModel(PresentationMakerViewModel parent)
        {
            _parent = parent;
            CloseTooManyPicturesErrorWindowCommand = new DelegateCommand(CloseTooManyPicturesErrorWindow);
        }

        public DelegateCommand CloseTooManyPicturesErrorWindowCommand { get; set; }

        private void CloseTooManyPicturesErrorWindow()
        {
            _parent.TooManyPicturesError.Hide();
            _parent.TooManyPicturesError = null;
        }
    }
}
