using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using RiggsBurnham_PresentationMaker.ViewModels;

namespace RiggsBurnham_PresentationMaker.Views
{
    public partial class ErrorWindow : Window
    {
        public ErrorWindow(PresentationMakerViewModel parent, string errorTitle, string errorDescription)
        {
            DataContext = new TooManyPicturesErrorViewModel(parent, errorTitle, errorDescription);
            InitializeComponent();
        }
    }
}
