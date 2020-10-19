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
    /// <summary>
    /// Interaction logic for TooManyPicturesErrorWindow.xaml
    /// </summary>
    public partial class TooManyPicturesErrorWindow : Window
    {
        public TooManyPicturesErrorWindow(PresentationMakerViewModel parent)
        {
            DataContext = new TooManyPicturesErrorViewModel(parent);
            InitializeComponent();
        }
    }
}
