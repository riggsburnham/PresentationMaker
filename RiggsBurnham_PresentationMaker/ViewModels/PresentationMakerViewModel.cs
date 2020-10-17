using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net.Mime;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using Prism.Commands;
using System.Drawing;
using System.Net.Http;
using System.Windows.Shapes;
using Shape = Microsoft.Office.Core.Shape;


//using Presentation = Syncfusion.Presentation.Presentation;

namespace RiggsBurnham_PresentationMaker.ViewModels
{
    public class PresentationMakerViewModel : INotifyPropertyChanged
    {
        #region private member variables

        private string _title = "";
        private string _description = "";
        #endregion

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

        #region constructor

        public PresentationMakerViewModel()
        {
            SavePowerpointCommand = new DelegateCommand(SavePowerpoint);
        }
        #endregion

        #region properties

        public string Title
        {
            get => _title;
            set
            {
                _title = value;
                NotifyPropertyChanged();
            }
        }

        public string Description
        {
            get => _description;
            set
            {
                _description = value;
                NotifyPropertyChanged();
            }
        }
        #endregion

        #region commands
        public DelegateCommand SavePowerpointCommand { get; set; }

        private void SavePowerpoint()
        {
            //Application pptApplication = new Application();
            //Slides slides;
            //Slide slide;
            
            //TextRange objText;
            
            //TextRange description;
            string imageTempPath = @"C:\Users\riggs\Pictures\Asuna Closeup.JPG";


            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Powerpoint files (*.pptx)|*.pptx|All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == true)
            {
                // create ppt file
                //Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
                //CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];
                

                //// create new slide
                //slides = pptPresentation.Slides;
                //slide = slides.AddSlide(1, customLayout);

                //// add title
                //objText = slide.Shapes[1].TextFrame.TextRange;
                //objText.Text = Title;
                //objText.Font.Name = "Arial";
                //objText.Font.Size = 32;

                //// add description
                //objText = slide.Shapes[2].TextFrame.TextRange;
                //objText.Text = Description;

                //// add images
                
                //Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[2];
                //slide.Shapes.AddPicture(imageTempPath, 
                //    Microsoft.Office.Core.MsoTriState.msoFalse, 
                //    Microsoft.Office.Core.MsoTriState.msoTrue, 
                //    shape.Left + shape.Width / 2, 
                //    shape.Top + shape.Height / 2, 
                //    shape.Width/2, 
                //    shape.Height/2);


                Application pptApplication = new Application();

                Slides slides;
                Slide slide;
                TextRange objText;

                Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
                CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];
                
                slides = pptPresentation.Slides;
                slide = slides.AddSlide(1, customLayout);

                // give title
                objText = slide.Shapes[1].TextFrame.TextRange;
                objText.Text = Title;
                objText.Font.Name = "Arial";
                objText.Font.Size = 32;

                // description...
                objText = slide.Shapes[2].TextFrame.TextRange;
                objText.Text = Description;


                Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[2];

                //Add first image
                slide.Shapes.AddPicture(imageTempPath, 
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    50, 100, 100, 100);

                //Add second image
                slide.Shapes.AddPicture(imageTempPath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 
                    200, 100, 100, 100);

                //Add third image
                slide.Shapes.AddPicture(imageTempPath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, 
                    350, 100, 100, 100);


                //Microsoft.Office.Interop.PowerPoint.Shape shape = slide1.Shapes[2];
                //slide1.Shapes.AddTitle();
                //objText = slide1.Shapes[1].TextFrame.TextRange;
                //objText.Text = "sample title yo!";
                //objText.Font.Name = "Arial";
                //objText.Font.Size = 32;

                //slide1.Shapes.AddTextbox(10,40,shape.Width, ).TextFrame.TextRange.Text = "Sample Description";



                // save new powerpoint
                pptPresentation.SaveAs(saveFileDialog.FileName, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                //ppt.SaveAs(saveFileDialog.FileName, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            }


        }
        #endregion
    }
}
