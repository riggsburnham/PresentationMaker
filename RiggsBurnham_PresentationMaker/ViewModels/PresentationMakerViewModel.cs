using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
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
using RiggsBurnham_PresentationMaker.Models;
using Image = System.Windows.Controls.Image;


namespace RiggsBurnham_PresentationMaker.ViewModels
{
    public class PresentationMakerViewModel : INotifyPropertyChanged
    {
        #region constants
        // these constants are for setting up the display area on the slide for where all pictures will go
        private const float PICTURE_BOX_TOP = 28.75f;
        private const float PICTURE_BOX_LEFT = 480f;
        private const float PICTURE_BOX_WIDTH = 414f;
        private const float PICTURE_BOX_HEIGHT = 457f;
        private const float PICTURE_BUFFER = 10f;
        #endregion

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
            // TODO: remove this once you are pulling images from google...
            string imageTempPath = @"C:\Users\riggs\Pictures\Asuna Closeup.JPG";

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Powerpoint files (*.pptx)|*.pptx|All files (*.*)|*.*";

            if (saveFileDialog.ShowDialog() == true)
            {
                Application pptApplication = new Application();
                Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
                CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];
                Slides slides = pptPresentation.Slides;
                Slide slide = slides.AddSlide(1, customLayout);

                // give title
                TextRange objText = slide.Shapes[1].TextFrame.TextRange;

                // modifying the title to only take up left half of page with a buffer gap of 10.
                slide.Shapes[1].Width /= 2;
                slide.Shapes[1].Width -= 10;
                objText.Text = Title;
                objText.Font.Name = "Arial";
                objText.Font.Size = 32;

                // description...
                objText = slide.Shapes[2].TextFrame.TextRange;

                // modifying the description to only take up left half of page with a buffer gap of 10.
                slide.Shapes[2].Width /= 2;
                slide.Shapes[2].Width -= 10;
                objText.Text = Description;
                objText.Font.Name = "Arial";
                objText.Font.Size = 16;

                // Add Images
                // TODO: this will be populated with selected images in the future
                List<string> imagePaths = new List<string>();
                imagePaths.Add(imageTempPath);
                imagePaths.Add(imageTempPath);
                imagePaths.Add(imageTempPath);
                imagePaths.Add(imageTempPath);

                // will most likely limit number of images on a slide to 4...
                PictureDimensions dimens = new PictureDimensions();
                System.Drawing.Image img;
                float runningHeight = 0;
                float runningWidth = 0;
                switch (imagePaths.Count)
                {
                    case 1:
                        // resize picture to fit inside box while retaining same aspect ratio...
                        img = System.Drawing.Image.FromFile(imageTempPath);
                        dimens = ResizePicture(img.Width, img.Height, 1);
                        slide.Shapes.AddPicture(
                            imageTempPath, 
                            Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 
                            PICTURE_BOX_LEFT, 
                            PICTURE_BOX_TOP, 
                            Convert.ToSingle(dimens.Width), 
                            Convert.ToSingle(dimens.Height)
                            );
                        break;
                    case 2:
                        runningHeight = 0;
                        for (int i = 0; i < 2; ++i)
                        {
                            img = System.Drawing.Image.FromFile(imagePaths[i]);
                            dimens = ResizePicture(img.Width, img.Height, 2);
                            switch (i)
                            {
                                case 0:
                                    // first picure will be placed above the second, will use returned values without modifying them
                                    slide.Shapes.AddPicture(
                                        imageTempPath, 
                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 
                                        PICTURE_BOX_LEFT, PICTURE_BOX_TOP, 
                                        Convert.ToSingle(dimens.Width), 
                                        Convert.ToSingle(dimens.Height)
                                        );
                                    runningHeight = Convert.ToSingle(dimens.Height);
                                    break;
                                case 1:
                                    // second picture will be placed below the first, will need to modify the top by the height the the picture before...
                                    slide.Shapes.AddPicture(
                                        imageTempPath, 
                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 
                                        PICTURE_BOX_LEFT, PICTURE_BOX_TOP + runningHeight + PICTURE_BUFFER, 
                                        Convert.ToSingle(dimens.Width), 
                                        Convert.ToSingle(dimens.Height)
                                        );
                                    break;
                            }
                            
                        }
                        break;
                    case 3:
                        runningHeight = 0;
                        runningWidth = 0;
                        for (int i = 0; i < 3; ++i)
                        {
                            img = System.Drawing.Image.FromFile(imagePaths[i]);
                            dimens = ResizePicture(img.Width, img.Height, 2);
                            switch (i)
                            {
                                case 0:
                                    // first picure will be placed above the second, will use returned values without modifying them
                                    slide.Shapes.AddPicture(
                                        imageTempPath, 
                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 
                                        PICTURE_BOX_LEFT, 
                                        PICTURE_BOX_TOP, 
                                        Convert.ToSingle(dimens.Width), 
                                        Convert.ToSingle(dimens.Height)
                                        );
                                    runningHeight = Convert.ToSingle(dimens.Height);
                                    break;
                                case 1:
                                    // second picture will be placed below the first, will need to modify the top by the height the the picture before...
                                    slide.Shapes.AddPicture(
                                        imageTempPath,
                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 
                                        PICTURE_BOX_LEFT, PICTURE_BOX_TOP + runningHeight + PICTURE_BUFFER, 
                                        Convert.ToSingle(dimens.Width), 
                                        Convert.ToSingle(dimens.Height)
                                        );
                                    runningWidth = Convert.ToSingle(dimens.Width);
                                    break;
                                case 2:
                                    // third picture will be placed to the right of the second, will need to modify both top and left by the height and width of previous pic
                                    slide.Shapes.AddPicture(
                                        imageTempPath, 
                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 
                                        PICTURE_BOX_LEFT + runningWidth + PICTURE_BUFFER, 
                                        PICTURE_BOX_TOP + runningHeight + PICTURE_BUFFER, 
                                        Convert.ToSingle(dimens.Width), 
                                        Convert.ToSingle(dimens.Height)
                                        );
                                    break;
                            }

                        }
                        break;
                    case 4:
                        runningHeight = 0;
                        runningWidth = 0;
                        for (int i = 0; i < 4; ++i)
                        {
                            img = System.Drawing.Image.FromFile(imagePaths[i]);
                            dimens = ResizePicture(img.Width, img.Height, 2);
                            switch (i)
                            {
                                case 0:
                                    // first picure will be placed above the second, will use returned values without modifying them
                                    slide.Shapes.AddPicture(
                                        imageTempPath, 
                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 
                                        PICTURE_BOX_LEFT, PICTURE_BOX_TOP, 
                                        Convert.ToSingle(dimens.Width), 
                                        Convert.ToSingle(dimens.Height)
                                        );
                                    runningWidth = Convert.ToSingle(dimens.Width);
                                    runningHeight = Convert.ToSingle(dimens.Height);
                                    break;
                                case 1:
                                    // second picture will be placed to the right of the first, will add width of first to left value
                                    slide.Shapes.AddPicture(
                                        imageTempPath, 
                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 
                                        PICTURE_BOX_LEFT + runningWidth + PICTURE_BUFFER, 
                                        PICTURE_BOX_TOP, 
                                        Convert.ToSingle(dimens.Width), 
                                        Convert.ToSingle(dimens.Height)
                                        );
                                    runningWidth = Convert.ToSingle(dimens.Width);
                                    runningHeight = Convert.ToSingle(dimens.Height);
                                    break;
                                case 2:
                                    // third picture will be placed below the first, will add height of the first to top value
                                    slide.Shapes.AddPicture(
                                        imageTempPath, 
                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue,
                                        PICTURE_BOX_LEFT, 
                                        PICTURE_BOX_TOP + runningHeight + PICTURE_BUFFER, 
                                        Convert.ToSingle(dimens.Width), 
                                        Convert.ToSingle(dimens.Height)
                                        );
                                    runningWidth = Convert.ToSingle(dimens.Width);
                                    break;
                                case 3:
                                    // fourth picture will be placed below the second, will add width of the third to width, height of the second to top
                                    slide.Shapes.AddPicture(
                                        imageTempPath, 
                                        Microsoft.Office.Core.MsoTriState.msoFalse,
                                        Microsoft.Office.Core.MsoTriState.msoTrue, 
                                        PICTURE_BOX_LEFT + runningWidth + PICTURE_BUFFER,
                                        PICTURE_BOX_TOP + runningHeight + PICTURE_BUFFER, 
                                        Convert.ToSingle(dimens.Width), 
                                        Convert.ToSingle(dimens.Height)
                                        );
                                    break;
                            }
                        }
                        break;
                }

                // save new powerpoint
                pptPresentation.SaveAs(saveFileDialog.FileName, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            }
        }

        /// <summary>
        /// Pass the width, height, number of pictures to this function in order to resize
        /// </summary>
        /// <param name="imageWidth"></param>
        /// <param name="imageHeight"></param>
        /// <param name="numPictures"></param>
        /// <param name="pictureNumber"></param>
        /// <returns>
        /// a class that holds the new size
        /// </returns>
        private PictureDimensions ResizePicture(double imageWidth, double imageHeight, int numPictures)
        {
            if (imageHeight > PICTURE_BOX_HEIGHT || imageWidth > PICTURE_BOX_WIDTH)
            {
                double resize = 0.9;
                double heightDifference = imageHeight - PICTURE_BOX_HEIGHT;
                double widthDifference = imageWidth - PICTURE_BOX_WIDTH;

                // want to modify the biggest side first to make sure it within picture box while retain aspect ratio
                if (heightDifference > widthDifference)
                {
                    // modify the height first then use same resize on the width
                    double newImageHeight = 0;
                    while (PICTURE_BOX_HEIGHT < imageHeight)
                    {
                        newImageHeight = imageHeight * resize;
                        if (imageHeight < PICTURE_BOX_HEIGHT)
                        {
                            imageHeight = newImageHeight;
                            imageHeight /= numPictures;
                        }
                        else
                        {
                            resize -= .1;
                            if (resize.Equals(0)) break;
                        }
                    }
                    imageWidth *= resize;
                    imageWidth /= numPictures;
                }
                else
                {
                    // modify the width first then use same resize on the height
                    double newImageWidth = 0;
                    while (PICTURE_BOX_WIDTH < imageWidth)
                    {
                        newImageWidth = imageWidth * resize;
                        if (newImageWidth < PICTURE_BOX_WIDTH)
                        {
                            imageWidth = newImageWidth;
                            imageWidth /= numPictures;
                        }
                        else
                        {
                            resize -= .1;
                            if (resize.Equals(0)) break;
                        }
                    }
                    imageHeight *= resize;
                    imageHeight /= numPictures;
                }
            }
            return new PictureDimensions(){Height = imageHeight, Width = imageWidth};
        }
        #endregion
    }
}
