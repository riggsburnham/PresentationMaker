using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.IO;
using System.Net;
using System.Net.Http;
using System.Windows.Shapes;
using RiggsBurnham_PresentationMaker.Models;
using Image = System.Windows.Controls.Image;
using Google.Apis;
using GoogleLibrary;
using ILibrary;


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
        private GoogleImages _googleImages;
        private ObservableCollection<IData> _images;
        private int _mainWindowHeight = 450;
        private int _mainWindowWidth = 800;
        private string _selectedImageUrl = "";
        List<string> _imagePaths;
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
            GoogleImages = new GoogleImages();
            ImagePaths = new List<string>();
            SavePowerpointCommand = new DelegateCommand(SavePowerpoint);
            SearchImagesCommand = new DelegateCommand(SearchImages);
            AddImageCommand = new DelegateCommand(AddImage);
            SelectedImageChangedCommand = new DelegateCommand<object>(SelectedImageChanged);

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

        public GoogleImages GoogleImages
        {
            get => _googleImages;
            set => _googleImages = value;
        }

        public ObservableCollection<IData> Images
        {
            get => _images;
            set
            {
                _images = value;
                NotifyPropertyChanged();
            }
        }

        public int MainWindowHeight
        {
            get => _mainWindowHeight;
            set => _mainWindowHeight = value;
        }

        public int MainWindowWidth
        {
            get => _mainWindowWidth;
            set => _mainWindowWidth = value;
        }

        public int ImagesWidth
        {
            get
            {
                // there are two columns, where the images are taking up the second column so, multiple width by .5 to get the width of the images gallery
                // want to display 2 images on each row so divide that result by 2
                // take that result and minus 11 for the width of the scrollbar
                double width = System.Convert.ToDouble(MainWindowWidth);
                double result = ((.5 * width) / 2) - 11;
                return System.Convert.ToInt32(result);
            }
        }

        public int ImageGalleryWidth
        {
            get
            {
                double width = System.Convert.ToDouble(MainWindowWidth);
                double result = .5 * width;
                return System.Convert.ToInt32(result);
            }
        }

        public string SelectedImageUrl
        {
            get => _selectedImageUrl;
            set
            {
                _selectedImageUrl = value;
                NotifyPropertyChanged();
            }
        }

        public List<string> ImagePaths
        {
            get => _imagePaths;
            set
            {
                _imagePaths = value;
                NotifyPropertyChanged();
            }
        }
        #endregion

        #region commands
        public DelegateCommand SavePowerpointCommand { get; set; }
        private void SavePowerpoint()
        {
            // TODO: remove this once you are pulling images from google...
            //string imageTempPath = @"C:\Users\riggs\Pictures\Asuna Closeup.JPG";

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

                

                // will most likely limit number of images on a slide to 4...
                PictureDimensions dimens = new PictureDimensions();
                System.Drawing.Image img;
                byte[] imgData;
                MemoryStream imgStream;
                float runningHeight = 0;
                float runningWidth = 0;
                switch (ImagePaths.Count)
                {
                    case 1:
                        // resize picture to fit inside box while retaining same aspect ratio...
                        //img = System.Drawing.Image.FromFile(imageTempPath);
                        imgData = new WebClient().DownloadData(ImagePaths[0]);
                        imgStream = new MemoryStream(imgData);
                        img = System.Drawing.Image.FromStream(imgStream);
                        //dimens = ResizePicture(img.Width, img.Height, 1);
                        dimens = ResizePicture(img.Width, img.Height, false);
                        slide.Shapes.AddPicture(
                            ImagePaths[0], 
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
                            //img = System.Drawing.Image.FromFile(ImagePaths[i]);
                            imgData = new WebClient().DownloadData(ImagePaths[i]);
                            imgStream = new MemoryStream(imgData);
                            img = System.Drawing.Image.FromStream(imgStream);
                            //dimens = ResizePicture(img.Width, img.Height, 2);
                            dimens = ResizePicture(img.Width, img.Height, true);
                            switch (i)
                            {
                                case 0:
                                    // first picure will be placed above the second, will use returned values without modifying them
                                    slide.Shapes.AddPicture(
                                        ImagePaths[i], 
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
                                        ImagePaths[i], 
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
                            //img = System.Drawing.Image.FromFile(ImagePaths[i]);
                            imgData = new WebClient().DownloadData(ImagePaths[i]);
                            imgStream = new MemoryStream(imgData);
                            img = System.Drawing.Image.FromStream(imgStream);
                            //dimens = ResizePicture(img.Width, img.Height, 2); // TODO: looks like i should've passed 3 here.. demo was working though
                            dimens = ResizePicture(img.Width, img.Height, true);
                            switch (i)
                            {
                                case 0:
                                    // first picure will be placed above the second, will use returned values without modifying them
                                    slide.Shapes.AddPicture(
                                        ImagePaths[i], 
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
                                        ImagePaths[i],
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
                                        ImagePaths[i], 
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
                            //img = System.Drawing.Image.FromFile(ImagePaths[i]);
                            imgData = new WebClient().DownloadData(ImagePaths[i]);
                            imgStream = new MemoryStream(imgData);
                            img = System.Drawing.Image.FromStream(imgStream);
                            //dimens = ResizePicture(img.Width, img.Height, 2);
                            dimens = ResizePicture(img.Width, img.Height, true);
                            switch (i)
                            {
                                case 0:
                                    // first picure will be placed above the second, will use returned values without modifying them
                                    slide.Shapes.AddPicture(
                                        ImagePaths[i], 
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
                                        ImagePaths[i], 
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
                                        ImagePaths[i], 
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
                                        ImagePaths[i], 
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

        public DelegateCommand SearchImagesCommand { get; set; }

        private void SearchImages()
        {
            List<string> titleString = Title.Split(' ').ToList();
            List<string> descriptionString = Description.Split(' ').ToList();
            titleString.AddRange(descriptionString);
            GoogleImages.SearchGoogleImages(titleString);
            Images = LoadGoogleImages();
        }

        public DelegateCommand AddImageCommand { get; set; }

        private void AddImage()
        {
            if (SelectedImageUrl == "") return;
            ImagePaths.Add(SelectedImageUrl);
        }

        public DelegateCommand<object> SelectedImageChangedCommand { get; set; }

        private void SelectedImageChanged(object selectedItem)
        {
            IData img = selectedItem as IData;
            if (img == null) return;
            SelectedImageUrl = img.URL;
        }
        #endregion

        #region private functions
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
        private PictureDimensions ResizePicture(double imageWidth, double imageHeight, bool moreThanOnePicture)
        {
            int gtOnePicture = 1;
            if (moreThanOnePicture == true)
            {
                gtOnePicture = 2;
            }
            if (imageHeight > (PICTURE_BOX_HEIGHT / gtOnePicture) || imageWidth > (PICTURE_BOX_WIDTH / gtOnePicture))
            {
                double resize = 0.9;
                double heightDifference = imageHeight - PICTURE_BOX_HEIGHT;
                double widthDifference = imageWidth - PICTURE_BOX_WIDTH;

                // want to modify the biggest side first to make sure it within picture box while retain aspect ratio
                if (heightDifference > widthDifference)
                {
                    // modify the height first then use same resize on the width
                    double newImageHeight = 0;
                    while ((PICTURE_BOX_HEIGHT / gtOnePicture) < imageHeight)
                    {
                        newImageHeight = imageHeight * resize;
                        if (newImageHeight < PICTURE_BOX_HEIGHT / gtOnePicture)
                        {
                            imageHeight = newImageHeight;
                            //imageHeight /= numPictures;
                        }
                        else
                        {
                            resize -= .1;
                            if (resize.Equals(0)) break;
                        }
                    }
                    imageWidth *= resize;
                    //imageWidth /= numPictures;
                }
                else
                {
                    // modify the width first then use same resize on the height
                    double newImageWidth = 0;
                    while ((PICTURE_BOX_WIDTH / gtOnePicture) < imageWidth)
                    {
                        newImageWidth = imageWidth * resize;
                        if (newImageWidth < (PICTURE_BOX_WIDTH / gtOnePicture))
                        {
                            imageWidth = newImageWidth;
                            //imageWidth /= numPictures;
                        }
                        else
                        {
                            resize -= .1;
                            if (resize.Equals(0)) break;
                        }
                    }
                    imageHeight *= resize;
                    //imageHeight /= numPictures;
                }
            }
            return new PictureDimensions() { Height = imageHeight, Width = imageWidth };
        }

        private ObservableCollection<IData> LoadGoogleImages()
        {
            ObservableCollection<IData> images = new ObservableCollection<IData>();
            foreach (Item item in _googleImages.GData.items)
            {
                foreach (Cse_Image img in item.pagemap.cse_image)
                {
                    images.Add(img);
                }
            }
            return images;
        }
        #endregion
    }
}
