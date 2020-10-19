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
using System.Windows;
using System.Windows.Shapes;
using RiggsBurnham_PresentationMaker.Models;
using Image = System.Windows.Controls.Image;
using Google.Apis;
using GoogleLibrary;
using ILibrary;
using Application = Microsoft.Office.Interop.PowerPoint.Application;


namespace RiggsBurnham_PresentationMaker.ViewModels
{
    public class PresentationMakerViewModel : INotifyPropertyChanged
    {
        #region constants
        // these constants are for setting up the display area on the slide for where all pictures will go in the powerpoint 
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
        private int _mainWindowWidth = 1000;
        private string _selectedImageUrl = "";
        private IData _selectedImage;
        private IData _selectedExportImage;
        private List<string> _imagePaths;
        private ObservableCollection<IData> _selectedImages;
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
            SelectedImages = new ObservableCollection<IData>();
            SavePowerpointCommand = new DelegateCommand(SavePowerpoint);
            SearchImagesCommand = new DelegateCommand(SearchImages);
            AddImageCommand = new DelegateCommand(AddImage);
            RemoveSelectedExportImageCommand = new DelegateCommand(RemoveSelectedExportImage);
            SelectedImageChangedCommand = new DelegateCommand<object>(SelectedImageChanged);
            SelectedExportImageChangedCommand = new DelegateCommand<object>(SelectedExportImageChanged);
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
                // there are 5 columns, where the images are taking up two of the columns, so multiple by .4 (2/5) to get the width of the images gallery
                // want to display 2 images on each row so divide that result by 2
                // take that result and minus 11 for the width of the scrollbar
                double width = System.Convert.ToDouble(MainWindowWidth);
                double result = ((.4 * width) / 2) - 11;
                return System.Convert.ToInt32(result);
            }
        }

        public int SelectedImagesWidth
        {
            get
            {
                // there are 5 columns, where the selected images are taking up 1 of the columns, so multiple by .2 (1/5) to get the width of the images gallery
                // want to display 1 images on each row so no division is necessary
                // take that result and minus 11 for the width of the scrollbar
                double width = System.Convert.ToDouble(MainWindowWidth);
                double result = (.2 * width) - 11;
                return System.Convert.ToInt32(result);
            }
        }

        public int ImageGalleryWidth
        {
            get
            {
                // there are 5 columns, where the images are taking up two of the columns, so multiple by .4 (2/5) to get the width of the images gallery
                double width = System.Convert.ToDouble(MainWindowWidth);
                double result = .4 * width;
                return System.Convert.ToInt32(result);
            }
        }

        public int SelectedImageGalleryWidth
        {
            get
            {
                // there are 5 columns, where the selected images are taking up 1 of the columns, so multiple by .2 (1/5) to get the width of the images gallery
                double width = System.Convert.ToDouble(MainWindowWidth);
                double result = .2 * width;
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

        public IData SelectedImage
        {
            get => _selectedImage;
            set
            {
                _selectedImage = value;
                NotifyPropertyChanged();
            }
        }

        public IData SelectedExportImage
        {
            get => _selectedExportImage;
            set
            {
                _selectedExportImage = value;
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

        public ObservableCollection<IData> SelectedImages
        {
            get => _selectedImages;
            set
            {
                _selectedImages = value;
                NotifyPropertyChanged();
            }
        }
        #endregion

        #region commands
        public DelegateCommand SavePowerpointCommand { get; set; }
        private void SavePowerpoint()
        {
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
                switch (SelectedImages.Count)
                {
                    case 1:
                        // resize picture to fit inside box while retaining same aspect ratio...
                        imgData = new WebClient().DownloadData(SelectedImages[0].URL);
                        imgStream = new MemoryStream(imgData);
                        img = System.Drawing.Image.FromStream(imgStream);
                        dimens = ResizePicture(img.Width, img.Height, false);
                        slide.Shapes.AddPicture(
                            SelectedImages[0].URL, 
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
                            imgData = new WebClient().DownloadData(SelectedImages[i].URL);
                            imgStream = new MemoryStream(imgData);
                            img = System.Drawing.Image.FromStream(imgStream);
                            dimens = ResizePicture(img.Width, img.Height, true);
                            switch (i)
                            {
                                case 0:
                                    // first picure will be placed above the second, will use returned values without modifying them
                                    slide.Shapes.AddPicture(
                                        SelectedImages[i].URL, 
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
                                        SelectedImages[i].URL, 
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
                            imgData = new WebClient().DownloadData(SelectedImages[i].URL);
                            imgStream = new MemoryStream(imgData);
                            img = System.Drawing.Image.FromStream(imgStream);
                            dimens = ResizePicture(img.Width, img.Height, true);
                            switch (i)
                            {
                                case 0:
                                    // first picure will be placed above the second, will use returned values without modifying them
                                    slide.Shapes.AddPicture(
                                        SelectedImages[i].URL, 
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
                                        SelectedImages[i].URL,
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
                                        SelectedImages[i].URL, 
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
                            imgData = new WebClient().DownloadData(SelectedImages[i].URL);
                            imgStream = new MemoryStream(imgData);
                            img = System.Drawing.Image.FromStream(imgStream);
                            dimens = ResizePicture(img.Width, img.Height, true);
                            switch (i)
                            {
                                case 0:
                                    // first picure will be placed above the second, will use returned values without modifying them
                                    slide.Shapes.AddPicture(
                                        SelectedImages[i].URL, 
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
                                        SelectedImages[i].URL, 
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
                                        SelectedImages[i].URL, 
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
                                        SelectedImages[i].URL, 
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

        //private void AddImage()
        //{
        //    if (SelectedImageUrl == "") return;
        //    ImagePaths.Add(SelectedImageUrl);
        //}
        private void AddImage()
        {
            if (SelectedImage == null) return;
            if (SelectedImages.Count == 4)
            {
                MessageBox.Show("Can only add 4 pictures, please remove one before adding",
                    "Error - Too many pictures");
                return;
            }
            SelectedImages.Add(SelectedImage);
        }

        public DelegateCommand RemoveSelectedExportImageCommand { get; set; }

        private void RemoveSelectedExportImage()
        {
            if (SelectedExportImage == null) return;
            if (!SelectedImages.Contains(SelectedExportImage)) return;
            SelectedImages.Remove(SelectedExportImage);
        }

        public DelegateCommand<object> SelectedImageChangedCommand { get; set; }

        //private void SelectedImageChanged(object selectedItem)
        //{
        //    IData img = selectedItem as IData;
        //    if (img == null) return;
        //    SelectedImageUrl = img.URL;
        //}

        private void SelectedImageChanged(object selectedItem)
        {
            IData img = selectedItem as IData;
            if (img == null) return;
            SelectedImage = img;
        }

        public DelegateCommand<object> SelectedExportImageChangedCommand { get; set; }

        private void SelectedExportImageChanged(object selectedItem)
        {
            IData img = selectedItem as IData;
            if (img == null) return;
            SelectedExportImage = img;
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
                        }
                        else
                        {
                            resize -= .1;
                            if (resize.Equals(0)) break;
                        }
                    }
                    imageWidth *= resize;
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
                        }
                        else
                        {
                            resize -= .1;
                            if (resize.Equals(0)) break;
                        }
                    }
                    imageHeight *= resize;
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
