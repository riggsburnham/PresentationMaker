using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using Prism.Commands;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Input;
using RiggsBurnham_PresentationMaker.Models;
using GoogleLibrary;
using ILibrary;
using RiggsBurnham_PresentationMaker.Views;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using TextRange = Microsoft.Office.Interop.PowerPoint.TextRange;


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
        private ObservableCollection<IData> _selectedImages;
        private ErrorWindow _tooManyPicturesError;
        private ErrorWindow _failedToLoadPictureError;
        private string _plainTextTitle = "";
        private string _rtfTextTitle = "";
        private string _plainTextDescription = "";
        private string _rtfTextDescription = "";
        private List<CharacterStyle> _csTitleList;
        private List<CharacterStyle> _csDescriptionList;
        private bool _failedToLoadPicture = false;
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
            SelectedImages = new ObservableCollection<IData>();
            SavePowerpointCommand = new DelegateCommand(async () => await SavePowerpoint());
            SearchImagesCommand = new DelegateCommand(async () => await SearchImages());
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
                Xceed.Wpf.Toolkit.RichTextBox rtBox = new Xceed.Wpf.Toolkit.RichTextBox(new FlowDocument());
                rtBox.Text = value;
                rtBox.TextFormatter = new Xceed.Wpf.Toolkit.RtfFormatter();
                _rtfTextTitle = rtBox.Text;
                rtBox.TextFormatter = new Xceed.Wpf.Toolkit.PlainTextFormatter();
                _plainTextTitle = rtBox.Text.Remove(rtBox.Text.Length - 2);
                _csTitleList = ConvertRtfTextToCharacterStyles(_rtfTextTitle);
                _title = _rtfTextTitle;
                NotifyPropertyChanged();
            }
        }

        public string Description
        {
            get => _description;
            set
            {
                //_description = value;
                Xceed.Wpf.Toolkit.RichTextBox rtBox = new Xceed.Wpf.Toolkit.RichTextBox(new FlowDocument());
                rtBox.Text = value;
                rtBox.TextFormatter = new Xceed.Wpf.Toolkit.RtfFormatter();
                _rtfTextDescription = rtBox.Text;
                rtBox.TextFormatter = new Xceed.Wpf.Toolkit.PlainTextFormatter();
                _plainTextDescription = rtBox.Text.Remove(rtBox.Text.Length - 2);
                _csDescriptionList = ConvertRtfTextToCharacterStyles(_rtfTextDescription);
                _description = _rtfTextDescription;
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

        public ObservableCollection<IData> SelectedImages
        {
            get => _selectedImages;
            set
            {
                _selectedImages = value;
                NotifyPropertyChanged();
            }
        }

        public ErrorWindow TooManyPicturesError
        {
            get => _tooManyPicturesError;
            set => _tooManyPicturesError = value;
        }

        public ErrorWindow FailedToLoadPictureError
        {
            get => _failedToLoadPictureError;
            set => _failedToLoadPictureError = value;
        }
        #endregion

        #region commands
        public DelegateCommand SavePowerpointCommand { get; set; }
        private async Task SavePowerpoint()
        {
            await Task.Run(PerformSave);
            if (_failedToLoadPicture == true)
            {
                FailedToLoadPictureError = new ErrorWindow(this, "!!! Error !!!", "Picture(s) failed to load");
                FailedToLoadPictureError.Show();
                _failedToLoadPicture = false;
            }
        }

        public DelegateCommand SearchImagesCommand { get; set; }

        private async Task SearchImages()
        {
            await Task.Run(PerformSearch);
        }

        public DelegateCommand AddImageCommand { get; set; }

        private void AddImage()
        {
            if (SelectedImage == null) return;
            if (SelectedImages.Count == 4)
            {
                TooManyPicturesError = new ErrorWindow(this, "!!! Error !!!", "Can only add 4 pictures, please remove one before adding.");
                TooManyPicturesError.Show();
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
        /// Pass the width, height, boolean value if there is more than 1 picture to be displayed
        /// </summary>
        /// <param name="imageWidth"></param>
        /// <param name="imageHeight"></param>
        /// <param name="moreThanOnePicture"></param>
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
                if (item.pagemap?.cse_image == null) continue;
                foreach (Cse_Image img in item.pagemap.cse_image)
                {
                    images.Add(img);
                }
            }
            return images;
        }

        private void PerformSave()
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

                objText.Text = _plainTextTitle;
                for (var i = 0; i < objText.Text.Length; ++i)
                {
                    if (_csTitleList[i].IsBold)
                    {
                        objText.Characters(i+1, 1).Font.Bold = MsoTriState.msoTrue;
                        // TODO: can further implement italic and, underline
                    }
                }
                objText.Font.Name = "Arial";
                objText.Font.Size = 32;

                // description...
                objText = slide.Shapes[2].TextFrame.TextRange;

                // modifying the description to only take up left half of page with a buffer gap of 10.
                slide.Shapes[2].Width /= 2;
                slide.Shapes[2].Width -= 10;

                objText.Text = _plainTextDescription;
                for (var i = 0; i < objText.Text.Length; ++i)
                {
                    if (_csDescriptionList[i].IsBold)
                    {
                        objText.Characters(i + 1, 1).Font.Bold = MsoTriState.msoTrue;
                        // TODO: can further implement italic and, underline
                    }
                }

                objText.Font.Name = "Arial";
                objText.Font.Size = 16;

                // will most likely limit number of images on a slide to 4...
                PictureDimensions dimens;
                float runningHeight = 0;
                float runningWidth = 0;
                bool failedLoadingPicture = false;
                switch (SelectedImages.Count)
                {
                    case 1:
                        // resize picture to fit inside box while retaining same aspect ratio...
                        try
                        {
                            dimens = CalculateDimensions(SelectedImages[0].URL);
                        }
                        catch
                        {
                            _failedToLoadPicture = true;
                            return;
                        }
                        
                        AddPicture(slide, SelectedImages[0].URL, PICTURE_BOX_LEFT, PICTURE_BOX_TOP, dimens.Width, dimens.Height);
                        break;
                    case 2:
                        runningHeight = 0;
                        for (int i = 0; i < 2; ++i)
                        {
                            try
                            {
                                dimens = CalculateDimensions(SelectedImages[i].URL);
                            }
                            catch
                            {
                                _failedToLoadPicture = true;
                                continue;
                            }
                            switch (i)
                            {
                                case 0:
                                    // first picure will be placed above the second, will use returned values without modifying them
                                    AddPicture(slide, SelectedImages[i].URL, PICTURE_BOX_LEFT, PICTURE_BOX_TOP, dimens.Width, dimens.Height);
                                    runningHeight = Convert.ToSingle(dimens.Height);
                                    break;
                                case 1:
                                    // second picture will be placed below the first, will need to modify the top by the height the the picture before...
                                    AddPicture(slide, SelectedImages[i].URL, PICTURE_BOX_LEFT, PICTURE_BOX_TOP + runningHeight + PICTURE_BUFFER, dimens.Width, dimens.Height);
                                    break;
                            }

                        }
                        break;
                    case 3:
                        runningHeight = 0;
                        runningWidth = 0;
                        for (int i = 0; i < 3; ++i)
                        {
                            try
                            {
                                dimens = CalculateDimensions(SelectedImages[i].URL);
                            }
                            catch
                            {
                                _failedToLoadPicture = true;
                                continue;
                            }
                            switch (i)
                            {
                                case 0:
                                    // first picure will be placed above the second, will use returned values without modifying them
                                    AddPicture(slide, SelectedImages[i].URL, PICTURE_BOX_LEFT, PICTURE_BOX_TOP, dimens.Width, dimens.Height);
                                    runningHeight = Convert.ToSingle(dimens.Height);
                                    break;
                                case 1:
                                    // second picture will be placed below the first, will need to modify the top by the height the the picture before...
                                    AddPicture(slide, SelectedImages[i].URL, PICTURE_BOX_LEFT, PICTURE_BOX_TOP + runningHeight + PICTURE_BUFFER, dimens.Width, dimens.Height);
                                    runningWidth = Convert.ToSingle(dimens.Width);
                                    break;
                                case 2:
                                    // third picture will be placed to the right of the second, will need to modify both top and left by the height and width of previous pic
                                    AddPicture(slide, SelectedImages[i].URL, PICTURE_BOX_LEFT + runningWidth + PICTURE_BUFFER, PICTURE_BOX_TOP + runningHeight + PICTURE_BUFFER, dimens.Width, dimens.Height);
                                    break;
                            }

                        }
                        break;
                    case 4:
                        runningHeight = 0;
                        runningWidth = 0;
                        float oldRunningHeight = 0;
                        float oldRunningWidth = 0;
                        for (int i = 0; i < 4; ++i)
                        {
                            try
                            {
                                dimens = CalculateDimensions(SelectedImages[i].URL);
                            }
                            catch
                            {
                                _failedToLoadPicture = true;
                                continue;
                            }
                            switch (i)
                            {
                                case 0:
                                    // first picure will be placed above the second, will use returned values without modifying them
                                    AddPicture(slide, SelectedImages[i].URL, PICTURE_BOX_LEFT, PICTURE_BOX_TOP, dimens.Width, dimens.Height);
                                    runningWidth = Convert.ToSingle(dimens.Width);
                                    runningHeight = Convert.ToSingle(dimens.Height);
                                    break;
                                case 1:
                                    // second picture will be placed to the right of the first, will add width of first to left value
                                    AddPicture(slide, SelectedImages[i].URL, PICTURE_BOX_LEFT + runningWidth + PICTURE_BUFFER, PICTURE_BOX_TOP, dimens.Width, dimens.Height);
                                    oldRunningWidth = runningWidth;
                                    runningWidth = Convert.ToSingle(dimens.Width);
                                    oldRunningHeight = runningHeight;
                                    runningHeight = Convert.ToSingle(dimens.Height);
                                    break;
                                case 2:
                                    // third picture will be placed below the first, will add height of the first to top value
                                    AddPicture(slide, SelectedImages[i].URL, PICTURE_BOX_LEFT, PICTURE_BOX_TOP + oldRunningHeight + PICTURE_BUFFER, dimens.Width, dimens.Height);
                                    runningWidth = Convert.ToSingle(dimens.Width);
                                    break;
                                case 3:
                                    // fourth picture will be placed below the second, will add width of the third to width, height of the second to top
                                    float widthShiftValue = runningWidth;
                                    if (oldRunningHeight > runningHeight)
                                    {
                                        if (oldRunningWidth > runningWidth)
                                        {
                                            widthShiftValue = oldRunningWidth;
                                        }
                                        else
                                        {
                                            widthShiftValue = runningWidth;
                                        }
                                    }
                                    AddPicture(slide, SelectedImages[i].URL, PICTURE_BOX_LEFT + widthShiftValue + PICTURE_BUFFER, PICTURE_BOX_TOP + runningHeight + PICTURE_BUFFER, dimens.Width, dimens.Height);
                                    break;
                            }
                        }
                        break;
                }
                // save new powerpoint
                pptPresentation.SaveAs(saveFileDialog.FileName, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                //if (failedLoadingPicture == true)
                //{
                //    //FailedToLoadPictureError = new ErrorWindow(this, "Failed Loading a Picture", "A picture failed to load");
                //    FailedToLoadPictureError.Show();
                //}
            }
        }

        private void PerformSearch()
        {
            List<string> titleString = _plainTextTitle.Split(' ').ToList();
            List<string> descriptionString = _plainTextDescription.Split(' ').ToList();
            titleString.AddRange(descriptionString);
            GoogleImages.SearchGoogleImages(titleString);
            Images = LoadGoogleImages();
        }

        private List<CharacterStyle> ConvertRtfTextToCharacterStyles(string rtfText)
        {
            List<CharacterStyle> csList = new List<CharacterStyle>();

            //_title = value;
            var rtfParts = rtfText.Split('{', '}');
            List<string> found = new List<string>();
            foreach (var part in rtfParts)
            {
                if (part.Contains("\\ltrch"))
                {
                    found.Add(part);
                }
            }
            //List<CharacterStyle> csList = new List<CharacterStyle>();
            int i = 0;
            bool isBold = false;
            string word = "";
            foreach (string styleWord in found)
            {
                if (styleWord.Contains("\\b\\ltrch"))
                {
                    isBold = true;
                    word = styleWord.Remove(0, 9);
                }
                else
                {
                    word = styleWord.Remove(0, 7);
                }
                foreach (char character in word)
                {
                    CharacterStyle cs = new CharacterStyle();
                    cs.Character = character;
                    cs.IsBold = isBold;
                    cs.Position = i;
                    csList.Add(cs);
                    ++i;
                }
                isBold = false;
            }

            return csList;
        }

        private void AddPicture(Slide slide, string url, float left, float top, double width, double height)
        {
            slide.Shapes.AddPicture(
                url,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoTrue,
                left, top,
                Convert.ToSingle(width),
                Convert.ToSingle(height)
            );
        }

        private PictureDimensions CalculateDimensions(string url)
        {
            //try
            //{
                PictureDimensions dimens = new PictureDimensions();
                byte[] imgData = new WebClient().DownloadData(url);
                MemoryStream imgStream = new MemoryStream(imgData);
                System.Drawing.Image img = System.Drawing.Image.FromStream(imgStream);
                dimens = ResizePicture(img.Width, img.Height, false);
                return dimens;
            //}
            //catch
            //{
            //    FailedToLoadPictureError = new ErrorWindow(this, "Failed Loading a Picture", "A picture failed to load");
            //    FailedToLoadPictureError.Show();
            //    throw;
            //}
        }
        #endregion
    }
}
