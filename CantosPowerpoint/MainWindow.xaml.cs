using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using System.Collections;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MessageBox = System.Windows.MessageBox;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;
using System.Windows.Input;
using System.Windows.Interop;

using System.Threading;
using System.Globalization;
using Infralution.Localization.Wpf;
using System.Reflection;
using System.Resources;

namespace Iglesia.CantosPowerpoint
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        // SetFocus will just focus the keyboard on your application, but not bring your process to front.
        // You don't need it here, SetForegroundWindow does the same.
        // Just for documentation.
        [DllImport("user32.dll")]
        static extern IntPtr SetFocus(HandleRef hWnd);

        /// <summary>
        /// The resource manager to use for this class.  Holding a strong reference to the
        /// Resource Manager keeps it in the cache while ever there are methods that
        /// are using it.
        /// </summary>
        private ResourceManager _resourceManager;

        private string CCName = "Cantos Del Camino";
        private string CEName = "Cantos Espirituales";

        private string CCPath = string.Empty;
        private string CEPath = string.Empty;

        private string currentSongBookPath = string.Empty;

        private double PPTVersion = 0;

        private PowerPoint.Application objPPT;
        private PowerPoint.Presentation objPresSongs;
        private PowerPoint.SlideShowWindow slideShowWin;

        private string Filename = Path.GetTempPath() + @"\ChurchSongs.pptx";

        private ArrayList SongsList = new ArrayList();

        public MainWindow()
        {
            InitializeComponent();

            // set the initial application UI Culture based on the users
            // current regional settings
            //
            CultureManager.UICulture = Thread.CurrentThread.CurrentCulture;
            CultureManager.UICultureChanged += new EventHandler(CultureManager_UICultureChanged);
            UpdateLanguageMenus();

            /*try
            {

                string filename = "CantosPowerpoint.exe";

                // Try to load the assembly.
                Assembly assem = Assembly.LoadFrom(filename);
                Debug.WriteLine(String.Format("File: {0}", filename));

                // Enumerate the resource files.
                string[] resNames = assem.GetManifestResourceNames();
                if (resNames.Length == 0)
                    Debug.WriteLine("   No resources found.");

                foreach (var resName in resNames)
                    Debug.WriteLine(String.Format("   Resource: {0}", resName.Replace(".resources", "")));

                string test = getStringByName("Window.Title");

            }
            catch (Exception)
            {
            }*/

            if (!FindSongBooksPath())
            {

                MessageBox.Show("We cannot continue because we couldn't find the songbooks. Exiting.", getStringByName("Window.Title"), MessageBoxButton.OK, MessageBoxImage.Error);

                // Exit application
                System.Windows.Application.Current.Shutdown();
            }

        }

        #region Localization functions
        /// <summary>
        /// Detach from UICultureChanged event
        /// </summary>
        /// <param name="e"></param>
        /// <remarks>
        /// If we don't detach from the event then the window will not get garbage collected
        /// </remarks>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            CultureManager.UICultureChanged -= new EventHandler(CultureManager_UICultureChanged);
        }

        /// <summary>
        /// Update the check state of the Language menu
        /// </summary>
        private void UpdateLanguageMenus()
        {
            string lang = CultureManager.UICulture.TwoLetterISOLanguageName.ToLower();
            _spanishMenuItem.IsChecked = (lang == "es");
            _englishMenuItem.IsChecked = (lang == "en");
            //_fileListBox.ItemsSource = System.Enum.GetValues(typeof(SampleEnum));
        }

        /// <summary>
        /// Update the language menus when the UI culture changes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CultureManager_UICultureChanged(object sender, EventArgs e)
        {
            UpdateLanguageMenus();
        }

        /// <summary>
        /// Select English as the User Interface language
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void _englishMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CultureManager.UICulture = new CultureInfo("en");
        }

        /// <summary>
        /// Select Spanish as the User Interface language
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void _spanishMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CultureManager.UICulture = new CultureInfo("es");
        }

        public string getStringByName(string resourceKey)
        {
            if (_resourceManager == null)
                _resourceManager = new ResourceManager("Iglesia.CantosPowerpoint.MainWindow", typeof(MainWindow).Assembly);

            if (_resourceManager != null)
                return _resourceManager.GetString(resourceKey, CultureManager.UICulture);
            else
                return null;
        }
        #endregion

        private Boolean FindSongBooksPath()
        {

            string dropboxFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Dropbox";

            CCPath = dropboxFolder + @"\HimnarioVirtual@2010\HV-" + CCName;

            // Test to see if Cantos del Camino songbook is present
            if (!Directory.Exists(CCPath))
            {

                FolderBrowserDialog fbd = new FolderBrowserDialog();
                fbd.RootFolder = System.Environment.SpecialFolder.MyComputer;
                fbd.ShowNewFolderButton = false;
                fbd.Description = "Select folder with \"" + CCName + "\" songbook";

                DialogResult result = fbd.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    CCPath = fbd.SelectedPath;
                    lblSongBook.Content = CCName;
                }
                else
                {
                    _ccMenuItem.Visibility = Visibility.Hidden;
                    _ceMenuItem.IsChecked = true;
                    CCPath = string.Empty;
                }
            }
            else
            {
                currentSongBookPath = CCPath;
                lblSongBook.Content = CCName;
            }

            CEPath = dropboxFolder + @"\HimnarioVirtual@2010\HV-Clásicos-" + CEName;

            // Test to see if Cantos Espirituales songbook is present
            if (!Directory.Exists(CEPath))
            {

                FolderBrowserDialog fbd = new FolderBrowserDialog();
                fbd.RootFolder = System.Environment.SpecialFolder.MyComputer;
                fbd.ShowNewFolderButton = false;
                fbd.Description = "Select folder with \"" + CEName + "\" songbook";

                DialogResult result = fbd.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    CEPath = fbd.SelectedPath;
                    lblSongBook.Content = CEName;
                }
                else
                {
                    _ceMenuItem.Visibility = Visibility.Hidden;
                    _ccMenuItem.IsChecked = true;
                    CEPath = string.Empty;
                }

            }
            else if (string.IsNullOrEmpty(CCPath))
            {
                currentSongBookPath = CEPath;
                lblSongBook.Content = CEName;
            }
            //MessageBox.Show("Songbooks Folder: " + SongBookPath, getStringByName("Window.Title"), MessageBoxButton.OK, MessageBoxImage.Information);

            return !string.IsNullOrEmpty(CCPath) || !string.IsNullOrEmpty(CEPath);

        }

        private void StartPowerpoint()
        {
            try
            {
                objPPT = new PowerPoint.Application();
                objPPT.Visible = MsoTriState.msoTrue;

                //Prevent Office Assistant (if installed) from displaying alert messages:
                try
                {
                    Boolean bAssistantOn = objPPT.Assistant.On;
                    objPPT.Assistant.On = false;
                }
                catch (Exception)
                {
                }

                objPresSongs = objPPT.Presentations.Add();
                objPresSongs.SaveAs(Filename, PowerPoint.PpSaveAsFileType.ppSaveAsDefault);

                _openPptMenuItem.IsEnabled = false;
                _closePptMenuItem.IsEnabled = true;

                tbarSlideShow.IsEnabled = true;
                tbarAddButton.IsEnabled = true;

                _ccMenuItem.IsEnabled = true;
                _ceMenuItem.IsEnabled = true;

                tboxSongsList.IsEnabled = true;
                tboxSongsList.Focus();

                if (double.TryParse(objPPT.Version, out PPTVersion))
                {
                    if (PPTVersion > 12.0)
                        tbarPresenterView.Visibility = Visibility.Visible;
                }

                SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
                //SetFocus(new HandleRef(null, Process.GetCurrentProcess().MainWindowHandle)); // not needed
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, getStringByName("Window.Title"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClosePowerpoint(Boolean Exiting = false)
        {
            try
            {
                objPresSongs.Close();
            }
            catch (Exception)
            {
            }

            try
            {
                objPPT.Quit();
            }
            catch (Exception)
            {
            }

            // Cleanup to ensure powerpoint is completely close
            /* // THIS CAUSED INSTABILITY, DON'T USE IT!
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if(slideShowWin != null)
                Marshal.FinalReleaseComObject(slideShowWin);
            if (objPresSongs != null)
                Marshal.FinalReleaseComObject(objPresSongs);
            if (objPPT != null)
                Marshal.FinalReleaseComObject(objPPT);
                */

            if (Exiting)
                return;

            _closePptMenuItem.IsEnabled = false;
            tbarSlideShow.IsEnabled = false;

            tbarAddButton.IsEnabled = false;
            tbarRemoveButton.IsEnabled = false;
            tbarPresenterView.IsEnabled = false;
            tbarStartButton.IsEnabled = false;
            tbarNextButton.IsEnabled = false;
            tbarPreviousButton.IsEnabled = false;
            tbarStopButton.IsEnabled = false;

            // Update slide number counter
            lblSlideNum.Content = "--";

            tboxSongsList.Clear();
            tboxSongsList.IsEnabled = false;

            lboxSongs.Items.Clear();

            _ccMenuItem.IsEnabled = false;
            _ceMenuItem.IsEnabled = false;

            _openPptMenuItem.IsEnabled = true;

            SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
            //SetFocus(new HandleRef(null, Process.GetCurrentProcess().MainWindowHandle)); // not needed

        }

        Boolean VerifyPowerpointIsOpen()
        {

            try
            {
                int count = objPresSongs.Slides.Count;
                return true;
            }
            catch (Exception ex)
            {

                if (ex.ToString().ToUpper().Contains("OBJECT DOES NOT EXIST"))
                {
                    objPPT = new PowerPoint.Application();
                    SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
                }
                return false;
            }

        }

        void ReOpenPresentation()
        {
            try
            {
                // We assume we already have a presentation created but was closed inadvertely
                if (objPPT.Presentations.Count == 0)
                    if (PPTVersion <= 12)
                    {
                        objPPT.Visible = MsoTriState.msoTrue;
                        objPresSongs = objPPT.Presentations.Open2007(Filename);
                    }
                    else
                        objPresSongs = objPPT.Presentations.Open(Filename);

                if (objPPT.Presentations.Count == 0)
                    throw new Exception();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to open presentation. Please restart application and try again." + "\n" +
                    ex.Message + "\n" + ex.StackTrace, getStringByName("Window.Title"), MessageBoxButton.OK, MessageBoxImage.Error);

                // Exit application
                System.Windows.Application.Current.Shutdown();
            }
        }
        Boolean VerifyPresentationIsOpen()
        {

            try
            {
                int count = objPresSongs.Slides.Count;

                return true;
            }
            catch (Exception ex)
            {

                if (ex.ToString().ToUpper().Contains("OBJECT DOES NOT EXIST"))
                {
                    VerifyPowerpointIsOpen();

                    MessageBox.Show("Powerpoint presentation missing. Restarting presentation...",
                        getStringByName("Window.Title"),
                        MessageBoxButton.OK, MessageBoxImage.Error);

                    objPresSongs = objPPT.Presentations.Add();

                    SongsList.Clear();

                    _closePptMenuItem.IsEnabled = true;
                    tbarSlideShow.IsEnabled = true;

                    tbarAddButton.IsEnabled = true;
                    tbarRemoveButton.IsEnabled = false;
                    tbarPresenterView.IsEnabled = false;
                    tbarStartButton.IsEnabled = false;
                    tbarNextButton.IsEnabled = false;
                    tbarPreviousButton.IsEnabled = false;
                    tbarStopButton.IsEnabled = false;

                    //rbCC.IsEnabled = true;
                    //rbCE.IsEnabled = true;

                    lboxSongs.Items.Clear();

                    tboxSongsList.Focus();

                    SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);

                    return false;

                }
                else
                    throw;
            }


        }

        private int AddSlides(int idxSlideInsert = -1)
        {
            try
            {

                if (!VerifyPresentationIsOpen())
                    idxSlideInsert = -1;

                PowerPoint.Slide objSlide;
                PowerPoint.CustomLayout objCustomLayout;

                // Parse numbers from string
                char[] delimiterChars = { ' ', ',', '.', ':', '\t', '\n' };
                string songsList = tboxSongsList.Text;

                string[] songNumbers = songsList.Split(delimiterChars);

                pBar.Value = 0;
                pBar.Maximum = songNumbers.Count() + 1; // Extra number to show progress while saving file

                //Create a new instance of our ProgressBar Delegate that points
                // to the ProgressBar's SetValue method.
                UpdateProgressBarDelegate updatePbDelegate =
                    new UpdateProgressBarDelegate(pBar.SetValue);

                int countSongs = 0;

                foreach (var number in songNumbers)
                {
                    int songNum;

                    // Check if valid number
                    if (int.TryParse(number, out songNum))
                    {

                        string[] files = Directory.GetFiles(currentSongBookPath,
                            String.Format("{0:000}*", songNum), SearchOption.TopDirectoryOnly);

                        // Try to find song file in songbook
                        // TODO, check when number matches more than one file
                        if (files.Count() > 0)
                        {

                            // Check where to start inserting slides
                            int SlideNumber;

                            if (idxSlideInsert < 0)
                                SlideNumber = objPresSongs.Slides.Count + 1;
                            else
                            {
                                int[] SelectedSlideNums = (int[])SongsList[idxSlideInsert + countSongs];
                                SlideNumber = SelectedSlideNums[0];
                            }

                            // Insert blank slide
                            objCustomLayout = objPresSongs.SlideMaster.CustomLayouts[1];
                            objSlide = objPresSongs.Slides.AddSlide(SlideNumber, objCustomLayout);
                            objSlide.FollowMasterBackground = MsoTriState.msoFalse;
                            objSlide.Background.Fill.ForeColor.RGB = Color.Black.ToArgb();
                            objSlide.Layout = PowerPoint.PpSlideLayout.ppLayoutBlank;

                            // Add line in the bottom
                            float BeginX = (float)((objPresSongs.PageSetup.SlideWidth / 2) - (objPresSongs.PageSetup.SlideWidth * .10));
                            float EndX = (float)((objPresSongs.PageSetup.SlideWidth / 2) + (objPresSongs.PageSetup.SlideWidth * .10));
                            float PosY = (float)(objPresSongs.PageSetup.SlideHeight - (objPresSongs.PageSetup.SlideHeight * 0.10));
                            PowerPoint.Shape objShape = objSlide.Shapes.AddLine(BeginX, PosY, EndX, PosY);

                            // Insert slides from songs file
                            int slidesAdded = objPresSongs.Slides.InsertFromFile(files[0], SlideNumber);

                            char[] delimiter = { '\\' };
                            string[] filenameParts = files[0].Split(delimiter);
                            string filename = filenameParts[filenameParts.Count() - 1];

                            string songbook = (((bool)_ccMenuItem.IsChecked) ? "CC" : "CE");

                            // Add (or insert) slide numbers in lists
                            int[] SlideNumbers = { SlideNumber, SlideNumber + slidesAdded };

                            if (idxSlideInsert < 0)
                            {
                                lboxSongs.Items.Add(songbook + ": " + filename);
                                SongsList.Add(SlideNumbers);
                            }
                            else {

                                lboxSongs.Items.Insert(idxSlideInsert + countSongs, songbook + ": " + filename);
                                SongsList.Insert(idxSlideInsert + countSongs, SlideNumbers);

                                // Increase slide numbers in rest of songs in list
                                for (int i = 0; i < lboxSongs.Items.Count - idxSlideInsert - countSongs - 1; i++)
                                {
                                    SlideNumbers = (int[])SongsList[idxSlideInsert + countSongs + i + 1];
                                    int[] NewSlideNums = { SlideNumbers[0] + slidesAdded + 1, SlideNumbers[1] + slidesAdded + 1 };

                                    SongsList[idxSlideInsert + countSongs + i + 1] = NewSlideNums;
                                }

                            }

                            int index = tboxSongsList.Text.IndexOf(number);
                            tboxSongsList.Text = tboxSongsList.Text.Remove(index, number.Length);

                            countSongs++;

                        }

                    }

                    pBar.Value += 1;

                    // Update the Value of the ProgressBar:
                    Dispatcher.Invoke(updatePbDelegate,
                        System.Windows.Threading.DispatcherPriority.Background,
                        new object[] { System.Windows.Controls.ProgressBar.ValueProperty, pBar.Value });

                }

                tboxSongsList.Text = tboxSongsList.Text.Trim();

                objPresSongs.SaveAs(Filename, PowerPoint.PpSaveAsFileType.ppSaveAsDefault);

                // Final upgrade to progress bar after saving file
                pBar.Value += 1;

                if (countSongs > 0)
                {
                    tbarPresenterView.IsEnabled = true;
                    tbarStartButton.IsEnabled = true;
                }
                else if (lboxSongs.Items.Count == 0)
                {
                    tbarPresenterView.IsEnabled = false;
                    tbarStartButton.IsEnabled = false;
                }

                return countSongs;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, getStringByName("Window.Title"), MessageBoxButton.OK, MessageBoxImage.Error);

                return 0;
            }
        }

        private void RemoveSong(string songSelected)
        {
            try
            {
                int songNumIdx = lboxSongs.Items.IndexOf(songSelected);

                if (!VerifyPresentationIsOpen())
                    return;

                int[] SlideNums = (int[])SongsList[songNumIdx];
                int SlidesToDelete = SlideNums[1] - SlideNums[0] + 1;

                pBar.Value = 0;
                pBar.Maximum = SlidesToDelete + 1; // Extra number to show progress while saving file

                //Create a new instance of our ProgressBar Delegate that points
                // to the ProgressBar's SetValue method.
                UpdateProgressBarDelegate updatePbDelegate =
                    new UpdateProgressBarDelegate(pBar.SetValue);

                for (int i = 0; i < SlidesToDelete; i++)
                {
                    objPresSongs.Slides[SlideNums[0]].Delete();

                    pBar.Value += 1;

                    // Update the Value of the ProgressBar:
                    Dispatcher.Invoke(updatePbDelegate,
                        System.Windows.Threading.DispatcherPriority.Background,
                        new object[] { System.Windows.Controls.ProgressBar.ValueProperty, pBar.Value });

                }

                // Decrease slide numbers in rest of songs in list
                if (lboxSongs.Items.Count - 1 > songNumIdx)
                    for (int i = 0; i < lboxSongs.Items.Count - songNumIdx - 1; i++)
                    {
                        SlideNums = (int[])SongsList[songNumIdx + i + 1];
                        int[] NewSlideNums = { SlideNums[0] - SlidesToDelete, SlideNums[1] - SlidesToDelete };

                        SongsList[songNumIdx + i + 1] = NewSlideNums;
                    }

                SongsList.RemoveAt(songNumIdx);
                lboxSongs.Items.Remove(songSelected);

                if (lboxSongs.Items.Count == 0)
                {
                    tbarRemoveButton.IsEnabled = false;
                    tbarStartButton.IsEnabled = false;
                    tbarPresenterView.IsEnabled = false;
                }

                objPresSongs.SaveAs(Filename, PowerPoint.PpSaveAsFileType.ppSaveAsDefault);

                // Final upgrade to progress bar after saving file
                pBar.Value += 1;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, getStringByName("Window.Title"), MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        void FocusCurrentSong(PowerPoint.SlideShowView slideShow)
        {

            for (int i = 0; i < SongsList.Count; i++)
            {
                int[] SlideNumbers = (int[])SongsList[i];

                if (slideShow.CurrentShowPosition >= SlideNumbers[0] &&
                    slideShow.CurrentShowPosition <= SlideNumbers[1])
                {
                    lboxSongs.SelectedIndex = i;
                    return;
                }
            }
        }

        private void StartSlideShow()
        {
            try
            {
                if (!VerifyPresentationIsOpen())
                    return;

                // Old versions don't have "ShowPresenterView" and the app will crash if this is used
                if (PPTVersion > 12)
                {
                    if ((bool)tbarPresenterView.IsChecked)
                        objPresSongs.SlideShowSettings.ShowPresenterView = MsoTriState.msoTrue;
                    else
                        objPresSongs.SlideShowSettings.ShowPresenterView = MsoTriState.msoFalse;
                }

                PowerPoint.SlideShowView slideShow = GetSlideshowOpen();

                if (lboxSongs.SelectedIndex >= 0)
                {
                    int[] SlideNums = (int[])SongsList[lboxSongs.SelectedIndex];
                    slideShow.GotoSlide(SlideNums[0]);
                }
                else
                    // Slideshow will start from beginning
                    // Move focus to first song in list
                    FocusCurrentSong(slideShow);

                // If more than one selected, change the selection back to just the first one
                if (lboxSongs.SelectedItems.Count > 1)
                {
                    int FirstSelectedIndex = lboxSongs.SelectedIndex;
                    lboxSongs.UnselectAll();
                    lboxSongs.SelectedIndex = FirstSelectedIndex;
                }

                // Update slide number counter
                lblSlideNum.Content = slideShow.CurrentShowPosition.ToString();

                imgStart.Source = new BitmapImage(new Uri(@"/images/Icon_Running_16x.png", UriKind.Relative));

                tbarNextButton.IsEnabled = true;
                tbarPreviousButton.IsEnabled = true;
                tbarStopButton.IsEnabled = true;

                SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
                //SetFocus(new HandleRef(null, Process.GetCurrentProcess().MainWindowHandle)); // not needed

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, getStringByName("Window.Title"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private PowerPoint.SlideShowView GetSlideshowOpen(Boolean Closing = false)
        {
            PowerPoint.SlideShowView slideShow = null;

            // Catch error when user closes slideshow view manually
            try
            {
                if (slideShowWin == null && !Closing)
                    slideShowWin = objPresSongs.SlideShowSettings.Run();

                slideShow = slideShowWin.View;
            }
            catch (Exception ex)
            {
                if (ex.ToString().ToUpper().Contains("OBJECT DOES NOT EXIST"))
                {
                    if (!Closing)
                    {
                        slideShowWin = objPresSongs.SlideShowSettings.Run();
                        slideShow = slideShowWin.View;
                    }
                }
                else
                    throw;
            }

            return slideShow;
        }

        private void AdvanceSlideShow(int countSlides)
        {
            try
            {
                if (!VerifyPresentationIsOpen())
                    return;

                PowerPoint.SlideShowView slideShow = GetSlideshowOpen();

                for (int i = 0; i < Math.Abs(countSlides); i++)
                    if (countSlides > 0 && slideShow.CurrentShowPosition <= objPresSongs.Slides.Count)
                        slideShow.Next();
                    else if (countSlides < 0 && slideShow.CurrentShowPosition > 1)
                        slideShow.Previous();

                // Update slide number counter
                lblSlideNum.Content = slideShow.CurrentShowPosition.ToString();

                // Move focus to song in list
                FocusCurrentSong(slideShow);

                SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
                //SetFocus(new HandleRef(null, Process.GetCurrentProcess().MainWindowHandle)); // not needed

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, getStringByName("Window.Title"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void StopSlideShow(Boolean Exiting = false)
        {
            try
            {
                if (!VerifyPresentationIsOpen() && !Exiting)
                    return;
                else if (Exiting)
                    return;

                PowerPoint.SlideShowView slideShow = GetSlideshowOpen(true);

                if (slideShow != null)
                    slideShow.Exit(); // This improperly closes presentation (file) in Office 2007

                // TEST, close presentation
                //objPresSongs.Close();

                // Reopen presentation if closed by error
                ReOpenPresentation();

                imgStart.Source = new BitmapImage(new Uri(@"/images/Icon_Start_16x.png", UriKind.Relative));

                tbarNextButton.IsEnabled = false;
                tbarPreviousButton.IsEnabled = false;
                tbarStopButton.IsEnabled = false;
                lblSlideNum.Content = "--";

                SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, getStringByName("Window.Title"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnOpenPpt_Click(object sender, RoutedEventArgs e)
        {
            StartPowerpoint();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            ClosePowerpoint();
        }

        private void tbarAddButton_Click(object sender, RoutedEventArgs e)
        {
            AddSlides(lboxSongs.SelectedIndex);
        }

        private void tbarStartButton_Click(object sender, RoutedEventArgs e)
        {
            StartSlideShow();
        }

        private void lboxSongs_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (lboxSongs.SelectedIndex >= 0)
                tbarRemoveButton.IsEnabled = true;
            else
                tbarRemoveButton.IsEnabled = false;
        }

        private void tbarStopButton_Click(object sender, RoutedEventArgs e)
        {
            StopSlideShow();
        }

        private void tbarRemoveButton_Click(object sender, RoutedEventArgs e)
        {
            /*if (lboxSongs.SelectedItems.Count > 0)
                foreach (object SongSelected in lboxSongs.SelectedItems)
                    RemoveSong(SongSelected);*/

            // Reuse delete procedure used when pressing Delete key
            System.Windows.Input.KeyEventArgs kea = new System.Windows.Input.KeyEventArgs(Keyboard.PrimaryDevice,
                new HwndSource(0, 0, 0, 0, 0, "", IntPtr.Zero), 0, Key.Delete);
            lboxSongs_KeyDown(sender, kea);
        }

        private void tbarNextButton_Click(object sender, RoutedEventArgs e)
        {
            AdvanceSlideShow(1);
        }

        private void tbarPreviousButton_Click(object sender, RoutedEventArgs e)
        {
            AdvanceSlideShow(-1);
        }

        private void lboxSongs_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key.Equals(Key.Delete) && lboxSongs.SelectedItems.Count > 0)
            {
                string[] strSelectedItems = new string[lboxSongs.SelectedItems.Count];
                lboxSongs.SelectedItems.CopyTo(strSelectedItems, 0);

                for (var i = 0; i < strSelectedItems.Length; i++)
                    RemoveSong(strSelectedItems[i]);
            }
        }

        //Create a Delegate that matches 
        //the Signature of the ProgressBar's SetValue method
        private delegate void UpdateProgressBarDelegate(
                System.Windows.DependencyProperty dp, Object value);

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            // If Slideshow or Presentation open, close them before exiting
            if (tbarStopButton.IsEnabled)
                StopSlideShow(true);

            if (_closePptMenuItem.IsEnabled)
                ClosePowerpoint(true);

            // Exit application
            System.Windows.Application.Current.Shutdown();
        }

        private void tbarPresenterView_Checked(object sender, RoutedEventArgs e)
        {
            tbarPresenterView.ToolTip = "[ON] Toggle Presenter View On/Off";

        }

        private void tbarPresenterView_Unchecked(object sender, RoutedEventArgs e)
        {
            tbarPresenterView.ToolTip = "[OFF] Toggle Presenter View On/Off";
        }

        private void _ccMenuItem_Click(object sender, RoutedEventArgs e)
        {
            currentSongBookPath = CCPath;
            lblSongBook.Content = CCName;
            _ceMenuItem.IsChecked = false;
        }

        private void _ceMenuItem_Click(object sender, RoutedEventArgs e)
        {
            currentSongBookPath = CEPath;
            lblSongBook.Content = CEName;
            _ccMenuItem.IsChecked = false;
        }
    }
}