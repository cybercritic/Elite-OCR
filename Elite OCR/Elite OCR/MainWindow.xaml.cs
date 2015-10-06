using Elite_OCR.Classes;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Elite_OCR
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private BitmapImage currentImage;
        private Bitmap currentBitmap;

        private static Thread screenMonitor;
        private string lastProcessedScreen;

        public MainWindow()
        {
            MainWindow.screenMonitor = new Thread(MonitorThread);

            InitializeComponent();

            this.cbMonitorScreen.IsChecked = Properties.Settings.Default.MonitorScreenshots;
            this.cbSaveCSV.IsChecked = Properties.Settings.Default.WriteCSV;
        }

        private void btTessPath_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.tbTessPath.Text = dialog.SelectedPath;
                Properties.Settings.Default.TesseractPath = dialog.SelectedPath;
                Properties.Settings.Default.Save();
            }
        }

        private void btEDScreensPath_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.tbEDScreens.Text = dialog.SelectedPath;
                Properties.Settings.Default.EDScreenPath = dialog.SelectedPath;
                Properties.Settings.Default.Save();
            }
        }

        private void btProcessing_Click(object sender, RoutedEventArgs e)
        {
            if(Properties.Settings.Default.CropMargin == new Thickness(0,0,0,0))
            {
                this.winMain.Width = 900;
                this.winMain.Height = 700;

                ImageShow showExample = new ImageShow();
                showExample.imImage.Source = this.ToBitmapImage(Properties.Resources.example);
                showExample.Width = Properties.Resources.example.Width;
                showExample.Height = Properties.Resources.example.Height;

                showExample.ShowDialog();
            }

            this.cnSettings.IsEnabled = false;
            this.cnSettings.Visibility = System.Windows.Visibility.Collapsed;

            this.cnResult.IsEnabled = false;
            this.cnResult.Visibility = System.Windows.Visibility.Collapsed;

            this.cnProcessing.IsEnabled = true;
            this.cnProcessing.Visibility = System.Windows.Visibility.Visible;
        }

        private void btSettings_Click(object sender, RoutedEventArgs e)
        {
            this.tbTessPath.Text = Properties.Settings.Default.TesseractPath;
            this.tbEDScreens.Text = Properties.Settings.Default.EDScreenPath;

            this.cnProcessing.IsEnabled = false;
            this.cnProcessing.Visibility = System.Windows.Visibility.Collapsed;

            this.cnResult.IsEnabled = false;
            this.cnResult.Visibility = System.Windows.Visibility.Collapsed;

            this.cnSettings.IsEnabled = true;
            this.cnSettings.Visibility = System.Windows.Visibility.Visible;
        }

        private void btLoadImage_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog ofdImage = new System.Windows.Forms.OpenFileDialog();
            ofdImage.Filter = "BMP files|*.bmp";
            ofdImage.InitialDirectory = Properties.Settings.Default.EDScreenPath;

            myCapture = true;
            if (ofdImage.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Uri tmp = new Uri(ofdImage.FileName);
                Bitmap bit = this.MakeGrayscale3(new Bitmap(ofdImage.FileName));

                this.currentImage = this.ToBitmapImage(bit);
                this.imImage.Source = this.currentImage;

                this.imImage.UpdateLayout();

                this.cropInit();
            }
        }

        private void btConvertImage_Click(object sender, RoutedEventArgs e)
        {
            // Create a CroppedBitmap based off of a xaml defined resource.
            //CroppedBitmap cb = new CroppedBitmap(this.currentImage, new Int32Rect(0, 0, newWidth, (int)this.currentImage.Height));
            //this.currentImage 

            try
            {
                this.imImage.Source = this.processImage(this.BitmapImage2Bitmap(this.currentImage));
            }
            catch { System.Windows.MessageBox.Show("Something went wrong, sorry.\nScreenshot resolution could be too high or bounds could be invalid.", "Error processing image.", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private RectangleF calcRawMargin(SizeF bitmap)
        {
            Thickness stored = Properties.Settings.Default.CropMargin;
            RectangleF result = new RectangleF();

            if (stored != null && stored != new Thickness(0, 0, 0, 0))
            {
                SizeF scale = new SizeF((float)bitmap.Width, (float)bitmap.Height);
                scale.Width /= Properties.Settings.Default.CropSize.Width;
                scale.Height /= Properties.Settings.Default.CropSize.Height;

                Thickness nt = new Thickness(stored.Left, stored.Top, stored.Right, stored.Bottom);

                PointF p1 = new PointF((float)(nt.Left * scale.Width), (float)(nt.Top * scale.Height));
                PointF p2 = new PointF((float)(nt.Right * scale.Width), (float)(nt.Bottom * scale.Height));

                result = new RectangleF((float)p1.X, (float)p1.Y,(float)(bitmap.Width - (p2.X + p1.X)),(float)(bitmap.Height - (p2.Y + p1.Y)));
            }

            return result;
        }

        private RectangleF calcMargin(SizeF bitmap)
        {
            PointF scale = new PointF(bitmap.Width / (float)this.imImage.ActualWidth, bitmap.Height / (float)this.imImage.ActualHeight);
            RectangleF newSize = new RectangleF();


            System.Windows.Point p1 = this.gdPicPre.TranslatePoint(new System.Windows.Point(this.cropRectangle.Margin.Left, this.cropRectangle.Margin.Top), this.imImage);
            p1.X = this.Clamp((float)p1.X, 0, (float)imImage.RenderSize.Width);
            p1.Y = this.Clamp((float)p1.Y, 0, (float)imImage.RenderSize.Height);

            newSize.X = (float)p1.X;
            newSize.Y = (float)p1.Y;

            System.Windows.Point p2 = this.gdPicPre.TranslatePoint(new System.Windows.Point(this.cropRectangle.Margin.Right, this.cropRectangle.Margin.Bottom), this.imImage);
            p2.X = this.Clamp((float)p2.X, 0, (float)imImage.RenderSize.Width);
            p2.Y = this.Clamp((float)p2.Y, 0, (float)imImage.RenderSize.Height);

            newSize.Width = (float)((this.imImage.RenderSize.Width - p2.X) - newSize.X);
            newSize.Height = (float)((this.imImage.RenderSize.Height - p2.Y) - newSize.Y);

            newSize.X *= scale.X;
            newSize.Y *= scale.Y;
            newSize.Width *= scale.X;
            newSize.Height *= scale.Y;

            Thickness marginP = new Thickness(p1.X, p1.Y, p2.X, p2.Y);

            if (Properties.Settings.Default.CropMargin != marginP)
            {
                Properties.Settings.Default.CropMargin = marginP;
                Properties.Settings.Default.CropSize = new SizeF((float)this.imImage.RenderSize.Width,(float)this.imImage.RenderSize.Height);
                Properties.Settings.Default.Save();
            }

            return newSize;
        }

        private BitmapImage processImage(Bitmap bitmap)
        {
            //newSize.Width = this.Clamp(newSize.Width, 0, bitmap.Width);
            //newSize.Height = this.Clamp(newSize.Height, 0, bitmap.Height);

            //int newWidth = (int)((float)bitmap.Width * ((float)this.cropRectangle.Width / (float)this.imImage.ActualWidth));
            
            Bitmap result = bitmap.Clone(this.calcRawMargin(bitmap.Size), bitmap.PixelFormat);
            result = new Bitmap(result, new System.Drawing.Size(result.Width * 2, result.Height * 2));

            result = this.covertToTessFrendly(result);

            //result = new Bitmap(result, new System.Drawing.Size(result.Width * 2, result.Height * 2));

            this.currentImage = ToBitmapImage(result);
            this.currentBitmap = result;
            return this.currentImage;
        }

        private Bitmap BitmapImage2Bitmap(BitmapImage bitmapImage)
        {
            // BitmapImage bitmapImage = new BitmapImage(new Uri("../Images/test.png", UriKind.Relative));

            using (MemoryStream outStream = new MemoryStream())
            {
                BitmapEncoder enc = new BmpBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(bitmapImage));
                enc.Save(outStream);
                System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(outStream);

                return new Bitmap(bitmap);
            }
        }

        private BitmapImage ToBitmapImage(Bitmap bitmap)
        {
            using (var memory = new MemoryStream())
            {
                bitmap.Save(memory, ImageFormat.Bmp);
                memory.Position = 0;

                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();

                return bitmapImage;
            }
        }

        private void btProcessImage_Click(object sender, RoutedEventArgs e)
        {
            TesseractOCR tess = new TesseractOCR(Properties.Settings.Default.TesseractPath);

            string result = "";
            try
            {
                result = tess.OCRFromBitmap(this.currentBitmap);
                result = result.Replace(",", string.Empty);
                result = result.Replace(".", string.Empty);
                result = result.Replace("-", "0");
                result = result.Replace("'", "0");

                string[] lines = result.Split('\n');
                string resultf = "";

                int cn = 0;
                foreach (string line in lines)
                {
                    string[] words = line.Split(' ');
                    string nline = "";

                    if (cn++ < 3)
                    {
                        resultf += line + "\n";
                        continue;
                    }
                    if (words.Length < 3)
                        continue;

                    int c = -1;
                    bool r = int.TryParse(words[1], out c);
                    bool v = int.TryParse(words[2], out c);

                    for (int i = words.Length - 1; i >= 0; i--)
                    {
                        string word = words[i].Trim();
                        bool addtab = (i != 0) && (i != 1 || r) && (i != 2 || v);

                        nline = nline.Insert(0, (addtab ? ";" : i != 0 ? " " : "") + word);
                    }
                    resultf += nline + "\n";
                }
                result = resultf;
            }
            catch { }//MessageBox.Show("Something went wrong, sorry.", "Error during OCR", MessageBoxButton.OK, MessageBoxImage.Error); }

            this.tbResult.Text = result.ToUpper();
        }

        private void btResult_Click(object sender, RoutedEventArgs e)
        {
            this.cnSettings.IsEnabled = false;
            this.cnSettings.Visibility = System.Windows.Visibility.Collapsed;

            this.cnProcessing.IsEnabled = false;
            this.cnProcessing.Visibility = System.Windows.Visibility.Collapsed;

            this.cnResult.IsEnabled = true;
            this.cnResult.Visibility = System.Windows.Visibility.Visible;
        }

        private bool myCapture = false;

        private void selectionRectangle_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && !myCapture)
            {
                var mousePosition = e.GetPosition(sender as UIElement);
                
                Thickness margin = this.cropRectangle.Margin;
                margin.Top = mousePosition.Y;
                margin.Left = mousePosition.X;
                
                this.cropRectangle.Margin = margin;
                this.calcMargin(new SizeF((float)this.imImage.Source.Width, (float)this.imImage.Source.Height));
            }
            else if (e.RightButton == MouseButtonState.Pressed && !IsMouseCaptured)
            {
                var mousePosition = e.GetPosition(sender as UIElement);
              
                Thickness margin = this.cropRectangle.Margin;
                margin.Bottom = this.gdPicPre.ActualHeight - mousePosition.Y;
                margin.Right = this.gdPicPre.ActualWidth - mousePosition.X;
                this.cropRectangle.Margin = margin;
                this.calcMargin(new SizeF((float)this.imImage.Source.Width, (float)this.imImage.Source.Height));
            }

            if (e.LeftButton == MouseButtonState.Released && myCapture)
                myCapture = false;
            if (e.RightButton == MouseButtonState.Released && myCapture)
                myCapture = false;

        }

        private void continueMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            e.Handled = false;
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private Bitmap covertToTessFrendly(Bitmap original)
        {
            unsafe
            {
                //create an empty bitmap the same size as original
                Bitmap newBitmap = new Bitmap(original.Width, original.Height);

                //lock the original bitmap in memory
                BitmapData originalData = original.LockBits(
                   new System.Drawing.Rectangle(0, 0, original.Width, original.Height),
                   ImageLockMode.ReadOnly, original.PixelFormat);

                //lock the new bitmap in memory
                BitmapData newData = newBitmap.LockBits(
                   new System.Drawing.Rectangle(0, 0, original.Width, original.Height),
                   ImageLockMode.WriteOnly, original.PixelFormat);

                //set the number of bytes per pixel
                int pixelSize = 4;

                for (int y = 0; y < original.Height; y++)
                {
                    //get the data from the original image
                    byte* oRow = (byte*)originalData.Scan0 + (y * originalData.Stride);

                    //get the data from the new image
                    byte* nRow = (byte*)newData.Scan0 + (y * newData.Stride);

                    for (int x = 0; x < original.Width; x++)
                    {
                        //create the grayscale version
                        if ((byte)oRow[x * pixelSize + 0] > 155 && (byte)oRow[x * pixelSize + 1] > 155 && (byte)oRow[x * pixelSize + 2] > 155)
                        {
                            nRow[x * pixelSize + 0] = (byte)(255 - (byte)oRow[x * pixelSize + 0]); ; //A
                            nRow[x * pixelSize + 1] = (byte)(255 - (byte)oRow[x * pixelSize + 1]); ; //R
                            nRow[x * pixelSize + 2] = (byte)(255 - (byte)oRow[x * pixelSize + 2]); ; //G
                            nRow[x * pixelSize + 3] = 255 ; //B
                        }
                        else
                        {
                            nRow[x * pixelSize + 0] = 255; //A
                            nRow[x * pixelSize + 1] = 255; //R
                            nRow[x * pixelSize + 2] = 255; //G
                            nRow[x * pixelSize + 3] = 255; //B
                        }
                    }
                }

                //unlock the bitmaps
                newBitmap.UnlockBits(newData);
                original.UnlockBits(originalData);

                return newBitmap;
            }
        }

        private Bitmap MakeGrayscale3(Bitmap original)
        {
            //create a blank bitmap the same size as original
            Bitmap newBitmap = new Bitmap(original.Width, original.Height);

            //get a graphics object from the new image
            Graphics g = Graphics.FromImage(newBitmap);

            //create the grayscale ColorMatrix
            ColorMatrix colorMatrix = new ColorMatrix(
               new float[][] 
                              {
                                 new float[] {.8f, .8f, .8f, 0, 0},//0.3
                                 new float[] {.59f, .59f, .59f, 0, 0},
                                 new float[] {.11f, .11f, .11f, 0, 0},
                                 new float[] {0, 0, 0, 1, 0},
                                 new float[] {0, 0, 0, 0, 1}
                              });

            //create some image attributes
            ImageAttributes attributes = new ImageAttributes();

            //set the color matrix attribute
            attributes.SetColorMatrix(colorMatrix);

            //draw the original image on the new image
            //using the grayscale color matrix
            g.DrawImage(original, new System.Drawing.Rectangle(0, 0, original.Width, original.Height),
               0, 0, original.Width, original.Height, GraphicsUnit.Pixel, attributes);

            //dispose the Graphics object
            g.Dispose();
            return newBitmap;
        }

        private void loadImage(string filename)
        {
            Uri tmp = new Uri(filename);
            Bitmap bit = this.MakeGrayscale3(new Bitmap(filename));

            this.currentImage = this.ToBitmapImage(bit);
            this.imImage.Source = this.currentImage;

            this.imImage.UpdateLayout();

            this.cropInit();
        }

        private void autoProcess()
        {
            try { this.imImage.Source = this.processImage(this.BitmapImage2Bitmap(this.currentImage)); }
            catch { }
        }

        private void saveCSV()
        {
            try
            {
                DateTime time = DateTime.UtcNow;
                string station = this.tbResult.Text.Substring(0, this.tbResult.Text.IndexOf("\n"));

                string folder = this.makeFolder();
                string filename = station + "_" + time.ToString("yyyy.MM.dd_HH.mm.ss") + ".txt";

                char[] invalid = System.IO.Path.GetInvalidFileNameChars();

                if (filename.IndexOfAny(invalid) >= 0)
                    filename = "unknown_" + time.ToString("yyyy.MM.dd_HH.mm.ss") + ".txt";
                
                File.WriteAllText(folder + filename, tbResult.Text);
            }
            catch { }
        }

        private string makeFolder()
        {
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            folder += "\\cybercritics";
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            folder += "\\critics elite OCR";
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            folder += "\\";

            return folder;
        }

        public void MonitorThread()
        {
            while(true)
            {
                if (!Properties.Settings.Default.MonitorScreenshots || Properties.Settings.Default.EDScreenPath == "" || Properties.Settings.Default.TesseractPath == "")
                    return;

                string[] files = Directory.GetFiles(Properties.Settings.Default.EDScreenPath);

                string newest = "";
                DateTime ntime = new DateTime(0);
                foreach(string file in files)
                {
                    DateTime time = File.GetCreationTimeUtc(file);
                    if(time > ntime)
                    {
                        newest = file;
                        ntime = time;
                    }
                }

                if (newest != this.lastProcessedScreen)
                {
                    Dispatcher.BeginInvoke(new ThreadStart(() => loadImage(newest)));
                    Thread.Sleep(10);
                    Dispatcher.BeginInvoke(new ThreadStart(() => autoProcess()));
                    Thread.Sleep(10);
                    Dispatcher.BeginInvoke(new ThreadStart(() => btProcessImage_Click(this,new RoutedEventArgs())));
                    Thread.Sleep(10);
                    if(Properties.Settings.Default.WriteCSV)
                        Dispatcher.BeginInvoke(new ThreadStart(() => saveCSV()));

                    this.lastProcessedScreen = newest;
                }

                Thread.Sleep(2000);
            }
        }

        public float Clamp(float value, float min, float max)
        {
            return (value < min) ? min : (value > max) ? max : value;
        }

        private void cropInit()
        {
            if (this.imImage.RenderSize.Width == 0)
                    return;

            if (Properties.Settings.Default.CropMargin != null && Properties.Settings.Default.CropMargin != new Thickness(0,0,0,0))
            {
                SizeF scale = new SizeF((float)this.imImage.RenderSize.Width, (float)this.imImage.RenderSize.Height);
                scale.Width /= Properties.Settings.Default.CropSize.Width;
                scale.Height /= Properties.Settings.Default.CropSize.Height;

                Thickness nt = new Thickness(Properties.Settings.Default.CropMargin.Left, Properties.Settings.Default.CropMargin.Top,
                                             Properties.Settings.Default.CropMargin.Right, Properties.Settings.Default.CropMargin.Bottom);

                System.Windows.Point p1 = this.imImage.TranslatePoint(new System.Windows.Point(nt.Left * scale.Width, nt.Top * scale.Height), this.gdPicPre);
                System.Windows.Point p2 = this.imImage.TranslatePoint(new System.Windows.Point(nt.Right * scale.Width, nt.Bottom * scale.Height), this.gdPicPre);

                this.cropRectangle.Margin = new Thickness(p1.X, p1.Y, p2.X, p2.Y);
                this.cropRectangle.UpdateLayout();
            }
        }

        private void gdPicPre_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            this.cropInit();
        }

        private void cbMonitorScreen_Checked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.MonitorScreenshots = true;
            Properties.Settings.Default.Save();

            if (!MainWindow.screenMonitor.IsAlive)
            {
                MainWindow.screenMonitor = new Thread(MonitorThread);
                MainWindow.screenMonitor.Start();
            }
        }

        private void cbMonitorScreen_Unchecked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.MonitorScreenshots = false;
            Properties.Settings.Default.Save();

            if (MainWindow.screenMonitor.IsAlive)
                MainWindow.screenMonitor.Abort();
        }

        private void cbSaveCSV_Checked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.WriteCSV = true;
            Properties.Settings.Default.Save();
        }

        private void cbSaveCSV_Unchecked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.WriteCSV = false;
            Properties.Settings.Default.Save();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (MainWindow.screenMonitor != null)
                MainWindow.screenMonitor.Abort();
        }

        private void btExample_Click(object sender, RoutedEventArgs e)
        {
            ImageShow showExample = new ImageShow();
            showExample.imImage.Source = this.ToBitmapImage(Properties.Resources.example);
            showExample.Width = Properties.Resources.example.Width;
            showExample.Height = Properties.Resources.example.Height;

            showExample.ShowDialog();
        }

    }
}
