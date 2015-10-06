// usage:
//
// TesseractOCR ocr = TesseractOCR(@"C:\bin\tesseract.exe");
// string result = ocr.OCRFromBitmap(bmp);
// textBox1.Text = result;
//

using System;
using System.IO;
using System.Diagnostics;
using System.Drawing;
using System.Text.RegularExpressions;

namespace Elite_OCR.Classes
{
    public class TesseractOCR
    {
        private string commandpath;
        private string outpath;
        private string tmppath;

        public TesseractOCR(string commandpath)
        {
            this.commandpath = commandpath + "\\tesseract.exe";
            tmppath = System.Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\out.tif";
            outpath = System.Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\out.txt";
        }
        public string analyze(string filename)
        {
            string args = filename + " " + outpath.Replace(".txt", "") + " -l eng -psm 6"; //+ " hocr";// ;
            ProcessStartInfo startinfo = new ProcessStartInfo(commandpath, args);
            startinfo.CreateNoWindow = true;
            startinfo.UseShellExecute = false;
            Process.Start(startinfo).WaitForExit();

            string ret = "";
            using (StreamReader r = new StreamReader(outpath))
            {
                string content = r.ReadToEnd();
                ret = content;
            }
            File.Delete(outpath);
            return ret;
        }
        public string OCRFromBitmap(Bitmap bmp)
        {
            bmp.Save(tmppath, System.Drawing.Imaging.ImageFormat.Tiff);
            string ret = analyze(tmppath);
            File.Delete(tmppath);

            ret = Regex.Replace(ret, @"[^\u0000-\u007F]", string.Empty);

            return ret;
        }
        public string OCRFromFile(string filename)
        {
            return analyze(filename);
        }
    }
}