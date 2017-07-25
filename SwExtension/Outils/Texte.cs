using LogDebugging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using System.Globalization;

namespace Outils
{
    public static class Texte
    {
        public static Image eConvertirEnBmp(this String text, int htImage, int lgImage)
        {
            try
            {
                int H = htImage;
                int W = lgImage;
                Font F = new Font(new FontFamily("Calibri"), (float)(Math.Ceiling(H / 1.34) + 1), GraphicsUnit.Pixel);

                //first, create a dummy bitmap just to get a graphics object
                Image Img = new Bitmap(1, 1);
                Graphics G = Graphics.FromImage(Img);

                //measure the string to see how big the image needs to be
                SizeF TextSize = G.MeasureString(text, F);

                //free up the dummy image and old graphics object
                Img.Dispose();
                G.Dispose();

                //create a new image of the right size
                Img = new Bitmap(W, H, PixelFormat.Format16bppRgb555);

                G = Graphics.FromImage(Img);

                G.CompositingQuality = CompositingQuality.HighQuality;
                G.InterpolationMode = InterpolationMode.HighQualityBicubic;
                G.SmoothingMode = SmoothingMode.AntiAlias;
                G.TextRenderingHint = TextRenderingHint.AntiAlias;

                //paint the background
                G.Clear(Color.White);

                //create a brush for the text
                Brush TextBrush = new SolidBrush(Color.Black);

                G.DrawString(text, F, TextBrush, (float)((W - TextSize.Width) * 0.5), (float)((H - TextSize.Height) * 0.5));

                G.Save();

                TextBrush.Dispose();
                G.Dispose();

                return Img;
            }
            catch
            {
                Log.Write("Erreur Image");
                return null;
            }
        }

        public static Image eConvertirEnBmp(this String text, int htImage)
        {
            return eConvertirEnBmp(text, htImage, htImage);
        }

        public static int eCountNewLine(this String s)
        {
            int n = 0;
            foreach (var c in s)
            {
                if (c == '\n') n++;
            }
            return n + 1;
        }

        public static bool eIsInteger(this String text)
        {
            int retNum;

            bool isNum = int.TryParse(Convert.ToString(text), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }

        public static bool eIsDouble(this String text)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(text), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }

        public static string eRepeter(this String Chaine, int Nb)
        {
            if (Nb > 0)
                return string.Join(Chaine, new string[Nb]);
            else
                return "";
        }

        public static Encoding eGetEncoding(string filename)
        {
            // Read the BOM
            var bom = new byte[4];
            using (var file = new FileStream(filename, FileMode.Open)) file.Read(bom, 0, 4);

            Encoding e = Encoding.ASCII;

            // Analyze the BOM
            if (bom[0] == 0x2b && bom[1] == 0x2f && bom[2] == 0x76) e = Encoding.UTF7;
            if (bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf) e = Encoding.UTF8;
            if (bom[0] == 0xff && bom[1] == 0xfe) e = Encoding.Unicode; //UTF-16LE
            if (bom[0] == 0xfe && bom[1] == 0xff) e = Encoding.BigEndianUnicode; //UTF-16BE
            if (bom[0] == 0 && bom[1] == 0 && bom[2] == 0xfe && bom[3] == 0xff) e = Encoding.UTF32;
            return e;
        }

        public static String eGetSafeFilename(this String filename, String remplace = "_")
        {

            return string.Join(remplace, filename.Split(Path.GetInvalidFileNameChars()));

        }

        public static Boolean eIsLike(this String text, String pattern, Boolean caseSensitiv = true)
        {
            try
            {
                CompareMethod cp = CompareMethod.Text;
                if (caseSensitiv)
                    cp = CompareMethod.Binary;

                return LikeOperator.LikeString(text, pattern, cp);
            }
            catch (Exception e)
            { Log.Message(e); }

            return false;
        }

        public static String IsRefAndNotEmpty(this String text, String valeurSiNon)
        {
            if (String.IsNullOrWhiteSpace(text))
                return valeurSiNon;

            return text;
        }

        public static String RemoveDiacritics(this String stIn)
        {
            string stFormD = stIn.Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder();

            for (int ich = 0; ich < stFormD.Length; ich++)
            {
                UnicodeCategory uc = CharUnicodeInfo.GetUnicodeCategory(stFormD[ich]);
                if (uc != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(stFormD[ich]);
                }
            }

            return (sb.ToString().Normalize(NormalizationForm.FormC));
        }
    }

    public class WindowsStringComparer : IComparer<string>
    {
        private ListSortDirection _Dir = ListSortDirection.Ascending;

        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        static extern int StrCmpLogicalW(String x, String y);

        public WindowsStringComparer() { }

        public WindowsStringComparer(ListSortDirection Dir)
        {
            _Dir = Dir;
        }

        public int Compare(string x, string y)
        {
            if (_Dir == ListSortDirection.Ascending)
                return StrCmpLogicalW(x, y);
            else
                return StrCmpLogicalW(y, x);
        }

    }
}
