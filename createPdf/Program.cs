using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ZXing;

namespace pdftest
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new PDF document
            PdfDocument document = new PdfDocument();
            document.Info.Title = "電子發票";

            // Create an empty page
            PdfPage page = document.AddPage();

            // 設定電子發票長寬
            var width = new XUnit();
            width.Centimeter = 5.7;

            var height = new XUnit();
            height.Centimeter = 9;

            page.Width = width;
            page.Height = height;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Create a font
            PrivateFontCollection pfcFonts = new System.Drawing.Text.PrivateFontCollection();

            string strFontPath2 = @"C:/Windows/Fonts/times.ttf";
            pfcFonts.AddFontFile(strFontPath2);
            string strFontPath1 = @"C:/Windows/Fonts/kaiu.ttf";
            pfcFonts.AddFontFile(strFontPath1);
            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            //  資料呈顯
            DrawWord(gfx, pfcFonts.Families[1], options, page, @"營業人企業識別標章", 10, 0, 30, XStringFormats.TopCenter);
            DrawWord(gfx, pfcFonts.Families[1], options, page, @"電子發票證明聯", 18, 0, 45, XStringFormats.TopCenter);
            DrawWord(gfx, pfcFonts.Families[1], options, page, @"102年05-06月", 23, 0, 70, XStringFormats.TopCenter);
            DrawWord(gfx, pfcFonts.Families[1], options, page, @"AB-11223344", 23, 0, 95, XStringFormats.TopCenter);
            DrawWord(gfx, pfcFonts.Families[1], options, page, @"2013-05-23 11:00:30", 9, 10, 125, XStringFormats.TopLeft);
            DrawWord(gfx, pfcFonts.Families[1], options, page, @"隨機碼 9999   總計340", 9, 10, 137, XStringFormats.TopLeft);
            DrawWord(gfx, pfcFonts.Families[1], options, page, @"賣方01234567", 9, 9, 149, XStringFormats.TopLeft);

            // 產生Code128
            var barCodeW = new XUnit();
            barCodeW.Centimeter = 5.1;
            var barCodeH = new XUnit();
            barCodeH.Centimeter = 0.6;

            BarcodeWriter bw = new BarcodeWriter();
            bw.Format = BarcodeFormat.CODABAR;
            bw.Options.Width = (int) barCodeW.Centimeter;
            bw.Options.Height = (int) barCodeH.Centimeter;
            bw.Options.PureBarcode = true;
            Bitmap code = bw.Write("999999999");

            gfx.DrawImage(code, 3, 165, barCodeW, barCodeH);

            var qrcodeSize = 2.0;
            // Left QRCode
            var w = new XUnit();
            w.Centimeter = qrcodeSize;
            var h = new XUnit();
            h.Centimeter = qrcodeSize;
            gfx.DrawImage(GenerateQRCode("12345678", Color.Black, Color.White), 10, 190, w, h);

            // Right QRCode
            w = new XUnit();
            w.Centimeter = qrcodeSize;
            h = new XUnit();
            h.Centimeter = qrcodeSize;
            var left = new XUnit();
            left.Centimeter = 3;

            gfx.DrawImage(GenerateQRCode("0999999", Color.Black, Color.White), left, 190, w, h);

            // Save the document...
            const string filename = "HelloWorld.pdf";
            document.Save(filename);
            // ...and start a viewer.
            //Process.Start(filename);
        }

        public static void DrawWord(XGraphics gfx, FontFamily pfcFonts, XPdfFontOptions options, PdfPage page, string content, int fontSize, int x, int y, XStringFormat format)
        {
            var font = new XFont(pfcFonts, fontSize, XFontStyle.Regular, options);
            gfx.DrawString(content, font, XBrushes.Black, new XRect(x, y, page.Width, page.Height), format);
        }

        public static Bitmap GenerateQRCode(string text, Color DarkColor, Color LightColor)
        {
            Gma.QrCodeNet.Encoding.QrEncoder Encoder = new Gma.QrCodeNet.Encoding.QrEncoder(Gma.QrCodeNet.Encoding.ErrorCorrectionLevel.H);
            Gma.QrCodeNet.Encoding.QrCode Code = Encoder.Encode(text);
            Bitmap TempBMP = new Bitmap(Code.Matrix.Width, Code.Matrix.Height);
            for (int X = 0; X <= Code.Matrix.Width - 1; X++)
            {
                for (int Y = 0; Y <= Code.Matrix.Height - 1; Y++)
                {
                    if (Code.Matrix.InternalArray[X, Y])
                        TempBMP.SetPixel(X, Y, DarkColor);
                    else
                        TempBMP.SetPixel(X, Y, LightColor);
                }
            }
            return TempBMP;
        }
    }
}
