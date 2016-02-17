using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Drawing;

namespace DESign_WordAddIn
{
    class ImageOperations
    {
        public void ConvertToBlackAndWhite(Bitmap sourceImage)
        {

            using (Graphics gr = Graphics.FromImage(sourceImage)) // SourceImage is a Bitmap object
            {
                var gray_matrix = new float[][] {
                new float[] { 0.299f, 0.299f, 0.299f, 0, 0 },
                new float[] { 0.587f, 0.587f, 0.587f, 0, 0 },
                new float[] { 0.114f, 0.114f, 0.114f, 0, 0 },
                new float[] { 0,      0,      0,      1, 0 },
                new float[] { 0,      0,      0,      0, 1 }
            };

                var ia = new System.Drawing.Imaging.ImageAttributes();
                ia.SetColorMatrix(new System.Drawing.Imaging.ColorMatrix(gray_matrix));
                ia.SetThreshold(0.8f); // Change this threshold as needed
                var rc = new Rectangle(0, 0, sourceImage.Width, sourceImage.Height);
                gr.DrawImage(sourceImage, rc, 0, 0, sourceImage.Width, sourceImage.Height, GraphicsUnit.Pixel, ia);
            }
        }

        public void Resize(Bitmap btm, int newWidth, int newHeight)
        {
            Bitmap newImage = new Bitmap(newWidth, newHeight);
            using (Graphics gr = Graphics.FromImage(newImage))
            {
                gr.SmoothingMode = SmoothingMode.HighQuality;
                gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                gr.PixelOffsetMode = PixelOffsetMode.HighQuality;
                gr.DrawImage(btm, new Rectangle(0, 0, newWidth, newHeight));
            }
        }

        public Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        public void createPDF(Bitmap bmp, string saveAsPath)
        {
            
            PdfDocument doc = new PdfDocument();
            PdfPage page = doc.AddPage(new PdfPage()); //new
            //doc.Pages.Add(new PdfPage()); //old

            page.Width = bmp.Height;
            page.Height = bmp.Width;
            page.Orientation = PdfSharp.PageOrientation.Landscape;
            XGraphics xgr = XGraphics.FromPdfPage(page);
            XImage image = XImage.FromGdiPlusImage(bmp);
            
            xgr.DrawImage(image, 0, 0);
            doc.Save(saveAsPath);
            doc.Close();
            
        }

        public void scalePDF(Bitmap bmp, string saveAsPath)
        {

            PdfDocument doc = new PdfDocument();
            PdfPage page = doc.AddPage(new PdfPage()); //new
            //doc.Pages.Add(new PdfPage()); //old

            page.Width = bmp.Height;
            page.Height = bmp.Width;
            page.Orientation = PdfSharp.PageOrientation.Landscape;
            XGraphics xgr = XGraphics.FromPdfPage(page);
            XImage image = XImage.FromGdiPlusImage(bmp);

            xgr.DrawImage(image, 0, 0);
            doc.Save(saveAsPath);
            doc.Close();

        }
    }

}
