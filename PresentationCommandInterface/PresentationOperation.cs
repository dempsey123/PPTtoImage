using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Syncfusion.Presentation;
using Syncfusion.OfficeChartToImageConverter;
using System.Drawing; // IMPORTANT: Import System.Drwing libarary.
using Syncfusion.UI;
using System.Drawing.Imaging;

namespace PresentationCommandInterface
{
    class PresentationOperation
    {
        private IPresentation presentation;
        public PresentationOperation()
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = Path.Combine(documents, @"presentation.pptx");
            //string filePath = @"C:\Users\Wentao Ruan\Downloads\AIPPT_Main.pptx";

            // Open a presentation at a specified file path. Set file mode to be ReadWrite mode.
            FileStream fileStream = new FileStream(filePath, FileMode.Open);

            presentation = Presentation.Open(fileStream);
        }

        public void GotoImage()
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string fileDir = Path.Combine(documents, @"AIMobileOffice/MSPowerPoint");
            Directory.CreateDirectory(fileDir);
            string filePath = Path.Combine(documents, @"AIMobileOffice/MSPowerPoint");

            int customWidth = 5000;
            int customHeight = 2800;

            for (int j = 0; j < presentation.Slides.Count; j++)
            {
                Stream stream = presentation.Slides[j].ConvertToImage(Syncfusion.Drawing.ImageFormat.Bmp);
                Bitmap bitmap = new Bitmap(customWidth, customHeight, PixelFormat.Format32bppPArgb);

                Graphics graphic = Graphics.FromImage(bitmap);
                bitmap.SetResolution(graphic.DpiX, graphic.DpiY);

                graphic.DrawImage(System.Drawing.Image.FromStream(stream), new Rectangle(0, 0, bitmap.Width, bitmap.Height));

                bitmap.Save(filePath + j.ToString());
            }

            presentation.Close();

           // presentation.ChartToImageConverter = new ChartToImageConverter();

          //  presentation.ChartToImageConverter.ScalingMode = Syncfusion.OfficeChart.ScalingMode.Best;

           // Image[] images = presentation.RenderAsImages(Syncfusion.Drawing.ImageType.Metafile); //convert format (size)

           
            //int i = 0;
            //foreach (Image image in images)
            //{
            //    image.Save(filePath + i.ToString());
            //    i++;
            //}
        }
    }
}
