using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Syncfusion.Presentation;
using Syncfusion.OfficeChartToImageConverter;
using System.Drawing;

namespace PresentationCommandInterface
{
    class PresentationOperation
    {
        private IPresentation presentation;
        public PresentationOperation()
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = Path.Combine(documents, @"AIMobileOffice\MSPowerPoint\presentation.pptx");
            //string filePath = @"C:\Users\Wentao Ruan\Downloads\AIPPT_Main.pptx";

            // Open a presentation at a specified file path. Set file mode to be ReadWrite mode.
            FileStream fileStream = new FileStream(filePath, FileMode.Open);

            presentation = Presentation.Open(fileStream);
        }

        public void GotoImage()
        {
            presentation.ChartToImageConverter = new ChartToImageConverter();

            presentation.ChartToImageConverter.ScalingMode = Syncfusion.OfficeChart.ScalingMode.Best;

            Image[] images = presentation.RenderAsImages(Syncfusion.Drawing.ImageType.Metafile);

            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string fileDir = Path.Combine(documents, @"AIMobileOffice/MSPowerPoint");
            Directory.CreateDirectory(fileDir);
            string filePath = Path.Combine(documents, @"AIMobileOffice/MSPowerPoint");
            int i = 0;
            foreach (Image image in images)
            {
                image.Save(filePath + i.ToString());
                i++;
            }
        }
    }
}
