using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;

namespace WindowsFormsApplication1
{

  public static class pdf
    {

        private static void MergeMultiplePDFIntoSinglePDF(string outputFilePath, string[] pdfFiles)
        {
            var now = DateTime.Now;
             PdfDocument document = new PdfDocument();
            foreach (string pdfFile in pdfFiles)
            {
                PdfDocument inputPDFDocument = PdfReader.Open(pdfFile, PdfDocumentOpenMode.Import);
                document.Version = inputPDFDocument.Version;
             
                foreach (PdfPage page in inputPDFDocument.Pages)
                {
                    document.AddPage(page);
                }
                // When document is add in pdf document remove file from folder  
                //System.IO.File.Delete(pdfFile);
            }
            // Set font for paging  
            XFont font = new XFont("Verdana", 9);
            XBrush brush = XBrushes.Black;
            // Create variable that store page count  
            string noPages = document.Pages.Count.ToString();
            // Set for loop of document page count and set page number using DrawString function of PdfSharp  
            for (int i = 0; i < document.Pages.Count; ++i)
            {
                PdfPage page = document.Pages[i];
              
              // Make a layout rectangle.  
                XRect layoutRectangle = new XRect(240 /*X*/ , page.Height - font.Height - 10 /*Y*/ , page.Width /*Width*/ , font.Height /*Height*/ );
                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                {
                    gfx.DrawString($" {now:F} -  Page " + (i + 1).ToString() + " of " + noPages, font, brush, layoutRectangle, XStringFormats.Center);
                }



            //using (XGraphics gfx = XGraphics.FromPdfPage(page))
            //{
            //  var xim = XImage.FromFile(ridsport);
            //  gfx.ScaleTransform(0.4);
            //  gfx.DrawImage(xim, new Point(120, 10));
            //}
            //  using (XGraphics gfx = XGraphics.FromPdfPage(page))
            //  {
            //    var xim = XImage.FromFile(complogo);
            //    gfx.ScaleTransform(0.15);
            //    gfx.DrawImage(xim, new Point(800, 10));
            //  }

            //  using (XGraphics gfx = XGraphics.FromPdfPage(page))
            //  {
            //    var xim = XImage.FromFile(datelogo);
            //    gfx.ScaleTransform(0.3);
            //    gfx.DrawImage(xim, new Point(550, 30));
            //  }

            //  using (XGraphics gfx = XGraphics.FromPdfPage(page))
            //  {
            //    var xim = XImage.FromFile(sponsorlogo);
            //    gfx.ScaleTransform(0.3);
            //    gfx.DrawImage(xim,new Point(2000,30));
            //   }

            //  using (XGraphics gfx = XGraphics.FromPdfPage(page))
            //  {
            //    var xim = XImage.FromFile(preliminary);
            //    gfx.ScaleTransform(0.5);
            //    gfx.DrawImage(xim, new Point(1300, 140));
            //  }

      }
      document.Options.CompressContentStreams = true;
            document.Options.NoCompression = false;
            // In the final stage, all documents are merged and save in your output file path.  
            document.Save(outputFilePath);
        }

        public static void Merge(string folder)
        {
            //string[] pdfs = Directory.GetFiles(folder,"*.pdf");

          var files = Directory.GetFiles(folder, "*.pdf").ToList();

          // Adjust the order for SM
          files.Sort(new PDFtoHTML.Comparer());

          string[] pdfs = files.ToArray();


            MergeMultiplePDFIntoSinglePDF(Path.Combine( Form1.mergedresultsFolder,"CombinedResults.pdf"), pdfs);
        }
    }
}
