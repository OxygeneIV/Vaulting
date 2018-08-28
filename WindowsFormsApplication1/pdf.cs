using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.IO;

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
            }
            document.Options.CompressContentStreams = true;
            document.Options.NoCompression = false;
            // In the final stage, all documents are merged and save in your output file path.  
            document.Save(outputFilePath);
        }

        public static void Merge(string folder)
        {
            string[] pdfs = Directory.GetFiles(folder,"*.pdf");
            
            MergeMultiplePDFIntoSinglePDF(Path.Combine( Form1.mergedresults,"CombinedResults.pdf"), pdfs);
        }
    }
}
