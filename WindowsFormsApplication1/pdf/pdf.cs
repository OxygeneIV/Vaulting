using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
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


            MergeMultiplePDFIntoSinglePDF(Path.Combine( Form1.mergedresultsFolder,"All_Results.pdf"), pdfs);
        }
    }
}
