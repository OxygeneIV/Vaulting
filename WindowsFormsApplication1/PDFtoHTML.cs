using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
//using Bytescout.PDF2HTML;
//using EvoPdf.PdfToHtml;
using FluentFTP;
//using RasterEdge.Imaging.Basic;
//using RasterEdge.XDoc.Converter;
//using RasterEdge.XDoc.PDF;
//using SautinSoft;

namespace WindowsFormsApplication1
{
  public class PDFtoHTML
  {

    public static string MakeFileNameWebSafe(string filename)
    {
      return filename.Replace(",", "-").Replace(" ", "-").Replace("å", "a").Replace("ü","y");
    }

    public static string klassHtml(string iframestring)
    {
      var index = $@"
      <html>
        <body bgcolor=white>
            {iframestring}
        </body>
      </html>
      ";
      return index;
    }

    public class Comparer : IComparer<string>
    {
      public int Compare(string left, string right)
      {
        var leftfilename  = Path.GetFileName(left);
        var rightfilename = Path.GetFileName(right);

        var leftSplit = leftfilename.Split('_').First();
        var rightSplit = rightfilename.Split('_').First();

        var leftfloat = double.Parse(leftSplit, CultureInfo.InvariantCulture.NumberFormat);
        var rightfloat = float.Parse(rightSplit, CultureInfo.InvariantCulture.NumberFormat);

        if (Math.Abs(leftfloat - rightfloat) < 0.001)
        {
          return 0;
        }

        if (leftfloat < rightfloat)
        {
          return -1;
        }
        else
        {
          return 1;
        }
      }


    }

    public static void GenerateHTML()
    {
      var folder = Form1.printedresults;
      var files = Directory.GetFiles(folder, "*.pdf").ToList();

      Dictionary<string,string> pdfLinks = new Dictionary<string, string>();
      List<string> klasshtmlfiles = new List<string>();

      // Adjust the order for SM
      files.Sort(new Comparer());

      foreach (var f in files)
      {
        var PdfFilename = Path.GetFileName(f);
        var shortFile   =  Path.GetFileNameWithoutExtension(f);

        string safeRemotePdfFile = MakeFileNameWebSafe(PdfFilename);
        string safeRemoteFolderName = MakeFileNameWebSafe(shortFile);

        var HTMLfolder = $"{safeRemoteFolderName}/";

        var remotePdfFile  = HTMLfolder + safeRemotePdfFile;
        var remoteHTMLFile = HTMLfolder + safeRemoteFolderName+".html";

        var remotePdfUrl = "http://privat.bahnhof.se/wb653561/" + remotePdfFile;

        pdfLinks[remotePdfUrl] = shortFile;

        var iframeurl =  $@"https://docs.google.com/viewer?url="+ remotePdfUrl + "&embedded=true";
        var iframe = $@"<iframe src=""{iframeurl}"" style=""width:100%; height:100%;"" ></iframe>";

        var klasshtml = klassHtml(iframe);

        var localHtml = Path.Combine(Form1.mergedresults, "klass.html");
        File.WriteAllText(localHtml, klasshtml);
        klasshtmlfiles.Add(remoteHTMLFile);
        // Create folder and klass.html

        FtpClient client1 = new FtpClient("privat.bahnhof.se") { Credentials = new NetworkCredential("wb653561", "foo123") };
        client1.Connect();
        client1.UploadFile(f,         remotePdfFile, createRemoteDir: true);
        client1.UploadFile(localHtml, remoteHTMLFile, createRemoteDir: true);

        client1.Disconnect();
      }

      var index = @"
      <html>
        <head>
          <title>SM/NM 2018</title>
        </head>
        <body bgcolor=white>
               <h1>SM/NM 2018</h1>
            DATA
        </body>
      </html>
      ";

      int i = 0;
      var text = "";
      foreach (KeyValuePair<string,string> kvp in pdfLinks)
      {
        var pdfurl = kvp.Key;
        var klassname = kvp.Value;
        var htmlfile = klasshtmlfiles[i];
        i++;

        text = text + 
               $@"<p>
                     <a href=""{htmlfile}"">{klassname}</a>
               </p>
              ";
        
      }

      index = index.Replace("DATA", text);

      var index_html = Path.Combine(Form1.mergedresults, "index.html");

      File.WriteAllText(index_html,index,Encoding.Unicode);

      // create an FTP client
      FtpClient client = new FtpClient("privat.bahnhof.se") {Credentials = new NetworkCredential("wb653561", "foo123")};
      client.Connect();
      client.UploadFile(index_html, "index.html");
      client.Disconnect();
      }
  }
}
