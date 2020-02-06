using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
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

        var leftSplit = leftfilename.Split(' ').First();
        var rightSplit = rightfilename.Split(' ').First();

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
      var folder = Form1.printedresultsFolder;
      var files = Directory.GetFiles(folder, "*.pdf").ToList();

      var singlefile = Path.Combine(Form1.mergedresultsFolder, "All_Results.pdf");

      // Test FTP
      var FTPserver = ConfigurationManager.AppSettings["ftpserver"];
      var FTPuser   = ConfigurationManager.AppSettings["ftpuser"];
      var FTPpwd    = ConfigurationManager.AppSettings["ftppwd"];
      var remoteworkingfolder = ConfigurationManager.AppSettings["remoteworkingfolder"];
      var remotepdfurl = ConfigurationManager.AppSettings["remotepdfurl"];

      try
      {
        FtpClient clienttest = new FtpClient(FTPserver) {Credentials = new NetworkCredential(FTPuser, FTPpwd)};
        clienttest.Connect();
        clienttest.SetWorkingDirectory(remoteworkingfolder);
        if (clienttest.DirectoryExists("smnmdummyfolder"))
        {
          clienttest.DeleteDirectory("smnmdummyfolder");
        }

        clienttest.CreateDirectory("smnmdummyfolder");

        if (!clienttest.DirectoryExists("smnmdummyfolder"))
        {
          clienttest.Disconnect();
          throw new Exception($"Failed to test create a folder on ftp server");
        }

        clienttest.Disconnect();
      }
      catch(Exception e)
      {
        throw new Exception($"Failed to publish {e.Message}");
      }

      Dictionary<string,string> pdfLinks = new Dictionary<string, string>();
      List<string> klasshtmlfiles = new List<string>();

      // Adjust the order for SM
      files.Sort(new Comparer());

      if (File.Exists(singlefile))
      {
        files.Insert(0, singlefile);
      }

      foreach (var f in files)
      {
        var PdfFilename = Path.GetFileName(f);
        var shortFile   =  Path.GetFileNameWithoutExtension(f);

        string safeRemotePdfFile = MakeFileNameWebSafe(PdfFilename);
        string safeRemoteFolderName = MakeFileNameWebSafe(shortFile);

        var HTMLfolder = $"{safeRemoteFolderName}/";

        var remotePdfFile  = HTMLfolder + safeRemotePdfFile;
        var remoteHTMLFile = HTMLfolder + safeRemoteFolderName+".html";

        var remotePdfUrl = remotepdfurl + remotePdfFile;

        pdfLinks[remotePdfUrl] = shortFile;

        var iframeurl =  $@"https://docs.google.com/viewer?url="+ remotePdfUrl + "&embedded=true";
        var iframe = $@"<embed src=""{iframeurl}"" style=""width:100%; height:100%;"" ></embed>";
          //var iframeurl = $@"https://docs.google.com/viewer?url=" + remotePdfUrl;
              // var iframe = $@"<embed src=""{iframeurl}#toolbar=0&navpanes=0&scrollbar=0"" style=""width:500px; height:1000px;"" ></embed>";
          iframe = "<a href=" + @""""+remotePdfUrl+@""""+ ">"+shortFile+"</a>";
                var klasshtml = klassHtml(iframe);
                // ditlagd
          pdfLinks[remotePdfUrl] = iframe;

                var localHtml = Path.Combine(Form1.mergedresultsFolder, "klass.html");
        File.WriteAllText(localHtml, klasshtml);
        klasshtmlfiles.Add(remoteHTMLFile); 
        // Create folder and klass.html

        FtpClient client = new FtpClient(FTPserver) { Credentials = new NetworkCredential(FTPuser, FTPpwd) };
        client.Connect();
        client.SetWorkingDirectory(remoteworkingfolder);
        client.UploadFile(f, remotePdfFile, createRemoteDir: true);
        client.UploadFile(localHtml, remoteHTMLFile, createRemoteDir: true);
        client.Disconnect();


        //FtpClient client1 = new FtpClient("privat.bahnhof.se") { Credentials = new NetworkCredential("wb653561", "foo123") };
        //client1.Connect();


        //client1.Disconnect();
      }

      var index = @"
      <html>
        <head>
          <title>Uppsala / Järlåsaskolan 2020-02-08</title>
        </head>
        <body bgcolor=white>
               <h1 align=""center"">Uppsala / Järlåsaskolan 2020-02-08</h1>
            <div align=""center"">
            DATA
           </div>

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

      var index_html = Path.Combine(Form1.mergedresultsFolder, "index.html");

      File.WriteAllText(index_html,index,Encoding.Unicode);

      // create an FTP client

      FtpClient client1 = new FtpClient(FTPserver) { Credentials = new NetworkCredential(FTPuser, FTPpwd) };
      client1.Connect();
      client1.SetWorkingDirectory(remoteworkingfolder);
      client1.UploadFile(index_html, "index.html");
      client1.Disconnect();


      //FtpClient client1 = new FtpClient("privat.bahnhof.se") {Credentials = new NetworkCredential("wb653561", "foo123")};
      //client1.Connect();
      //client1.UploadFile(index_html, "index.html");
      //client1.Disconnect();
      }
  }
}
