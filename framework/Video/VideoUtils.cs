using System.Collections.Generic;

namespace Framework.Video
{
    public static class VideoUtils
    {
        //public static List<string> Split(string filePath, int videoLengthInMinutes)
        //{
        //    var newFile1 = new MediaFile(filePath+"_01.mp4");
        //    var newFile2 = new MediaFile(filePath+"_02.mp4");
        //    var files = new List<string>();
        //    using (var engine = new Engine())
        //    {
        //        var orig = new MediaFile(filePath);
        //        engine.GetMetadata(orig);
        //        var duration = orig.Metadata.Duration;
        //        var options = new ConversionOptions();
        //        options.CutMedia(TimeSpan.FromSeconds(0),TimeSpan.FromMinutes(videoLengthInMinutes));
        //        engine.Convert(orig, newFile1, options);
        //        options.CutMedia(TimeSpan.FromMinutes(videoLengthInMinutes), TimeSpan.FromMinutes(60));
        //        engine.Convert(orig, newFile2, options);
        //    }
        //    files.Add(newFile1.Filename);
        //    files.Add(newFile2.Filename);
        //    return files;
        //}
    }
}
