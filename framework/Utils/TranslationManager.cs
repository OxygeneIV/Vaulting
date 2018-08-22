using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using Framework.Extensions;
using Framework.PageObjects;
using NLog;

namespace Framework.Utils
{
    public class TranslationManager
    {
        protected static Logger Log = LogManager.GetCurrentClassLogger();
        private string _version;
        private bool _active;
        private string _language;

        private string _testCaseFolder;
        private string _zippedFolder;
        private int _packageContainer;
        private LanguageIdentifier _languageId;

        public enum LanguageIdentifier
        {
            English,
            German,
            Spanish,
            French,
            Japanese,
            Malayalam,
            Polish,
            Swedish,
            ChineseSimplified,
            SriLankian,
            Skip
        }

        public static  LanguageIdentifier LanguageId
        {
            get { return Instance._languageId; }
            set { Instance._languageId = value; }
        }

        public static string TestCaseFolder
        {
            get { return Instance._testCaseFolder; }
            set { Instance._testCaseFolder = value; }
        }

        public static int PackageContainer
        {
            get { return Instance._packageContainer; }
            set { Instance._packageContainer = value; }
        }

        public static string ZippedFolder
        {
            get { return Instance._zippedFolder; }
            set { Instance._zippedFolder = value; }
        }

        public static string Version
        {
            get { return Instance._version; }
            set { Instance._version = value; }
        }

        public static string Language
        {
            get { return Instance._language; }
            set { Instance._language = value; }
        }

        public static TranslationManager Instance => _instance ?? (_instance = new TranslationManager());

        public static bool Active
        {
            get { return Instance._active; }
            set { Instance._active = value; }
        }

        private static TranslationManager _instance;

        

        public static void Translate(KeyValuePair<Enum, Func<PageObject>> kvp)
        {
            if (!Active) return;
            var po = kvp.Value();
            var t = kvp.Key.GetType().Name;
            Screenshot(po, t + "-" + kvp.Key,kvp.Key.ToString().EndsWith("_Tooltip"));
        }

        private int _translationImage;


        private static void Screenshot(PageObject pageObject, string name, bool trueSnapshot = false)
        {
            string oldBorder = null;

            try
            {
                oldBorder = pageObject.HighlighElement();              
                var preliminaryScreenshotNumber = Instance._translationImage + 1;
                var fullname = $"Translation_{Version}_{Language}_{preliminaryScreenshotNumber:D3}_" + name;
                var path = Path.Combine(TestCaseFolder, fullname);
                if (trueSnapshot)
                {
                    // We need to wait for the page to be updated
                    System.Threading.Thread.Sleep(2000);
                    ScreenCapture.CaptureActiveWindowToFile(path + ".png", ImageFormat.Png);
                }
                else
                {
                    pageObject.WebDriver.TakeScreenshot(path);
                }
                Instance._translationImage++;

            }
            catch(Exception e)
            {
                var mess =
                    $"Failed to make Translation Screenshot of '{name}' using driver {pageObject.WebDriver.Nickname}";
                Log.Warn(mess);
                Log.Warn(e.Message);
                throw;
            }
            finally
            {
                pageObject.HighlighElement(oldBorder);
            }
        }
    }
}
