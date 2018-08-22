using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

using Castle.Core.Internal;
using Framework.Utils;
using Framework.WaitHelpers;
using Framework.WebDriver;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using NLog;
using NLog.Targets;
using OpenQA.Selenium.Remote;
//using Viedoc.viedoc.me;
//using Viedoc.viedoc.pages.login;

using Parallel = System.Threading.Tasks.Parallel;

namespace Tests.Base
{
    [TestClass]
    public abstract class TestBase
    {
        public TValue GetMethodAttributeValue<TAttribute, TValue>(Func<TAttribute, TValue> valueSelector) where TAttribute : Attribute, new()
        {
            var attr = GetMethodAttribute<TAttribute>();
            return attr != null ? valueSelector(attr) : default(TValue);
        }

        private TAttribute GetMethodAttribute<TAttribute>() where TAttribute : Attribute,new()
        {
            var methodInfo = GetType().GetMethod(TestContext.TestName);
            var attr = methodInfo.GetCustomAttributes(typeof(TAttribute), true);
            if (attr.IsNullOrEmpty())
            {
                return null;
            }

            return attr.FirstOrDefault() as TAttribute;
        }

        private TAttribute GetClassAttribute<TAttribute>() where TAttribute : Attribute, new()
        {
            var methodInfo = GetType();
            var attr = methodInfo.GetCustomAttributes(typeof(TAttribute), true);
            if (attr.IsNullOrEmpty())
            {
                return null;
            }
            return attr.FirstOrDefault() as TAttribute;
        }


        public bool PreStartBrowser
        {
            get
            {
                var method = GetMethodAttribute<BrowserAttribute>();
                if (method != null)
                    return method.Initiate;

                var cl = GetClassAttribute<BrowserAttribute>();
                if (cl != null)
                    return cl.Initiate;

                return false;

            }
        }

        private static readonly object Lock = new object();
        internal ConcurrentBag<Driver> DriversInTest = new ConcurrentBag<Driver>();

        private bool TestIsAlreadyPassed = true;

        // <sessionId,Nickname>
        private readonly Dictionary<string,string> _driverNicknameDictionary = new Dictionary<string, string>();
        private readonly Dictionary<string, bool> _driverBrowserstackDictionary = new Dictionary<string, bool>();

        protected static readonly Logger Log = LogManager.GetCurrentClassLogger();
        protected static readonly Logger NotesLog = LogManager.GetLogger("Notes");

        //internal static LoginPage ViedocLoginPage;
        //internal ViedocMeLoginPage ViedocMeLoginPage;

        // Save a timestamp for this session
        private static string UtcNow => DateTime.UtcNow.ToString("yyyyMMdd_HHmmss") + "_UTC";

        private static readonly string TestRunUtcStart = TestRunUtcStart ?? UtcNow; 

        private static readonly string ResultFolder =ConfigurationManager.AppSettings["TestResults"];

        internal static readonly bool DoTakeScreenshot = Convert.ToBoolean(ConfigurationManager.AppSettings["TakeScreenshot"]);

        public string TestClass => TestContext.FullyQualifiedTestClassName;
        public string TestName => TestContext.TestName;

        public string TestDir => Directory.Exists(ResultFolder) ? 
            Path.Combine(ResultFolder, TestRunUtcStart) : Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "TestResults", TestRunUtcStart);

        public string TestClassFolder => Path.Combine(TestDir, TestClass);

        public string TestTranslationFolder => Path.Combine(TestDir, "Translation");

        public string TestCaseFolder => Path.Combine(TestClassFolder,TestName);
        public TestContext TestContext { get; set; }
        protected System.Diagnostics.Stopwatch StopWatch;
        private string _currentLogFile;
        private string _currentNotesFile;

        protected bool UsingBrowserStack;

        private readonly string _utcStarttime = UtcNow;

        public TimeSpan Duration;

        public  string TestProperty(string key)
        {
            return TestProperty<string>(key);
        }

        public bool TestPropertyExists(string key)
        {
            return TestContext.Properties.Contains(key);
        }

        public void LogTestContext()
        {
            //Log.Info("TestContext...");
            //var settings = new JsonSerializerSettings() { ContractResolver = new Driver.MyContractResolver() };
            //string json = JsonConvert.SerializeObject(TestContext.Properties, Formatting.Indented, settings);
            //Log.Info(json);
        }

        public T TestProperty<T>(string key)
        {
            var value = Convert.ToString(TestContext.Properties[key]);
            var converter = TypeDescriptor.GetConverter(typeof(T));
            var f = (T)converter.ConvertFromInvariantString(value);
            return f;
        }

        public static Dictionary<string, string> LoadConfig(string settingfile)
        {
            var dic = new Dictionary<string, string>();
            var fullpath = $"Capabilities/{settingfile}";
            if (!File.Exists(fullpath)) return dic;

            var settingdata = File.ReadAllLines(fullpath);
            foreach (var setting in settingdata)
            {
                var sidx = setting.IndexOf("=", StringComparison.Ordinal);
                if (sidx < 0) continue;
                var skey = setting.Substring(0, sidx);
                var svalue = setting.Substring(sidx + 1);
                if (!dic.ContainsKey(skey))
                {
                    dic.Add(skey, svalue);
                }
            }
            return dic;
        }


        /// <summary>
        /// Currently we only start a chrome browser
        /// </summary>
        public Driver CreateBrowserInstance(Driver.Browser browser, string nickName = null, bool usingBrowserStack = false, string capabilities = null)
        {
            Driver driverToReturn;
           
            if (usingBrowserStack)
            {                
                var browserStackSettings =
                    new Dictionary<string, string>
                    {
                        ["browserstack.user"] = TestProperty("browserstack.user"),
                        ["browserstack.key"]  = TestProperty("browserstack.key"),
                        ["browserstack.hub"]  = TestProperty("browserstack.hub"),
                        ["browserstack.selenium_version"] = TestProperty("browserstack.selenium_version"),
                        ["browserstack.geckodriver"] = TestProperty("browserstack.geckodriver"),
                        ["browserstack.local"] = TestProperty("browserstack.local"),
                        ["browserstack.use_w3c"] = TestProperty("browserstack.use_w3c"),
                        ["browserstack.idleTimeout"] = "300",
                        ["browserstack.ie.arch"] = TestProperty("browserstack.ie.arch"),
                    };

                if (TestName.ToLower().Contains("addremovephoneproperties"))
                {
                    browserStackSettings["browserstack.local"] = "true";
                }

                if (TestName.ToLower().Contains("ActivateDeactivateSMSTwoFactorAuthForUser".ToLower()))
                {
                    browserStackSettings["browserstack.local"] = "true";
                }
                

                var capabilitiesFile = capabilities ?? TestProperty("capabilities");
                Log.Info($"Loading capabilities file {capabilitiesFile}");
                var browserStackCapabilities = LoadConfig(capabilitiesFile);

                LocalFileDetector detector = new LocalFileDetector();
                driverToReturn = Driver.CreateDriver(browserStackSettings, browserStackCapabilities, TestClass, TestName);
                ((RemoteWebDriver)(driverToReturn.WrappedDriver)).FileDetector = detector;
            }
            else
            {
                driverToReturn = Driver.CreateDriver(browser, TestCaseFolder);
            }

            // Use unique name unless named
            if (nickName == null)
            {
                var sid = driverToReturn.SessionId;
                if (sid != null)
                {
                    nickName = sid;
                }
                else
                {
                    throw new Exception("SessionId from WebDriver is null");
                }
            }

            driverToReturn.Nickname = nickName;
            driverToReturn.OnQuit = OnQuit;
            Log.Info($"Nickname of new session {nickName}");
            Log.Info($"Window handle of new session {driverToReturn.CurrentWindowHandle}");

            driverToReturn.windowHandles[driverToReturn.CurrentWindowHandle] = "Init Window";
            _driverNicknameDictionary[driverToReturn.SessionId] = nickName;
            _driverBrowserstackDictionary[driverToReturn.SessionId] = usingBrowserStack;
            Log.Info("Adding driver to DriverInTest-bag...");
            DriversInTest.Add(driverToReturn);

            if (!driverToReturn.IsMobile())
            {
                try
                {
                    Log.Info("Try maximize");
                    if (driverToReturn.BrowserType == Driver.Browser.Chrome && driverToReturn.Platform == Driver.OS.OSX)
                        driverToReturn.Manage().Window.Size = new Size(1910, 1070);
                    else
                        driverToReturn.Manage().Window.Maximize();
                }
                catch (Exception e)
                {
                    Log.Warn(e, "Maximize failed, trying to set size");
                    driverToReturn.Manage().Window.Size = new Size(1910, 1070);
                }

                try
                {
                    Log.Info(
                        $"Window size = {driverToReturn.Manage().Window.Size.Width} , {driverToReturn.Manage().Window.Size.Height}");
                }
                catch (Exception e)
                {
                    Log.Warn(e, "Trying to get size failed");
                }
            }
            Log.Info("Returning driver");
            return driverToReturn;
        }

 

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            //if (!TranslationManager.Active) return;

            //var wildname = "trans_" + TranslationManager.Language + "*.zip";
            //string[] files = Directory.GetFiles(TranslationManager.ZippedFolder,wildname);

            //foreach (var file in files)
            //{
            //    try
            //    {
            //        TFSutil.TFS.AddWorkitemAttachment(TranslationManager.PackageContainer, file);
            //    }
            //    catch (Exception e)
            //    {
            //        Log.Warn(e, "Translation AddWorkitemAttachment failed");
            //    }
            //}
        }

        /// <summary>
        /// Instantiate the first driver and ViedocLoginPage for the test
        /// </summary>
        [TestInitialize]
        public void BaseTestInit()
        {
            try
            {

                // Create the test result area
                Directory.CreateDirectory(TestCaseFolder);
                var target = (FileTarget) LogManager.Configuration.FindTargetByName("UITestLogFile");

                // Reconfigure logger
                _currentLogFile = Path.Combine(TestCaseFolder, TestName + "-" + _utcStarttime + ".log");
                target.FileName = _currentLogFile;


                var target2 = (FileTarget)LogManager.Configuration.FindTargetByName("UITestNotesLogFile");

                // Reconfigure logger
                _currentNotesFile = Path.Combine(TestCaseFolder, TestName + "-" + _utcStarttime + "-Notes" + ".log");
                target2.FileName = _currentNotesFile;

                LogManager.ReconfigExistingLoggers();
                
                Log.Info($"Setting up testcase {TestContext.FullyQualifiedTestClassName}");

                //var testrunId = TestProperty<int>("testrun");
                StopWatch = System.Diagnostics.Stopwatch.StartNew();
                StopWatch.Restart();

                // Check if we want to rerun failed tests in a testrun
                //if (testrunId > 0)
                //{
                //    TestIsAlreadyPassed =  TFSutil.TFS.IsTestPassed(testrunId,TestContext.TestName);

                //    if (TestIsAlreadyPassed)
                //       Assert.Inconclusive("Test already passed, skipping");

                //    // or else we have a failed test (or an exception wich is OK)
                //}

//                LogTestContext();

                //UsingBrowserStack = TestProperty<bool>("BrowserStack");
                //Log.Info($"Using browserstack : {UsingBrowserStack}");

                //var useLocalTesting = TestProperty<bool>("browserstack.local");
                //Log.Info($"Browserstack Local testing : {useLocalTesting}");

                // Translation
                //if (TestProperty<bool>("translation"))
                //{
                //    this.InitTranslation();
                //    TranslationManager.TestCaseFolder = Path.Combine(TestCaseFolder, TranslationManager.Language);
                //    TranslationManager.ZippedFolder = TestTranslationFolder;

                //    if (!Directory.Exists(TranslationManager.ZippedFolder))
                //        Directory.CreateDirectory(TranslationManager.ZippedFolder);

                //    if (!Directory.Exists(TranslationManager.TestCaseFolder))
                //        Directory.CreateDirectory(TranslationManager.TestCaseFolder);
                //}
            }
            catch (AssertInconclusiveException)
            {
                Log.Info("Setting inconclusive error, test already run");
                throw;
            }
            catch(Exception e)
            {
                Log.Fatal(e,$"TestInit failed {e.Message}");
                throw;
            }
        }


        private void TeardownSnapshotsWhenFailure()
        {
            //Log.Info("Taking snapshot from all drivers...");
            //var driverCount = DriversInTest.Count;
            //Log.Info($"Number of active drivers  = {driverCount}");

            //if (driverCount <= 0) return;

            //var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = driverCount };

            //var loopResults = Parallel.ForEach(
            //    DriversInTest,
            //    parallelOptions,
            //    driver =>
            //    {
            //        this.TakeScreenshot(driver, $"TeardownScreenshot_{driver.Nickname}");
            //    });

            //Log.Info($@"Snapshots completed = {loopResults.IsCompleted}");
        }


        [TestCleanup]
        public void BaseTestCleanup()
        {
            StopWatch.Stop();

            var passed = TestContext.CurrentTestOutcome == UnitTestOutcome.Passed;

            Duration = StopWatch.Elapsed;
            //var testrunId = TestProperty<int>("testrun");

            //if (testrunId > 0 && TestIsAlreadyPassed)
            //{
            //    Log.Info($"Test {TestContext.FullyQualifiedTestClassName} was already passed for testrun {testrunId}...");
            //    // Test was never run
            //    return;
            //}

            Log.Info($"Final status in test  = {TestContext.CurrentTestOutcome}");

            if (TestContext.CurrentTestOutcome != UnitTestOutcome.Passed)
            {
                TeardownSnapshotsWhenFailure();
            }

            // Quit all drivers gracefully
            QuitDriversInTest();
            
            //var browserstackData             = Path.Combine(TestCaseFolder, "Browserstack.txt");
            //var browserstackVideoUrls        = Path.Combine(TestCaseFolder, "BrowserstackVideos.txt");

            //Browserstack - operations, wrap in try/catch, should not stop progress of teardown
            try
            {
                //if (UsingBrowserStack)
                //{
                //    string user = TestProperty("browserstack.user");
                //    string key = TestProperty("browserstack.key");

                //    Browserstack stack = new Browserstack(user,key);


                //    var recordOnPassed = TestProperty<bool>("pcg.recordingOnSuccess");
                //    var recordOnError = TestProperty<bool>("pcg.recordingOnError");

                //    var allBrowserStackSessions = _driverBrowserstackDictionary.Where(kvp => kvp.Value)
                //        .Select(kvp => kvp.Key).ToList();

                //    Log.Info($"Browserstack session count : {allBrowserStackSessions.Count}");

                //    foreach (var browserStackSession in allBrowserStackSessions)
                //    {
                //        try
                //        {
                //            Log.Info($"Browserstack session : {browserStackSession}");

                //            // Wait for status to Completed
                //            var completed = stack.WaitForBrowserstackSessionCompleted(browserStackSession);

                //            if (!completed)
                //                continue;

                //            // Set the test status
                //            var browserstackSetStatus = stack.SetSessionStatus(browserStackSession,
                //                TestContext.CurrentTestOutcome.ToString(), "Because we can!");

                //            if (!browserstackSetStatus)
                //                continue;

                //            var downloadRecording = recordOnPassed && passed || recordOnError && !passed;
                //            var nickName = _driverNicknameDictionary[browserStackSession];
                //            var recording = Path.Combine(TestCaseFolder, $"{nickName}_{browserStackSession}.mp4");
                //            stack.HandleSessionData(browserStackSession, browserstackData, browserstackVideoUrls, downloadRecording, recording);
                //        }
                //        catch (Exception e)
                //        {
                //            Log.Warn($"Trouble gett session info from Browserstack : {e.Message}");
                //        }
                //    }
                //}
            }
            catch(Exception e)
            {
                Log.Warn($"Error retrieving data from Browserstack {e.Message}");
            }

            // files to be added to test result
            var files2Add = new List<string>
            { _currentLogFile,
              _currentNotesFile,
              //browserstackData,
              //browserstackVideoUrls
            };

            
            files2Add.AddRange(Directory.EnumerateFiles(TestCaseFolder, "*.png"));
            files2Add.AddRange(Directory.EnumerateFiles(TestCaseFolder, "*.pdf"));
            files2Add.AddRange(Directory.EnumerateFiles(TestCaseFolder, "*.xlsx"));
            files2Add.AddRange(Directory.EnumerateFiles(TestCaseFolder, "*.mp4"));

            // Attach results to current test
            try
            {
                foreach (var file in files2Add)
                {
                    if (!File.Exists(file)) continue;
                    Log.Info($"Adding file '{Path.GetFileName(file)}' to test context");
                    TestContext.AddResultFile(file);
                }
            }
            catch (Exception e)
            {
                Log.Warn($"Failed to attach file to test result {e.Message}");
            }

            // Update old test case if rerun
            //try
            //{
            //    if (testrunId > 0 && !TestIsAlreadyPassed)
            //    {
            //        TFSutil.TFS.UpdateTestResult(testrunId,Duration,TestContext,files2Add);          
            //    }
            //}
            //catch (Exception e)
            //{
            //    Log.Warn($"Failed to update test result {e.Message}");
            //}
 

            // Translation
          
            //if (TranslationManager.Active)
            //{
            //    Log.Info("Adding translation...");
            //    foreach (var image in Directory.EnumerateFiles(TranslationManager.TestCaseFolder, "*.png"))
            //    {
            //        TestContext.AddResultFile(image);
            //    }

            //    foreach (var image in Directory.EnumerateFiles(TranslationManager.TestCaseFolder, "*.png"))
            //    {
            //        var fil = Path.GetFileName(image);
            //        var imsplit = fil.Split('_');
            //        var name = imsplit[4];

            //        bool CommonOrClinic = name.ToLower().StartsWith("common-") ||
            //                              name.ToLower().StartsWith("clinic-");

            //        var okLang = TranslationManager.Language == "SE" || TranslationManager.Language == "EN" ||
            //                     TranslationManager.Language == "JP" || TranslationManager.Language == "PRC" || TranslationManager.Language == "XX" ||
            //                     TranslationManager.Language == "SR";

            //        if (!CommonOrClinic && !okLang) continue;

            //        var filename = "trans_" + TranslationManager.Language + "_" + _translationSequenceNumber + "_.zip";
            //        var z = Path.Combine(TranslationManager.ZippedFolder, filename);

            //        var shortName = Path.GetFileName(image);
            //        Log.Info($"Adding {shortName} to zip-file");
            //        using (var archive = ZipFile.Open(z, ZipArchiveMode.Update))
            //        {
            //            archive.CreateEntryFromFile(image, shortName, CompressionLevel.Optimal);
            //        }
            //        var length = new FileInfo(z).Length;

            //        if (length > 3670016)
            //        {
            //            _translationSequenceNumber++;
            //        }
            //    }
            //}
        }

        protected List<string> SavedFileList(string searchPattern = null)
        {
            var files = searchPattern != null ? Directory.GetFiles(TestCaseFolder, searchPattern) : Directory.GetFiles(TestCaseFolder);
            return files.ToList();
        }

        protected string WaitForNewFileDownloaded(Action action, string searchPattern = null)
        {
            Log.Info("Wait for new file...");
            var currentList = SavedFileList(searchPattern);
            Log.Info($"currentList count  = {currentList.Count}");
            var newList = new List<string>();

            action();

            Wait.UntilOrThrow(() =>
            {
                newList = SavedFileList(searchPattern);
                Log.Info($"New Filelist count  = {newList.Count}");

                newList = newList.Except(currentList).ToList();

                if (!newList.Any()) return false;

                Log.Info("Wait for new file completed...");
                return true;
            });

            var fileInfo = new FileInfo(newList.First());
            Log.Info($"New filename = {fileInfo.Name}");

            // Wait for stable size
            Wait.UntilOrThrow(() =>
            {
                var length1 = fileInfo.Length;
                System.Threading.Thread.Sleep(3000);
                var length2 = fileInfo.Length;
                return length1 == length2;
            });

            Log.Info($"New filename when stabilzed download = {fileInfo.Name}");
            return fileInfo.Name;
        }



        private void OnQuit(Driver driver)
        {
            // Not a thread safe method to remove items from list, so add a lock
            lock (Lock)
            {
                DriversInTest = new ConcurrentBag<Driver>(DriversInTest.Except(new[] { driver }));
            }
        }

 



        public void QuitDriversInTest()
        {
            var driverCount = DriversInTest.Count;
            Log.Info($"Number of active drivers  = {driverCount}");

            if (driverCount == 0)
                return;

            Log.Info("Quitting all drivers in test");
            var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = driverCount };

            var loopResults = Parallel.ForEach(
                 DriversInTest,
                 parallelOptions,
                 driver =>
                 {
                     driver.Quit();
                 });

           
            Log.Info($@"Driver.Quit status from parallell quit = {loopResults.IsCompleted}");

            driverCount = DriversInTest.Count;

            // We should have an empty list now....
            Log.Info($@"Number of active drivers left = {driverCount}");
        }
    }
}
