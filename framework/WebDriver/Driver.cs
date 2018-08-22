using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Framework.PageObjects;
using Framework.WaitHelpers;
using Newtonsoft.Json;
using NLog;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Interfaces;
using OpenQA.Selenium.Appium.MultiTouch;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Safari;

namespace Framework.WebDriver
{

    public partial class Driver : IParent
    {
        private readonly IWebDriver _webDriver;

        public IWebDriver WrappedDriver => _webDriver;

        protected static Logger Log = LogManager.GetCurrentClassLogger();

        //private Dictionary<string,string>  DriverInfo = new Dictionary<string, string>();

        public bool IsBrowserstackSession = false;

        public Driver WebDriver => this;

        /// <summary>
        /// Get/Set Screen orientation if rotatable
        /// Returns Portrait if non-rotatable
        /// </summary>
        //public ScreenOrientation ScreenOrientation
        //{
        //    get => _webDriver is IRotatable ? IosDriver.Orientation : ScreenOrientation.Portrait;
        //    set
        //    {
        //        if (_webDriver is IRotatable)
        //            IosDriver.Orientation = value;
        //    }
        //}

        public string WindowHandle
        {
            get => CurrentWindowHandle;
            set => throw new NotImplementedException();
        }

        public string SessionId => ((RemoteWebDriver)_webDriver).SessionId.ToString();

        public ISearchContext SearchContext => _webDriver;


        public void HideKeyboard()
        {
            Log.Info("Hiding keyboard");
            switch (_webDriver)
            {
                case AppiumDriver w:
                    try
                    {
                        w.HideKeyboard();
                        Log.Info("Hiding keyboard was ok");
                    }
                    catch (Exception e)
                    {
                        Log.Info("Hiding keyboard failed");
                        Log.Info(e);
                    }

                    break;

                default:
                    Log.Info($"Hiding keyboard not supported for driver '{_webDriver}'");
                    break;
            }
        }

        public AppiumDriver AppiumDriver => (AppiumDriver) _webDriver;

        // nickName, windowHandle
        public Dictionary<string,string> windowHandles = new Dictionary<string, string>();

        public bool IsMobile()
        {
            if (_webDriver.GetType() == typeof(AppiumDriver))
                return true;
            return false;
        }

        private string SwitchToNewWindow(Action action, string nickname)
        {
            Log.Info("Switching To New Window after action");
            var currentHandles = WindowHandles;
            Log.Info($"Current Handle count = {currentHandles.Count}, performing action");
            action();
            Wait.UntilOrThrow(() => WindowHandles.Count > currentHandles.Count, message:"Wait for new window to appear");

            var newWinCount = WindowHandles.Count;
            Log.Info($"New window count = {newWinCount}");

            var newHandle = WindowHandles.Except(currentHandles).Single();
            Log.Info($"Got new window handle : {newHandle}");

            windowHandles[newHandle] = nickname;
            SwitchToWindow(newHandle);
            Log.Info($"Current Handle count after new ones created = {WindowHandles.Count}");
            return newHandle;
        }

        public T SwitchToNewWindow<T>(Action action, string nickname) where T : PageObject,new()
        {
            var newWindow = SwitchToNewWindow(action, nickname);
            var po = PageObjectFactory.Init<T>(WebDriver);
            po.WindowHandle = newWindow;
            return po;
        }

        internal void SwitchToWindow(string windowHandle)
        {
           var success = windowHandles.ContainsKey(windowHandle);

            if (!success)
            {
                throw new Exception($"Window handle {windowHandle} not found in driver's internal list of window handles");
            }

            var nickName = windowHandles[windowHandle];

            // Context switching looks ok for all devices/browsers

            Log.Info($"Got window Nickname  = {nickName} ,  Handle = {windowHandle}");
            SwitchTo().Window(windowHandle);
            Log.Info("Switch completed");

            Log.Info("Check the new window is currentHandle");
            Wait.UntilOrThrow(() => CurrentWindowHandle == windowHandle, 20, message: "Window not the current Window");
            try
            {
                Log.Info($"Current Browser Title is now : {WebDriver.Title}");
                Log.Info($"Current Browser URL is now :   {WebDriver.Url}");
             
            }
            catch (Exception)
            {
                Log.Warn("Failed to get Browser Title...");
            }



            if (IsMobile())
            {
                // Check if we have focus in the current document
                const string focus = "return document.hasFocus();";
                var focus1 = ExecuteJavascript<bool>(focus);
                var focusInTabbed = false;

                // One window means we do no need any TAB switching
                if (focus1 || WindowHandles.Count == 1)
                {
                    Log.Info($"Focus = {focus1}");
                    Log.Info($"Wincount = { WindowHandles.Count}");
                    focusInTabbed = true;
                }
                else
                {
                    // get device width
                    var deviceWidth = (double)WebDriver.AppiumDriver.Manage().Window.Size.Width;
                    var tabs = (double)WindowHandles.Count;
                    Log.Info($"Device width = {deviceWidth}");
                    Log.Info($"Tab count = {tabs}");

                    // Assume same width is applied to the tabs
                    var tabwidth = (double)(deviceWidth / tabs);
                    Log.Info($"tab width = {tabwidth}");

                    // Try with 80 to reach the TAB row, have not measured....
                    const int yclick = 80;

                    // Walk through the tabs, click and check if the drivers context/window has focus in its doc
                    for (var i = 0; i < tabs; i++)
                    {
                       // Click in the middle of the TAB
                        var xclick = tabwidth * (i + 0.5);
                        Log.Info($"Clicking at {xclick},{yclick}");
                        var ta = new TouchAction(WebDriver.AppiumDriver);
                        ta.Tap(xclick, yclick).Perform();

                        // Check if our document did get focus now
                        focusInTabbed = Wait.Until(() => ExecuteJavascript<bool>(focus), 3, 500);

                        if(focusInTabbed)
                            break;
                    }
                }

                if(!focusInTabbed)
                    throw new Exception("Never succeeded to set the current driver window in focus");
            }
        }

        public enum OS
        {
            Windows,
            OSX,
            Android
        }

        public OS Platform
        {
            get
            {
                var t = Capabilities.Platform.PlatformType;

                if (((string)WebDriver.Capabilities.GetCapability("platformName")).ToLower() == "darwin")
                {
                    return OS.OSX;
                }

                if (t == PlatformType.Android)
                    return OS.Android;

                return OS.Windows;
            }
        }

        public bool IsIe8()
        {
            var browser = Capabilities.GetCapability(@"browserName") as string;
            var version = Capabilities.GetCapability(@"version") as string;
           
            var isIe8 = browser == "internet explorer" && version == "8";
            Log.Debug($"Browser {browser} Version {version}, Is IE8 : {isIe8}");
            return isIe8;

        }

        public bool IsIe9()
        {
            var browser = Capabilities.GetCapability(@"browserName") as string;
            var version = Capabilities.GetCapability(@"version") as string;

            var isIe9 = browser == "internet explorer" && version == "9";
            Log.Debug($"Browser {browser} Version {version}, Is IE9 : {isIe9}");
            return isIe9;

        }

        public bool IsIe10()
        {
            var browser = Capabilities.GetCapability(@"browserName") as string;
            var version = Capabilities.GetCapability(@"version") as string;

            var isIe10 = browser == "internet explorer" && version == "10";
            Log.Debug($"Browser {browser} Version {version}, Is IE10 : {isIe10}");
            return isIe10;

        }

        public bool IsAndroid()
        {
            return Capabilities.GetCapability(@"platformName") as string == "Android";
        }

        public class DriverInfo
        {
            public enum OS
            {
                Unknown,
                Windows,
                OSX,
                iOS,
                Android
            }

            public OS Os;
            public string OsVersion = null;
            public string BrowserName = null;

        }

        public Browser BrowserType
        {
            get
            {
                var browser = Capabilities.GetCapability(@"browserName") as string;

                switch (browser.ToLower())
                {
                    case "internet explorer":
                        return Browser.InternetExplorer;
                    case "firefox":
                        return Browser.Firefox;
                    case "chrome":
                        return Browser.Chrome;
                    case "safari":
                        return Browser.Safari;
                    case "edge" :
                        return Browser.Edge;
                    case "microsoftedge":
                        return Browser.Edge;
                    default:
                        throw new Exception($"Unknown browser type {browser}");
                }
            }
        }

        public void Dispose()
        {
            _webDriver.Dispose();
        }

        // Follow the browser stack wording
        // firefox, chrome, internet explorer, safari, opera, edge, iPad, iPhone, android
        public enum Browser
        {
            Chrome,
            Firefox,
            Edge,
            Opera,
            InternetExplorer,
            IE,
            Safari,
            iPad,
            iPhone,
            Android
        }


        public ICapabilities Capabilities => ((RemoteWebDriver) (_webDriver)).Capabilities;

        public Driver Close()
        {
            var handle = CurrentWindowHandle;
            var mywin = windowHandles.Single(p => p.Key == handle);
            var lastWindow = false;

            Log.Debug($"Closing window {mywin.Key} / {mywin.Value}");

            if (WindowHandles.Count == 1)
            {
                lastWindow = true;
            }

            var beforeCount = _webDriver.WindowHandles.Count;

            if (IsMobile() && !IsAndroid())
            {
                var focusInTabbed = false;
                // get device width
                var deviceWidth = (double) WebDriver.AppiumDriver.Manage().Window.Size.Width;
                var tabs = (double) WindowHandles.Count;
                Log.Info($"Device width = {deviceWidth}");
                Log.Info($"Tab count = {tabs}");
                const string focus = "return document.hasFocus();";

                // Assume same width is applied to the tabs
                var tabwidth = (double) (deviceWidth / tabs);
                Log.Info($"tab width = {tabwidth}");

                // Try with 80 to reach the TAB row, have not measured....
                const int yclick = 80;

                var tabIndex = 0;

                // Walk through the tabs, click and check if the drivers context/window has focus in its doc
                for (var i = 0; i < tabs; i++)
                {
                    tabIndex = i;
                    // Click in the middle of the TAB
                    var xclick = tabwidth * (i + 0.5);
                    Log.Info($"Clicking at {xclick},{yclick}");
                    var ta = new TouchAction(WebDriver.AppiumDriver);
                    ta.Tap(xclick, yclick).Perform();

                    // Check if our document did get focus now
                    focusInTabbed = Wait.Until(() => ExecuteJavascript<bool>(focus), 3, 500);

                    if (focusInTabbed)
                        break;
                }

                if (!focusInTabbed)
                    throw new Exception("Never succeeded to set the current driver window in focus");

                // Click at the left of the TAB
                var xclickKill = tabwidth * (tabIndex) + 20;
                Log.Info($"Clicking at {xclickKill},{yclick}");
                var taKill = new TouchAction(WebDriver.AppiumDriver);
                taKill.Tap(xclickKill, yclick).Perform();

            }
            else
            {
                _webDriver.Close();
            }

            // Remove from dictionary
            windowHandles.Remove(mywin.Key);

            foreach (var key in windowHandles)
            {
                Log.Debug($"Internal handle left = {key.Key}");
            }

            // No more windows to handle, driver is gone
            if (lastWindow) return this;

            Wait.UntilOrThrow(() =>
                {
                    var h = _webDriver.WindowHandles;
                    return !h.Contains(mywin.Key);
                },
                message: "Wait for handle to disappear from WindowHandles");

            //If possible switch to the last remaining window
            if (WindowHandles.Any())
            {
                _webDriver.SwitchTo().Window(WindowHandles.Last());
            }

            return this;
        }

        public void Quit()
        {
            _webDriver.Quit();
            OnQuit(this);
        }

        public IOptions Manage()
        {
            return _webDriver.Manage();
        }

        public INavigation Navigate()
        {
            return _webDriver.Navigate();
        }

        public ITargetLocator SwitchTo()
        {
            return _webDriver.SwitchTo();
        }

        public string Url => _webDriver.Url;
        public Driver(IWebDriver driver)
        {
            _webDriver = driver;
            try
            {
                Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(0);
            }
            catch (Exception e)
            {
                Log.Warn("Failed to set timeout");
                Log.Warn(e);
            }
            
            // Store facts about the current Driver from browserstack
           // this.IsDevice = Capabilities.


        }

        public string Nickname { get; set; }

        public string Title => _webDriver.Title;

        public string PageSource => _webDriver.PageSource;

        public string CurrentWindowHandle => _webDriver.CurrentWindowHandle;

        public ReadOnlyCollection<string> WindowHandles => _webDriver.WindowHandles;

        public Action<Driver> OnQuit { get; set; }

        public Actions Actions => new Actions(_webDriver);

        public TouchAction TouchAction => new TouchAction((IPerformsTouchActions)_webDriver);


        public T ExecuteJavascript<T>(string script, params object[] args)
        {
            return (T) ExecuteScript(script, args);
        }

        public void ExecuteJavascript(string script, params object[] args)
        {
            ExecuteScript(script, args);
        }

        private object ExecuteScript(string script, params object[] args)
        {
            Log.Info($"Executing script '{script}' with args {string.Join(",",args)}");
            return ((IJavaScriptExecutor) _webDriver).ExecuteScript(script, args);
        }

        private object ExecuteAsyncScript(string script, params object[] args)
        {
            return ((IJavaScriptExecutor) _webDriver).ExecuteAsyncScript(script, args);
        }

        private List<string> ScreenshotNames = new List<string>();

        public void TakeScreenshot(string path)
        {
            string fullPath = null;

            try
            {
                var time = DateTime.UtcNow.ToString("-yyyyMMdd_HHmmss") + "_UTC.png";
                var shot = ((ITakesScreenshot) _webDriver).GetScreenshot();
                fullPath = path + time;
                shot.SaveAsFile(fullPath, ScreenshotImageFormat.Png);
                ScreenshotNames.Add(path);
            }
            catch (Exception e)
            {
                Log.Warn($"Failed to take screenshot {e.Message}");
                Log.Warn($"Tried path {fullPath}");
            }
        }
        

        //Local run on regular browser
        public static Driver CreateDriver(Browser browser, string downloadDirectory)
        {

            IWebDriver webdriver;

            switch (browser)
            {
                case Browser.Chrome:
                    var options = new ChromeOptions();
                    
                    options.AddArguments("start-maximized");
                    options.AddArguments("disable-infobars");
                    //var downloadDirectory = "C:\\Temp";
                    options.AddUserProfilePreference("download.default_directory", downloadDirectory);
                    options.AddUserProfilePreference("download.prompt_for_download", false);
                    options.AddUserProfilePreference("download.directory_upgrade", true);

                    //options.AddExcludedArgument("enable-automation");
                    options.AddUserProfilePreference("credentials_enable_service", false);
                    webdriver = new ChromeDriver(options);
                    webdriver.Manage().Timeouts().PageLoad = new TimeSpan(0,2,0);
                    break;
                case Browser.Firefox:
                    var ffopts = new FirefoxOptions {UseLegacyImplementation = false};

                    var firefoxProfile = new FirefoxProfile();
                   
                    firefoxProfile.SetPreference("browser.download.folderList", 2);
                    firefoxProfile.SetPreference("browser.download.manager.focusWhenStarting", false);
                    firefoxProfile.SetPreference("browser.download.dir",@"c:\tools");
                    firefoxProfile.SetPreference("browser.download.useDownloadDir", true);
                    firefoxProfile.SetPreference("browser.helperApps.alwaysAsk.force", false);
                    firefoxProfile.SetPreference("browser.download.manager.alertOnEXEOpen", false);
                    firefoxProfile.SetPreference("browser.download.manager.closeWhenDone", true);
                    firefoxProfile.SetPreference("browser.download.manager.showAlertOnComplete", false);
                    firefoxProfile.SetPreference("browser.download.manager.useWindow", false);
                    firefoxProfile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/force-download, text/xml");
                   
                    ffopts.Profile = firefoxProfile;

                    webdriver = new FirefoxDriver(ffopts);

                    break;
                case Browser.Edge:
                    var edgeopts = new EdgeOptions {PageLoadStrategy = PageLoadStrategy.Eager};
                    webdriver = new EdgeDriver(edgeopts);
                    break;
                case Browser.InternetExplorer:
                    // IE
                    var opts = new InternetExplorerOptions
                    {
                        EnableNativeEvents = true,
                        EnsureCleanSession = true,
                        EnablePersistentHover = false,
                        RequireWindowFocus = false,
                    };
                    webdriver = new InternetExplorerDriver(opts);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(browser), browser, null);
            }

            var caps = ((RemoteWebDriver) webdriver).Capabilities;

            Log.Info("Capabilities looks like this");
            var json = JsonConvert.SerializeObject(caps);

            Log.Info(json);

       
            var driver = new Driver(webdriver);
            return driver;
        }

        public static Driver CreateDriver(Dictionary<string,string> browserStackCon, Dictionary<string, string> capabilities, string project, string testname)
        {
            var browser = capabilities["browser"];
            var hub = browserStackCon["browserstack.hub"];
            IWebDriver webdriver;

            if (hub == null)
                throw new Exception("No hub defined for BrowserStack");

            browserStackCon.Remove("browserstack.hub");

            DesiredCapabilities desiredCapabilities;
            Log.Info($"Creating driver for '{browser}'");

            switch (browser)
            {
                case "IE":
                    var opts = new InternetExplorerOptions
                    {
                        EnsureCleanSession = true,
                        EnablePersistentHover = false,
                        RequireWindowFocus = false,
                        PageLoadStrategy = PageLoadStrategy.Normal,
                        EnableNativeEvents = true
                    };
                    desiredCapabilities = (DesiredCapabilities)opts.ToCapabilities();
                    //desiredCapabilities.SetCapability("browserstack.ie.enablePopups", true);

                    desiredCapabilities = new DesiredCapabilities();
                    //desiredCapabilities.SetCapability("ie.ensureCleanSession", true);
                    //desiredCapabilities.SetCapability("enablePersistentHover", false);
                    //desiredCapabilities.SetCapability("requireWindowFocus", true);
                    //desiredCapabilities.SetCapability("nativeEvents", true);
                    //desiredCapabilities.SetCapability("ie.pageLoadStrategy", PageLoadStrategy.Normal);
                    desiredCapabilities.SetCapability("browserstack.ie.enablePopups", true);
                    //desiredCapabilities.SetCapability("browserstack.ie.driver", "3.9.0");
                    
                    break;
                case "Chrome":
                    var options = new ChromeOptions
                    {
                        PageLoadStrategy = PageLoadStrategy.Normal
                    };

                    options.AddArguments("start-maximized");
                    options.AddArguments("disable-infobars");
                    options.AddUserProfilePreference("credentials_enable_service", false);
                    var downloadDirectory = "C:\\Temp";
                    options.AddUserProfilePreference("download.default_directory", downloadDirectory);
                    options.AddUserProfilePreference("download.prompt_for_download", false);
                    options.AddUserProfilePreference("download.directory_upgrade", true);
                    options.AddUserProfilePreference("disable-popup-blocking", "true");
                    desiredCapabilities = (DesiredCapabilities)options.ToCapabilities();
                    break;

                case "Firefox":
                    var ffopts = new FirefoxOptions
                    {
                        PageLoadStrategy = PageLoadStrategy.Normal,
                        LogLevel = FirefoxDriverLogLevel.Warn
                    };

                    var firefoxProfile = new FirefoxProfile();
                    firefoxProfile.SetPreference("browser.download.folderList", 0);
                    firefoxProfile.SetPreference("browser.download.manager.focusWhenStarting", false);
                    firefoxProfile.SetPreference("browser.download.useDownloadDir", true);
                    firefoxProfile.SetPreference("browser.helperApps.alwaysAsk.force", false);
                    firefoxProfile.SetPreference("browser.download.manager.alertOnEXEOpen", false);
                    firefoxProfile.SetPreference("browser.download.manager.closeWhenDone", true);
                    firefoxProfile.SetPreference("browser.download.manager.showAlertOnComplete", false);
                    firefoxProfile.SetPreference("browser.download.manager.useWindow", false);
                    firefoxProfile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/pdf");
                    ffopts.Profile = firefoxProfile;

                    desiredCapabilities = (DesiredCapabilities)ffopts.ToCapabilities();
                    desiredCapabilities.SetCapability("browserstack.video", true);
                    desiredCapabilities.SetCapability(FirefoxDriver.ProfileCapabilityName, firefoxProfile.ToBase64String());
                    break;

                case "Safari":
                    var safariOptions = new SafariOptions
                    {
                        PageLoadStrategy = PageLoadStrategy.Normal
                    };

                    desiredCapabilities = (DesiredCapabilities)safariOptions.ToCapabilities();
                    desiredCapabilities.SetCapability("browserstack.safari.allowAllCookies", true);
                    desiredCapabilities.SetCapability("browserstack.safari.enablePopups", true);
                    desiredCapabilities.SetCapability("browserstack.safari.driver", "2.48");

                    break;
                case "Edge":
                    var edgeopts = new EdgeOptions { PageLoadStrategy = PageLoadStrategy.Eager };
                    desiredCapabilities = (DesiredCapabilities)edgeopts.ToCapabilities();
                    desiredCapabilities.SetCapability("browserstack.edge.enablePopups", true);
                    
                    break;

                case "iPad":
                case "iPhone":
                case "WebdriveriPad":

                    var safariOptionsiPad = new SafariOptions();
                    desiredCapabilities = (DesiredCapabilities)safariOptionsiPad.ToCapabilities();
                    desiredCapabilities.SetCapability("browserstack.safari.allowAllCookies", true);
                    desiredCapabilities.SetCapability("browserstack.safari.enablePopups", true);
                    
                  
                    break;


                case "Android":
                case "android":
                    
                    desiredCapabilities = new DesiredCapabilities();
                    desiredCapabilities.SetCapability("unicodeKeyboard", true);
                    desiredCapabilities.SetCapability("resetKeyboard", true);  

                    break;
                default:
                    throw new Exception($"Unknown browser sent to BrowserStack {browser}");
            }

            desiredCapabilities.SetCapability("project", project);
            desiredCapabilities.SetCapability("name", testname);

            // Set browserstack params
            foreach (var key in browserStackCon.Keys)
            {
                var value = browserStackCon[key];
                Log.Debug($"Setting browserstack capability {key}={value}");
                desiredCapabilities.SetCapability(key, value);
            }

            // Set desired capabilities
            foreach (var key in capabilities.Keys)
            {
                var value = capabilities[key];
                Log.Info($"Setting desired capability {key}={value}");
                desiredCapabilities.SetCapability(key, value);
            }

            // Fetch driver
            Log.Info("Creating driver...");


            if (browser == "iPad" || browser == "iPhone")
            {
                Log.Info("Creating IOS driver...(Appium)");
                webdriver = new AppiumDriver(new Uri(hub), desiredCapabilities,new TimeSpan(0, 2, 0));
            }
            else if (browser == "WebdriveriPad")
            {

                webdriver = new AppiumDriver(new Uri(hub), desiredCapabilities, new TimeSpan(0, 2, 0));
            }
            else if (browser.ToLower() == "android")
            {
                Log.Info("Creating Android driver...(Appium)");

                webdriver = new AppiumDriver(new Uri(hub), desiredCapabilities, new TimeSpan(0, 2, 0));
            }
            else
            {
                Log.Info("Creating Desktop driver...");
                webdriver = new RemoteWebDriver(new Uri(hub), desiredCapabilities,new TimeSpan(0,2,0));
            }
            // var 
            
         
            Log.Info("Driver created...");

            var caps = ((RemoteWebDriver)webdriver).Capabilities;

            try
            {
                switch (browser)
                {
                    case "IE":
                        
                        Log.Info("Adding PageLoad Timeout");
                        webdriver.Manage().Timeouts().PageLoad = TimeSpan.FromMinutes(2);
                        Log.Info($"Got capability nativeEvents    = {caps.GetCapability("nativeEvents")}");
                        Log.Info(
                            $"Got capability ie.ensureCleanSession    = {caps.GetCapability("ie.ensureCleanSession")}");
                        Log.Info(
                            $"Got capability enablePersistentHover = {caps.GetCapability("enablePersistentHover")}");
                        Log.Info(
                            $"Got capability RequireWindowFocus    = {caps.GetCapability("RequireWindowFocus")}");
                        Log.Info(
                            $"Got capability browserstack.ie.driver    = {caps.GetCapability("browserstack.ie.driver")}");
                        Log.Info(
                            $"Got capability browserstack.ie.enablePopups    = {caps.GetCapability("browserstack.ie.enablePopups")}");

                        break;
                }
            }
            catch (Exception)
            {
                // ignored
            }

            Log.Info("Capabilities...");
            var settings = new JsonSerializerSettings() { ContractResolver = new MyContractResolver() };
            var json = JsonConvert.SerializeObject(caps, Formatting.Indented,settings);
            Log.Info(json);

            var driver = new Driver(webdriver) {IsBrowserstackSession = true};
            return driver;
        }
    }
}
