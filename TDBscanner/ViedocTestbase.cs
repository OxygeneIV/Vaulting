using System;
using System.Collections.Generic;
using System.IO;
using Framework.PageObjects;
using Framework.WaitHelpers;
using Framework.WebDriver;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace Tests.Base
{
    [Browser(true)]
    public class ViedocTestbase : TestBase
    {
        protected string Password => TestProperty("password");
        protected string Email => TestProperty("user1");
        protected string Fullname => TestProperty("userfullname");
        protected string Fullname2 => TestProperty("userfullname2");

        protected string Email2 => TestProperty("user2");
        protected string Email3 => TestProperty("user3");
       
        protected string Site = "site 01";
        protected string StudyName => TestProperty("studyname");
        protected string Visit1 => "Scheduled 1";
        protected string UnscheduledVisit1 => "Unscheduled 1";
        protected string PublishableStudy => TestProperty("publishable");

        protected string WorkflowStudy => TestProperty("workflowstudy");

        protected string Organization => TestProperty("organization");


        public static Dictionary<string, string> LoadEnvironment(string settingfile)
        {
            var dic = new Dictionary<string, string>();
            //var fullpath = $"Environments/{settingfile}";
            //if (!File.Exists(fullpath))
            //{
            //    throw new Exception("Need environment file (internaltest or stage)");
            //}

            //var settingdata = File.ReadAllLines(fullpath);
            //foreach (var setting in settingdata)
            //{
            //    var sidx = setting.IndexOf("=", StringComparison.Ordinal);
            //    if (sidx < 0) continue;
            //    var skey = setting.Substring(0, sidx);
            //    var svalue = setting.Substring(sidx + 1);
            //    if (!dic.ContainsKey(skey))
            //    {
            //        dic.Add(skey, svalue);
            //    }
            //}
            return dic;
        }


        [TestInitialize]
        public void StartBrowser()
        {
            Log.Info("Viedoc Testbase Testcase Init");

            //Environment.SetEnvironmentVariable("TESTENV", TestProperty("env").Trim());

            //// Viedoc specific          
            //var env = TestProperty("env").Trim() + ".cfg";

            //var envDictionary = LoadEnvironment(env);
            //Log.Info("Appending variables from environment file to TestContext.Properties dictionary");

            //foreach (var kvp in envDictionary)
            //{
            //    if (TestPropertyExists(kvp.Key))
            //    {
            //        // Overwritten by test run parameter
            //        Log.Info(
            //            $"variable {kvp.Key} is set (overridden by build def parameter), using it's value {kvp.Value}");
            //    }
            //    else
            //    {
            //        TestContext.Properties.Add(kvp.Key, kvp.Value);
            //    }
            //}


            if (!PreStartBrowser) return;

            // Create the default browser
            var browser = TestProperty("browser");
            var theBrowser = browser != null
                ? (Driver.Browser)Enum.Parse(typeof(Driver.Browser), browser)
                : Driver.Browser.Chrome;

            //if (ViedocMeAttr.Initiate)
            //{
            //    bool forceLocalChrome = TestProperty("pcg.viedocmecapability").Trim().Length > 0;
            //    useBrowserStack = _usingBrowserStack && forceLocalChrome == false;
            //}

            var driver = CreateBrowserInstance(theBrowser, "Main", UsingBrowserStack);

            // Navigate
            var url = TestProperty("applicationUrl");
            driver.Navigate().GoToUrl(url);
            //if (url.ToLower().Contains("4me"))
            //{
            //    ViedocMeLoginPage = PageObjectFactory.Init<ViedocMeLoginPage>(driver);
            //    Wait.UntilOrThrow(() => ViedocMeLoginPage.Displayed, 60, 2000, "Wait for ViedocMe Login Page");
            //}
            //else
            //{
            //    ViedocLoginPage = PageObjectFactory.Init<LoginPage>(driver);
            //    Wait.UntilOrThrow(() => ViedocLoginPage.Displayed, 30, 2000, "Wait for Viedoc Login Page");
            //}

        }


    }
}