using System;
using System.Collections.Generic;
using Framework.PageObjects;
using Framework.WaitHelpers;
using Framework.WebDriver;

namespace Viedoc.viedoc.pages.components
{

    [Locator("dd a")]
    public class TabControlItem : PageObject
    {
    }

    [Locator("li.active")]
    public class TabControlArea : PageObject
    {
    }

    [Locator(How.XPath, @".//*[@class='tabs']/..")]
    public class TabControl<T> : PageObject  where T : TabControlItem , new()                                             
    {
        public virtual IEnumerable<T> Tabs { get; set; }

        protected Dictionary<Enum, string> ItemDictionary = new Dictionary<Enum, string>();

        public T GetTabButton(Enum tabItem)
        {
            var str = ItemDictionary[tabItem];
            var tab = Tabs.GetElement(str);
            return tab;
        }

        // Clicks on tab and returns the Tab Area Object
        public virtual void OpenTab<TU>(Enum tabItem, out TU tabArea) where TU : TabControlArea, new()
        {
            var str = ItemDictionary[tabItem];
            var tab = Tabs.GetElement(str);
            if(WebDriver.BrowserType == Driver.Browser.InternetExplorer || WebDriver.BrowserType== Driver.Browser.IE)
               Wait.UntilOrThrow(() => tab.Displayed, message: "Wait for tab to be Displayed");
            else
              Wait.UntilOrThrow(() => tab.Displayed && tab.Enabled,message:"Wait for tab to be Enabled");

            //tab.Click();
            tab.JavaClick();
            
            tabArea = PageObjectFactory.Init<TU>(this);
            System.Threading.Thread.Sleep(1000);
        }

        // Clicks on tab and returns the Tab Area Object
        public virtual TU OpenTab<TU>(Enum tabItem) where TU : TabControlArea, new()
        {
            var str = ItemDictionary[tabItem];
            var tab = Tabs.GetElement(str);
            if (WebDriver.BrowserType == Driver.Browser.InternetExplorer || WebDriver.BrowserType == Driver.Browser.IE)
                Wait.UntilOrThrow(() => tab.Displayed, message: "Wait for tab to be Displayed");
            else
                Wait.UntilOrThrow(() => tab.Displayed && tab.Enabled, message: "Wait for tab to be Enabled");

            //tab.Click();
            tab.JavaClick();

            var tabArea = PageObjectFactory.Init<TU>(this);
            System.Threading.Thread.Sleep(1000);
            return tabArea;
        }

    }
}