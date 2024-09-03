using System.Collections.Generic;
using Framework.PageObjects;
using Framework.WebDriver;

namespace Viedoc.viedoc.pages.components.elements
{
    /// <summary>
    /// Default Viedoc TableCell
    /// </summary>
    [Locator("td, th")]
    public class TableCell : PageObject
    {
        protected Link link_ = null;
        public IEnumerable<Link> Links;

        public virtual void LinkClick()
        {
            if (WebDriver.BrowserType == Driver.Browser.Firefox || WebDriver.IsMobile())
            {
                link_.JavaClick();
            }
            else
            {
                link_.JavaClick();
                //link_.Click();
            }
        }

        public string LinkUrl
        {
            get { return link_.GetAttribute("href"); }
        }
        public string LinkUrls
        {
            get { return link_.GetAttribute("href"); }
        }


        public void JavaSetText(string text, bool append = false)
        {
            JavaSetAttribute(text);
            WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());

        }

        public override string Text
        {
            get
            {
                Log.Info("Fetching text...");
                var text = DoAction(i => i.Text);
                Log.Info($"Fetching text => '{text}'");

                if (Driver.Browser.Edge == WebDriver.BrowserType && text.EndsWith("\r\n"))
                    text = text.TrimEnd(System.Environment.NewLine.ToCharArray());

                return text;
            }
        }
    }
}
