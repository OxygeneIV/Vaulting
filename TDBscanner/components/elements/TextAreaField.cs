using System;
using Framework.PageObjects;
using Framework.WebDriver;

namespace Viedoc.viedoc.pages.components.elements
{
    /// <summary>
    /// Default Viedoc Textarea element
    /// </summary>
    [Locator(How.Css, "textarea")]
    public class TextAreaField : PageObject
    {
        public string TextValue => GetAttribute("value");

        public virtual void SetTextJQuery(string text)
        {
            var txtIsDisplayed = Displayed;

            if (!txtIsDisplayed)
            {
                throw new Exception("Element not found to do JQuery SetText on");
            }
            var loc = GetWrappedElement();
            var str = $"$(arguments[0]).val(\"{text}\").change();";
            Log.Info($"JQuery set text string = {str}");
            WebDriver.ExecuteJavascript(str, loc);
        }

        public virtual void SetText(string text, bool append = false)
        {
            if (WebDriver.BrowserType == Driver.Browser.InternetExplorer)
            {
                if (WebDriver.IsIe8())
                {
                    ActionDoubleClick();
                    System.Threading.Thread.Sleep(100);
                    JavaSetInnerHtml(text);
                    System.Threading.Thread.Sleep(100);
                    WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
                    //// Try get the value and check
                    var txtie8 = TextValue;
                    Log.Info($"Re-read the textfield value for IE 8 : {txtie8}");

                }
                else if (WebDriver.IsIe9())
                {
                    if (!append)
                        SetTextJQuery("");
                    // Sub

                    var text2 = text.Replace(Environment.NewLine, @"\n");

                    SetTextJQuery(text2);

                    // Try get the value and check
                    var txtie8 = TextValue;
                    Log.Info($"Re-read the textfield vale : {txtie8}");
                }
                else if (WebDriver.IsIe10())
                {
                    if (!append)
                        SetTextJQuery("");
                    // Sub

                    var text2 = text.Replace(Environment.NewLine, @"\n");

                    SetTextJQuery(text2);

                    // Try get the value and check
                    var txtie8 = TextValue;
                    Log.Info($"Re-read the textfield vale : {txtie8}");
                }
                //else if (WebDriver.IsIe9())
                //{
                //    if (!append)
                //        Clear();
                //    //Click();
                //    System.Threading.Thread.Sleep(100);
                //    JavaSetInnerHTML(text);
                //    //SendKeys(text);
                //    System.Threading.Thread.Sleep(100);
                //    WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
                //    // Try get the value and check
                //    var txtie8 = TextValue;
                //    Log.Info($"Re-read the textfield vale : {txtie8}");
                //}
                else
                {

                    var currentText = TextValue;

                    if (!append)
                        Clear();

                    SendKeys(text);
                    var txt = TextValue;
                    Log.Info($"Re-read the textfield vale : {txt}");

                    if (txt.EndsWith(text)) return;

                    Clear();
                    SendKeys(currentText + text);
                }
            }else if (WebDriver.BrowserType == Driver.Browser.Edge)
            {
                if (!append)
                    Clear();

                var value = text.Replace("\r\n", "\r");
                SendKeys(value);
                //JavaClick();
                //JavaSetValue(text);
                //WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
            }
            else
            {
                if (!append)
                    Clear();

                SendKeys(text);
                if (WebDriver.BrowserType == Driver.Browser.Safari)
                {
                    WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
                }
            }

            if (WebDriver.IsMobile())
            {
                WebDriver.HideKeyboard();// ("Done","HideKeyboardStrategy");
            }
        }

        public virtual bool IsCleared()
        {
            return TextValue == string.Empty;
        }

    }
}