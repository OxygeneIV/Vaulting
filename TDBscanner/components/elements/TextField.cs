using System;
using Framework.PageObjects;
using Framework.WebDriver;

namespace Viedoc.viedoc.pages.components.elements
{
    [Locator(How.Css, "input[type=text]")]
    public class TextField : PageObject
    {
        public virtual string TextValue => GetAttribute("value");

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
            WebDriver.ExecuteJavascript(str,loc);
        }

        public virtual void SetText2_Ie9(string text, bool append = false)
        {
            if (!append)
                Clear();
            Click();
            System.Threading.Thread.Sleep(100);
            SendKeys(text);
            System.Threading.Thread.Sleep(100);
            WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
            // Try get the value and check
            var txtie8 = TextValue;
            Log.Info($"Re-read the IE 9 textfield vale : {txtie8}");
        }

        public virtual void SetText2_Ie10(string text, bool append = false)
        {
            if (!append)
                Clear();
            Click();
            System.Threading.Thread.Sleep(100);
            SendKeys(text);
            System.Threading.Thread.Sleep(100);
            WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
            // Try get the value and check
            var txtie8 = TextValue;
            Log.Info($"Re-read the IE 10 textfield vale : {txtie8}");
        }
        /// <summary>
        /// Set text in field
        /// </summary>
        /// <param name="text"></param>
        /// <param name="append"></param>
        public virtual void SetText(string text, bool append = false)
        {
            if (WebDriver.BrowserType == Driver.Browser.InternetExplorer)
            {
                if (WebDriver.IsIe8())
                {
                    ActionDoubleClick();
                    System.Threading.Thread.Sleep(100);
                    JavaSetAttribute(text);
                    System.Threading.Thread.Sleep(100);
                    //// Try get the value and check
                    var txtie8 = TextValue;
                    Log.Info($"Re-read the textfield value for IE 8 : {txtie8}");
                    WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
                }
                else if (WebDriver.IsIe9())
                {
                    if (!append)
                        SetTextJQuery("");

                    SetTextJQuery(text);

                    // Try get the value and check
                    var txtie8 = TextValue;
                    Log.Info($"Re-read the textfield vale : {txtie8}");


                    //if (!append)
                    //    Clear();
                    //Click();
                    //System.Threading.Thread.Sleep(100);
                    //SendKeys(text);
                    //System.Threading.Thread.Sleep(100);
                    //WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
                    //// Try get the value and check
                    //var txtie8 = TextValue;
                    //Log.Info($"Re-read the textfield vale : {txtie8}");
                }
                else if (WebDriver.IsIe10())
                {
                    if (!append)
                        SetTextJQuery("");

                    SetTextJQuery(text);

                    // Try get the value and check
                    var txtie8 = TextValue;
                    Log.Info($"Re-read the textfield vale : {txtie8}");
                }
                //else if (WebDriver.IsIe10())
                //{
                //    if (!append)
                //        Clear();
                //    Click();
                //    System.Threading.Thread.Sleep(100);
                //    SendKeys(text);
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
            }
            else
            {
                if (!append)
                    Clear();

                SendKeys(text);


            }

            if (WebDriver.IsMobile())
            {
                //WebDriver.IosDriver.HideKeyboard();// ("Done","HideKeyboardStrategy");
                WebDriver.HideKeyboard();
            }
        }

        /// <summary>
        /// Check if field is empty
        /// </summary>
        /// <returns></returns>
        public bool IsCleared()
        {
            return TextValue == string.Empty;
        }


        public void SetPassword(string password)
        {
            if (WebDriver.IsIe8())
            {
                SetText(password);
                WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
            }
            else if (WebDriver.IsIe10())
            {
                Click();
                System.Threading.Thread.Sleep(1000);
                SetText(password);
                System.Threading.Thread.Sleep(500);
                WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
            }
            else if (WebDriver.BrowserType == Driver.Browser.Safari)
            {
                SetText(password);
                WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
            }
            else if (WebDriver.BrowserType == Driver.Browser.iPad)
            {
                SetText(password);
                WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
            }
            else
            {
                SetText(password);
            }
        }

        public virtual void SetTextOnChange(string text)
        {
            SetText(text);
            WebDriver.ExecuteJavascript("$(arguments[0]).change();", GetWrappedElement());
        }

        public virtual bool IsReadOnly()
        {
            return Enabled == false;
        }
    }
}
