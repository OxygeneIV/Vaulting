using System;
using Framework.PageObjects;

namespace Viedoc.viedoc.pages.components.elements
{
    /// <summary>
    /// Default Viedoc input fiels item
    /// </summary>
    [Locator(How.Css, "input")]
    public class InputField : PageObject
    {
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
    }
}