using Framework.Extensions;
using Framework.PageObjects;
using Framework.WaitHelpers;
using Framework.WebDriver;
using Viedoc.viedoc.pages.components.elements;

namespace Viedoc.viedoc.pages.components.modals
{
    /// <summary>
    ///Modal Page base class
    /// </summary>
    [Locator(".fancybox-wrap .modal")]
    public abstract class ModalPage : PageObject
    {
        [Locator(".close")] protected virtual Button _close { get; set; }
        [Locator(".save")] protected virtual Button btnSave { get; set; }

        public virtual void Save()
        {
            Wait.UntilOrThrow(() => btnSave.Displayed, 10, 1000,"Wait for save button to be displayed");
            btnSave.Click();
        }

        public virtual void SaveOrClose()
        {
            if(btnSave.Exists() && btnSave.Displayed)
               btnSave.Click();
            else
            {
                Close();
            }
        }


        public virtual void Close()
        {
            Log.Debug("Closing modal");

            if (WebDriver.IsMobile())
            {
                _close.JavaClick();
                System.Threading.Thread.Sleep(500);
                this.WaitUntilGone();
            }
            else
            if (WebDriver.BrowserType == Driver.Browser.Firefox)
            {
                try
                {
                    System.Threading.Thread.Sleep(300);
                    _close.AjaxClick();
                    this.WaitUntilGone();
                }
                catch
                {
                    System.Threading.Thread.Sleep(1000);
                    _close.AjaxClick();
                    this.WaitUntilGone();
                }
                finally
                {
                    System.Threading.Thread.Sleep(1000);
                }
            }
            else
            {
                Wait.UntilOrThrow(() => _close.Displayed, 10, 1000, "Wait for clos button to be displayed");
                _close.AjaxClick();
            }
        }
    }
}