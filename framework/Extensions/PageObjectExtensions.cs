using System;
using System.Drawing;
using System.Threading;
using Framework.PageObjects;
using Framework.WaitHelpers;
using NLog;
using OpenQA.Selenium;

namespace Framework.Extensions
{
    public static class PageObjectExtensions
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        public static Bitmap GetElementScreenShot(this PageObject pageObject)
        {
            //pageObject.ScrollIntoView();
  
  

            var driver = pageObject.WebDriver.SearchContext;
            var sc =  ((ITakesScreenshot)driver).GetScreenshot();
            sc.SaveAsFile(@"c:\temp\img.bmp",ScreenshotImageFormat.Bmp);
            //var img = Image.FromStream(new MemoryStream(sc.AsByteArray)) as Bitmap;
            //var img = Image.FromStream(new MemoryStream(sc.AsByteArray)) as Bitmap;
            var img = Image.FromFile(@"c:\temp\img.bmp");
            if (img == null)
            {
                throw new Exception("Failed to take PageObject snapshot");
            }
            var imgsize = img.Size;
            var rec = new Rectangle(pageObject.Location, pageObject.Size);
            var bmpImage = new Bitmap(img);
            var bmp =  bmpImage.Clone(rec , bmpImage.PixelFormat);
            return bmp;
        }


        public static string HighlighElement(this PageObject pageObject, string border = "4px dashed magenta")
        {
            pageObject.ScrollIntoView();
            var me = pageObject.GetWrappedElement();
         
            const string currentBorderQuery = "return arguments[0].style.border";
            var oldBorder = "";
            try
            {
                oldBorder = pageObject.WebDriver.ExecuteJavascript<string>(currentBorderQuery, me);
                Log.Info($"Old border = '{oldBorder}'");
            }
            catch(Exception)
            {
                Log.Warn("Failed/Missing to get border property, returning null");
            }

            var bordersetting = $"arguments[0].style.border='{border}'";
            pageObject.WebDriver.ExecuteJavascript(bordersetting, me);

            Wait.UntilOrThrow(() =>
            {
                var modborder = pageObject.WebDriver.ExecuteJavascript<string>(currentBorderQuery, me);
                return modborder.Contains(border);
            }, 5, 500,$"Could not find the new border '{border}' in border property");

            return oldBorder;
        }

        /// <summary>
        /// Action method: Move mouse to element
        /// </summary>
        /// <param name="pageObject"></param>
        public static void MoveToElement(this PageObject pageObject)
        {
            Log.Info($"Move to {pageObject.Location.X},{pageObject.Location.Y}");
            Log.Info($"Size : {pageObject.Size.Width},{pageObject.Size.Height}");

            pageObject.Actions.MoveToElement(pageObject.GetWrappedElement()).Build().Perform();
        }

        /// <summary>
        ///  Move mouse to element using java
        /// </summary>
        /// <param name="pageObject"></param>
        /// <param name="target"></param>
        public static void JavaMoveToElement(this PageObject pageObject, PageObject target)
        {
            var scr =
                "function simulate(f,c,d,e){var b,a=null;for(b in eventMatchers)if(eventMatchers[b].test(c)){a=b;break}if(!a)return!1;document.createEvent?(b=document.createEvent(a),a==\"HTMLEvents\"?b.initEvent(c,!0,!0):b.initMouseEvent(c,!0,!0,document.defaultView,0,d,e,d,e,!1,!1,!1,!1,0,null),f.dispatchEvent(b)):(a=document.createEventObject(),a.detail=0,a.screenX=d,a.screenY=e,a.clientX=d,a.clientY=e,a.ctrlKey=!1,a.altKey=!1,a.shiftKey=!1,a.metaKey=!1,a.button=1,f.fireEvent(\"on\"+c,a));return!0} var eventMatchers={HTMLEvents:/^(?:load|unload|abort|error|select|change|submit|reset|focus|blur|resize|scroll)$/,MouseEvents:/^(?:click|dblclick|mouse(?:down|up|over|move|out))$/}; " +
                "simulate(arguments[0],\"mousemove\",arguments[1],arguments[2]);";

            var me = pageObject.GetWrappedElement();
            var x = target.Location.X;
            var y = target.Location.Y;
            var sz = target.Size;
            var xx = sz.Width  / 2;
            var yy = sz.Height / 2;

            pageObject.WebDriver.ExecuteJavascript(scr,me,x+xx,y+yy);
        }


        public static void JavaHover(this PageObject pageObject)
        {
            Log.Debug("Hover with java");

            var javaScript = @"if(document.createEvent){var evObj = document.createEvent('MouseEvents');evObj.initEvent('mouseover', 

            true, false); arguments[0].dispatchEvent(evObj);
        } else if(document.createEventObject) { arguments[0].fireEvent('onmouseover');
        }";


            var me = pageObject.GetWrappedElement();
            pageObject.WebDriver.ExecuteJavascript(javaScript, me);
            Log.Debug("Hover with java done");
        }

        public static void JavaFocus(this PageObject pageObject)
        {
            Log.Debug("Focus with java");
            var javaScript = "arguments[0].focus();";
            var me = pageObject.GetWrappedElement();
            pageObject.WebDriver.ExecuteJavascript(javaScript, me);
            Log.Debug("Focus with java done");
        }
        public static void WaitForAjax<T>(this T pageObject, int timeoutSecs = 10, bool throwException = false) where T : PageObject
        {
            var device = (string)pageObject.WebDriver.Capabilities.GetCapability("device");

            if (device != null && device.Contains("iPad"))
            {
                var ajaxIsComplete = pageObject.WebDriver.ExecuteJavascript<bool>("return jQuery.active == 0;");
                Log.Debug($"Ajax complete (iPad) => {ajaxIsComplete}");

                Log.Warn($"No ajax normally for device {device}");
                Thread.Sleep(500);
                return;
            }

            try
            {
                var ajaxHasBeenDetected = false;
                for (var i = 0; i < timeoutSecs; i++)
                {
                    var ajaxIsComplete = false;
                    Log.Debug($"Check Ajax at {DateTime.UtcNow.TimeOfDay}");
                    try
                    {
                        ajaxIsComplete = pageObject.WebDriver.ExecuteJavascript<bool>("return jQuery.active == 0;");
                        Log.Debug($"Ajax complete => {ajaxIsComplete}");
                    }
                    catch(Exception e)
                    {
                        Log.Warn(e, $"Ajax call said => {e.Message}");
                    }

                    if (ajaxIsComplete)
                    {
                        if(ajaxHasBeenDetected)
                            Thread.Sleep(500);
                        return;
                    }
                    Log.Debug("Ajax not completed!");
                    ajaxHasBeenDetected = true;
                    Thread.Sleep(1000);
                }
                if (throwException)
                {
                    throw new Exception("WebDriver timed out waiting for AJAX call to complete");
                }
            }
            catch(Exception e)
            {
                Log.Debug(e, $"Check Ajax Failed {e.Message}");
                throw;
            }
        }

        public static void WaitUntilGone<T>(this T pageObject, int timeout = 30, int pollIntervalMilliSeconds = 500) where T : PageObject
        {
            Log.Info($"Waiting for element {typeof(T)} to be gone !");
            var gone = Wait.Until(() => !pageObject.IsValid(),
                         timeout,
                         pollIntervalMilliSeconds);

            Log.Info($"Waiting for element gone => {gone}");

            if (gone)
            {
                pageObject.CachedElement = null;
                return;
            }

            Log.Info("Element never gone!");

            throw new Exception($"{typeof(T)} never gone!!");
        }

        public static void TakeScreenshot<T>(this T pageObject, string name) where T : PageObject
        {
            pageObject.WebDriver.TakeScreenshot(name);
        }
    }
}
