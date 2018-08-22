using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Threading;
using System.Threading.Tasks;
using Castle.Core.Internal;
using Framework.Exceptions;
using Framework.Extensions;
using Framework.Utils;
using Framework.WaitHelpers;
using Framework.WebDriver;
using Framework.TimeoutManagement;
using NLog;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.MultiTouch;
using OpenQA.Selenium.Interactions;

using Polly;
using Polly.Retry;

namespace Framework.PageObjects
{

    /// <summary>
    /// Base class of all PageObject
    /// </summary>
    public class PageObject : IParent
    {

        protected Logger Log { get; private set; }

        protected static Logger NotesLog = LogManager.GetLogger("Notes");


        public string WindowHandle { get; set; } = null;

        protected readonly List<KeyValuePair<Enum,Func<PageObject>>> Translation = new List<KeyValuePair<Enum, Func<PageObject>>>();

        public void AddTranslation(Enum e, Func<PageObject> p)
        {
             Translation.Add(new KeyValuePair<Enum, Func<PageObject>>(e,p));    
        }

        public TU SwitchToNewWindow<TU>(Action action, string nickName) where TU : PageObject, new()
        {
            return WebDriver.SwitchToNewWindow<TU>(action, nickName);
        }

        private List<KeyValuePair<Enum, Func<PageObject>>> GetTranslationObjects(params Enum[] keys)
        {
            var list = new List<KeyValuePair<Enum, Func<PageObject>>>();
            if (keys.IsNullOrEmpty())
            {
                list = Translation.ToList();
            }
            else
            {
                foreach (var key in keys)
                {
                    list.AddRange(Translation.Where(kvp => Equals(kvp.Key, key)).ToList());
                }
            }
            return list;
        }

        public void SwipeRight()
        {
            var myMobilePoint = MobilePoint();
            var p2 = new Point
            {
                Y = myMobilePoint.Y,
                X = myMobilePoint.X + 500
            };
            Swipe(myMobilePoint,p2);
        }

        public void SwipeLeft()
        {
            var myMobilePoint = MobilePoint();
            var p2 = new Point
            {
                Y = myMobilePoint.Y,
                X = myMobilePoint.X - 500
            };
            Swipe(myMobilePoint, p2);
        }

        public void Swipe(Point p1, Point p2, int timeout = 500)
        {

            Log.Debug($"Swiping from {p1.X},{p1.Y} to {p2.X},{p2.Y}");
            var touchAction = WebDriver.TouchAction;
            var swipeaction = touchAction.Press(p1.X, p1.Y).Wait(timeout).MoveTo(p2.X, p2.Y).Release();
            swipeaction.Perform();
            Log.Debug($"Swiping from {p1.X},{p1.Y} to {p2.X},{p2.Y} completed");
        }

        private Point MobilePoint()
        {
            var multipleTabsAdder = WebDriver.WindowHandles.Count > 1 ? 60 : 0;
            Log.Info($"MultipleTabsAdder = {multipleTabsAdder}");

            var pos1 = Location;
            Log.Debug($"Initial Center is {pos1.X},{pos1.Y}");

            if (pos1.Y == 0)
            {
                Log.Debug("Wait for Y to be > 0") ;

                var ans = Wait.Until(() =>
                {
                    var ypos = Location.Y;
                    Log.Info($"Y - pos == {ypos}");
                    return ypos > 0;
                }, 10, 2000);

                if (!ans)
                {
                    Log.Debug("Wait for Y to be > 0 never occured, we have probably scrolled to top");
                }

            }

            pos1 = Location;
            var centerX = pos1.X + Size.Width / 2;
            var centerY = pos1.Y + Size.Height / 2;
            centerY = centerY + 60;// 64;
            centerY = centerY + multipleTabsAdder;
          
            Log.Debug($"New Center is {centerX},{centerY}");
            var p = new Point(centerX,centerY);
            return p;
        }

        public void Tap()
        {
            if (! WebDriver.IsMobile())
            {
                throw new Exception("Can not tap unless Appium driver");
            }

            if (WebDriver.IsAndroid())
            {
                Click();
                return;
            }


            int width;
            int height;

            try
            {
                width = WebDriver.Manage().Window.Size.Width;
                height = WebDriver.Manage().Window.Size.Height;
            }
            catch
            {
                var sz = WebDriver.Capabilities.GetCapability("deviceScreenSize");
                Log.Info($"Device screen size = {sz}");
                var szx = sz.ToString().Split('x').First();
                var szy = sz.ToString().Split('x').Last();
                width = int.Parse(szy);
                height = int.Parse(szx);
            }



            var p = MobilePoint();

            var swipeY = height - 40; 

            Log.Debug($"swipeY is {swipeY}");

            if (p.X > width)
            {
                Log.Debug($"Center is > width... {p.X} > {width}");

                Log.Debug($"Swiping from 600,{swipeY} to 100,{swipeY}");
                var swipeaction = new TouchAction(WebDriver.AppiumDriver).Press(600, swipeY).Wait(100).MoveTo(100, swipeY)
                    .Release();
                swipeaction.Perform();
                Log.Debug($"Swiping from 600,{swipeY} to 100,{swipeY} performed");
                System.Threading.Thread.Sleep(1000);
                p = MobilePoint();
            }

            if (p.X < 0)
            {
                Log.Debug($"Center is < 0... {p.X}");
                Log.Debug($"Swiping from 400,{swipeY} to 100,{swipeY}");

                var swipeaction = new TouchAction(WebDriver.AppiumDriver).Press(100, swipeY).Wait(2000).MoveTo(400, swipeY)
                    .Release();
                swipeaction.Perform();
                Log.Debug($"Swiping from 400,{swipeY} to 100,{swipeY} performed");
                System.Threading.Thread.Sleep(1000);
                p = MobilePoint();
            }

            Log.Debug("Doing tap action");
            var tapaction = new TouchAction(WebDriver.AppiumDriver).Tap(p.X, p.Y, 1);
            tapaction.Perform();
            Log.Debug("Tap action completed");

        }





      

        /// <summary>
        /// 
        /// </summary>
        /// <param name="labelKeys"></param>
        public void TranslateIt(params Enum[] labelKeys)
        {
            try
            {
                if (!TranslationManager.Active) return;
                var pos = GetTranslationObjects(labelKeys);
                if (!pos.Any() && labelKeys.Any())
                {
                    foreach (var l in labelKeys)
                    {
                       Log.Warn($"Translation key missing: {l}");
                       NotesLog.Warn($"Translation key missing: {l} , intentional ?");
                    }
                    throw new Exception("No matching translation key found, ignore this one");
                }

                foreach (var po in pos)
                {
                    TranslationManager.Translate(po);
                }
            }
            catch (Exception e)
            {
                Log.Warn(e, "Unregistered labelKeys sent in to TranslateIt-method");
                NotesLog.Warn(e, "Unregistered labelKeys sent in to TranslateIt-method");
            }
        }

        public virtual void TranslateRegistration()
        {
            
        }

        /// <summary>
        /// Parent of this PageObject ( Driver or other PageObject)
        /// </summary>
        protected internal IParent Parent { get; set; }

        /// <summary>
        /// Locator for this PageObject
        /// </summary>
        public LocatorAttribute Locator { get; set; }
        public List<LocatorAttribute> Locators { get; set; }


        /// <summary>
        /// The currently cached web element
        /// </summary>
        protected internal IWebElement CachedElement;

        public virtual Func<string, LocatorAttribute> LocatorFunc { get; set; }


        /// <summary>
        ///   The frameworks internal Webdriver, NOT the Selenium Webdriver
        /// </summary>
        /// <returns>
        ///   The <see cref="Driver" />.
        /// </returns>
        public Driver WebDriver => Parent.WebDriver;


        /// <summary>
        /// Return me as a SearchContext object
        /// </summary>
        //public ISearchContext SearchContext => DoAction(i => i);
        public ISearchContext SearchContext => GetWrappedElement();

        /// <summary>
        /// Clear element
        /// </summary>
        public virtual void Clear()
        {
            {
                Log.Info("Clearing...");
                DoAction(i => i.Clear());
                Log.Info("Clearing done");
            }
        }

        /// <summary>
        /// Send keystrokes to element
        /// </summary>
        public void SendKeys(string text)
        {
            Log.Info("Sending Keys...");
            Log.Info($"Text to send : '{text}'");

            if (text == null)
            {
                Log.Info("Text string is null, ignore Sending Keys...");
                return;
            }

            if (WebDriver.BrowserType == Driver.Browser.InternetExplorer)
            {
                DoAction(i => i.SendKeys(text),60);
            }
            else
            {
                DoAction(i => i.SendKeys(text));
            }
            Log.Info("Sending Keys done");
        }

        public void JavaSetAttribute(string text, string attr = "value")
        {
            Log.Info($"Set attribute {attr} using javascript...");
            Log.Info($"Text to set : '{text}'");
            DoAction(i => WebDriver.ExecuteJavascript($"arguments[0].setAttribute('{attr}','{text}');", i));
            Log.Info("Set text using javascript done");
        }

        public void JavaSetValue(string text)
        {
            Log.Info("Set value using javascript...");
            Log.Info($"Value to set  : '{text}'");
            DoAction(i => WebDriver.ExecuteJavascript($"arguments[0].value='{text}';", i));
            //DoAction(i => WebDriver.ExecuteJavascript($"$('textarea').val('{text}');"));
            Log.Info("Set text using javascript done");
        }

        public void JavaSetInnerHtml(string text)
        {
            Log.Info("Set text using javascript to InnerHTML...");
            Log.Info($"Text to set : '{text}'");
            DoAction(i => WebDriver.ExecuteJavascript($"arguments[0].innerHTML='{text}';", i));
            Log.Info("Set text using javascript done");
        }

        public bool IsValid()
        {
            try
            {
                Log.Debug("Checking if I'm still valid");
                var me = CachedElement;
               

                if (me != null)
                {
                    Log.Debug($"me != null => {me}");
                    var meCandidates = FindMeCandidates();
                    Log.Debug("Check if I am in the collection found");
                    if (!WebDriver.IsMobile())
                    {
                        if (meCandidates.Contains(me))
                        {
                            Log.Info("My cache is still a valid element, my parent found exactly me.");
                            return true;
                        }
                        Log.Debug("No, I am not in the collection found");

                        Log.Info(
                            meCandidates.Count > 0
                                ? "I was found but I'm now a new element!"
                                : "No candidates matching me were found.");
                    }
                    else
                    {
                        Log.Debug("Check if I am in the collection found for a mobile...");
                        return meCandidates.Count > 0  && meCandidates.Count(mes => mes.TagName == me.TagName)>0;                
                    }
                }
            }
            catch (StaleElementReferenceException e)
            {
                Log.Info(e, "No one there to find me...so I'm gone too.");
            }
            catch (Exception e)
            {
                Log.Warn(e, $"Unknown exception !{e.Message}");
                Log.Warn(e.InnerException);
                Log.Warn(e.StackTrace);
            }

            return false;
        }

        /// <summary>
        /// Click on element
        /// </summary>
        public virtual void Click()
        {
             
            if (WebDriver.BrowserType == Driver.Browser.Edge)
                ScrollIntoView();

                if (WebDriver.BrowserType == Driver.Browser.InternetExplorer ||
                    WebDriver.BrowserType == Driver.Browser.IE)
                {
                    Log.Info("IE Clicking...");
                    DoAction(i => i.Click(), 120);
                    Log.Info("IE Clicking done");
                }
                else
                {
                    Log.Info("Clicking...");
                    DoAction(i => i.Click(), 120);
                    Log.Info("Clicking done");
                }
        }

        public virtual void ClickAndAcceptCommandTimeout()
        {

            if (WebDriver.BrowserType == Driver.Browser.Edge)
                ScrollIntoView();

            try
            {
                if (WebDriver.BrowserType == Driver.Browser.InternetExplorer ||
                    WebDriver.BrowserType == Driver.Browser.IE)
                {
                    Log.Info("IE Clicking...");
                    DoAction(i => i.Click(), 60);
                    Log.Info("IE Clicking done");
                }
                else
                {
                    Log.Info("Clicking...");
                    DoAction(i => i.Click(), 60);
                    Log.Info("Clicking done");
                }
            }
            catch(Exception e)
            {
                Log.Info("Got an error in click...");
                Log.Warn($"{e.Message}");
                Log.Warn($"{e.InnerException}");
            }
            finally
            {
                Log.Info("Returning from click, assuming we at least clicked the element");
            }
        }

        /// <summary>
        /// Click on element
        /// </summary>
        public virtual void Click(int doActionTimeout)
        {

            if (WebDriver.BrowserType == Driver.Browser.Edge)
                ScrollIntoView();

            if (WebDriver.BrowserType == Driver.Browser.InternetExplorer || WebDriver.BrowserType == Driver.Browser.IE)
            {
                Log.Info("IE Clicking...");
                DoAction(i => i.Click(), 60);
                Log.Info("IE Clicking done");
            }
            else
            {
                Log.Info($"Clicking with DoActionTimmeout = {doActionTimeout}...");
                DoAction(i => i.Click(), doActionTimeout);
                Log.Info($"Clicking done with DoActionTimmeout = {doActionTimeout}...");
            }
        }


        /// <summary>
        /// Click on element and wait for ajax request to become 0
        /// </summary>
        public virtual void AjaxClick(int timeout = 10)
        {
            Log.Info("Clicking and wait for Ajax requests...");
            Click();
            Log.Info("Clicking completed in AjaxClick operation...");
            this.WaitForAjax(timeout);
            Log.Info("Clicking and wait for Ajax requests done");
        }

        public virtual void AjaxJavaClick()
        {
            Log.Info("Clicking using Javascript and wait for Ajax requests...");
            JavaClick();
            this.WaitForAjax();
            Log.Info("Clicking using Javascript and wait for Ajax done");
        }

        /// <summary>
        /// Click on element using javascript
        /// </summary>
        public virtual void JavaClick()
        {
            Log.Info("Clicking using javascript...");
            DoAction(i => WebDriver.ExecuteJavascript("arguments[0].click();", i));
            Log.Info("Clicking using javascrip done");
        }

        /// <summary>
        /// Get text (innerHTML) of element
        /// </summary>
        public string JavaText()
        {
            Log.Info("Getting text using javascript...");
            var r = DoAction(i => WebDriver.ExecuteJavascript<string>("return arguments[0].innerHTML;", i));
            Log.Info($"Getting text using javascript done , returned : {r}");
            return r;
        }

        /// <summary>
        /// Submit (the WebDriver implementation)
        /// </summary>
        public virtual void Submit()
        {
            Log.Info("Submitting...");
            DoAction(i => i.Submit());
            Log.Info("Submitting done");
        }

        /// <summary>
        /// Get the style value for a given key
        /// </summary>
        /// <param name="key"></param>
        /// <param name="removeUnitFromValue"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public T GetStyleValue<T>(string key, bool removeUnitFromValue = true)
        {
            Log.Info($"Get style value '{key}'");
            var style = GetAttribute("style");
            var kvps = style.Split(';');
            foreach (var kvp in kvps)
            {
                var vars = kvp.Split(':');
                if (vars[0].Trim() != key)
                {
                    continue;
                }
                var str = vars[1].Trim();
                // Remove "px" and "%"
                if (removeUnitFromValue)
                    str = str.Replace("px", "")
                             .Replace("%", "");

                Log.Info($"Get style value '{key}' => '{str}'");
                return (T)Convert.ChangeType(str, typeof(T));
            }
            throw new Exception($"Style parameter {key} not found");
        }

        private RetryPolicy GetRetryPolicyAsync()
        {
            Log.Debug("Creating Retry policy");
            var retries = 0;
            var waitAndRetryPolicy = Policy
                .Handle<NotFoundException>()
                .Or<NoSuchElementException>()
                .Or<StaleElementReferenceException>()
                .Or<ElementNotVisibleException>()
                .Or<InvalidOperationException>()
                .Or<InvalidElementStateException>()
                .Or<WebDriverException>()
                .Or<InvalidCastException>()
                .WaitAndRetryForeverAsync(
                    (attempt, context) => TimeSpan.FromMilliseconds(Math.Min(Math.Pow(1.8, attempt) * 100, 1000)),
                    (exception, calculatedWaitDuration, context) =>
                    {
                        retries++;
                        Log.Info($"Retries = {retries} for exception {exception}");

                        // Either me or a parent is stale 
                        if (exception.GetType() == typeof(StaleElementReferenceException))
                        {
                            Log.Info("Clearing cached element");
                            CachedElement = null;
                        }
                        else if (exception.GetType() == typeof(InvalidOperationException) && exception.Message.Contains("Element reference not seen before"))
                        {
                            Log.Info("Firefox issue ! , clearing cache");
                            CachedElement = null;
                        }
                        else if (exception.GetType() == typeof(InvalidOperationException) && exception.Message.Contains("Element does not exist in cache"))
                        {
                            Log.Info("Appium issue ! , clearing cache");
                            CachedElement = null;
                        }
                        else if (exception.GetType() == typeof(InvalidOperationException) && exception.Message.Contains("Element is no longer attached to the DOM"))
                        {
                            Log.Info("Appium issue ! , clearing cache");
                            CachedElement = null;
                        }
                        else if (exception.GetType() == typeof(NoSuchElementException) && exception.Message.Contains("Web element reference not seen before"))
                        {
                            Log.Info("Firefox issue ! , clearing cache");
                            CachedElement = null;
                        }
                        context["Err"] = exception;
                    });
            Log.Debug("Creating Retry policy completed");
            return waitAndRetryPolicy;
        }

        private RetryPolicy GetRetryPolicy()
        {
            Log.Debug("Creating Retry policy");
            var retries = 0;
            var waitAndRetryPolicy = Policy
                .Handle<NotFoundException>()
                .Or<NoSuchElementException>()
                .Or<StaleElementReferenceException>()
                .Or<ElementNotVisibleException>()
                .Or<InvalidOperationException>()
                .Or<InvalidElementStateException>()
                .Or<WebDriverException>()
                .Or<InvalidCastException>()
                .WaitAndRetryForever(             
                    (attempt,context) => TimeSpan.FromMilliseconds(Math.Min(Math.Pow(1.8, attempt) * 100,1000)),
                    (exception, calculatedWaitDuration,context) =>
                    {
                        retries++;
                        Log.Info($"Retries = {retries} for exception {exception}");

                        // Either me or a parent is stale 
                        if (exception.GetType() == typeof(StaleElementReferenceException))
                        {
                            Log.Info("Clearing cached element");
                            CachedElement = null;
                        }
                        else if (exception.GetType() == typeof(InvalidOperationException) && exception.Message.Contains("Element reference not seen before"))
                        {
                            Log.Info("Firefox issue ! , clearing cache");
                            CachedElement = null;
                        }
                        else if (exception.GetType() == typeof(InvalidOperationException) && exception.Message.Contains("Element does not exist in cache"))
                        {
                            Log.Info("Appium issue ! , clearing cache");
                            CachedElement = null;
                        }
                        else if (exception.GetType() == typeof(InvalidOperationException) && exception.Message.Contains("Element is no longer attached to the DOM"))
                        {
                            Log.Info("Appium issue ! , clearing cache");
                            CachedElement = null;
                        }
                        else if (exception.GetType() == typeof(NoSuchElementException) && exception.Message.Contains("Web element reference not seen before"))
                        {
                            Log.Info("Firefox issue ! , clearing cache");
                            CachedElement = null;
                        }
                        context["Err"] = exception;
                    });
            Log.Debug("Creating Retry policy completed");
            return waitAndRetryPolicy;
        }

 
        private RetryPolicy SingleGetRetryPolicy()
        {
            var retries = 0;
            var waitAndRetryPolicy = Policy
                .Handle<NotFoundException>()
                .Or<NoSuchElementException>()
                .Or<StaleElementReferenceException>()
                .Or<ElementNotVisibleException>()
                .Or<InvalidOperationException>()
                .Or<InvalidElementStateException>()
                .Or<WebDriverException>()
                .WaitAndRetryForever(
                attempt => TimeSpan.FromMilliseconds(Math.Pow(1.8, attempt) * 100),
                (exception, calculatedWaitDuration) =>
                {
                    retries++;
                    Log.Info($"Retries = {retries} for exception {exception}");

                    if (exception.GetType() == typeof(NoSuchElementException))
                    {
                        Log.Info("Element not found, aborting");
                        CachedElement = null;
                        throw new NoSuchElementException("FROM INSIDE");
                    }

                    if (exception.GetType() == typeof(MultipleElementHitsException))
                    {
                        Log.Info("Multiple elements found, aborting");
                        CachedElement = null;
                        throw new MultipleElementHitsException("FROM INSIDE");
                    }

                    // Either me or a prent is stale 
                    if (exception.GetType() == typeof(StaleElementReferenceException))
                    {
                        Log.Info("Clearing cached element");
                        CachedElement = null;
                    }

                    if (exception.GetType() == typeof(InvalidOperationException) && exception.Message.Contains("Element is no longer attached to the DOM"))
                    {
                        Log.Info("Appium issue ! , clearing cache");
                        CachedElement = null;
                    }

                    //LogManager.Flush();

                });
            return waitAndRetryPolicy;
        }

        protected virtual void DoAction2(Action<CancellationToken> action)
        {
            Log.Info("Starting simple action...");
            CancellationTokenSource src = new CancellationTokenSource();

            CancellationToken cancellationToken = src.Token;

            var timeoutPolicy = FaultHandling.TimeoutPolicy(6);
            var waitAndRetryPolicy = GetRetryPolicy();

            var wrap = Policy.Wrap(timeoutPolicy, waitAndRetryPolicy);

            var result = wrap.ExecuteAndCapture(action,cancellationToken);//, src.Token);
            if (result.Outcome == OutcomeType.Successful)
            {
                Log.Info("Simple Action completed in time");
                return;
            }
            Log.Warn($"Simple Action {action} failed");
            //src.Cancel();
            throw result.FinalException;
        }

        protected virtual void DoAction3(Action action)
        {
            Log.Info("Starting simple action...");
          
            var timeoutPolicy = FaultHandling.TimeoutPolicy(6);
            var waitAndRetryPolicy = GetRetryPolicy();

            var wrap = Policy.Wrap(timeoutPolicy, waitAndRetryPolicy);
            var result = wrap.ExecuteAndCapture(action);
            if (result.Outcome == OutcomeType.Successful)
            {
                Log.Info("Simple Action completed in time");
                return;
            }
            Log.Warn($"Simple Action {action} failed");
            //src.Cancel();
            throw result.FinalException;
        }

        protected virtual void DoAction(Action<IWebElement> action, int timeout = 20)
        {
            Log.Info("Starting...");
            var timeoutPolicy = FaultHandling.TimeoutPolicy(timeout);
            var waitAndRetryPolicy = GetRetryPolicy();
            var wrap = Policy.Wrap(timeoutPolicy, waitAndRetryPolicy);
            var result = wrap.ExecuteAndCapture(()=>action(GetWrappedElement()));
            if (result.Outcome == OutcomeType.Successful)
            {
                Log.Info("Action completed in time");
                return;
            }
            Log.Warn($"Action {action} failed");
            throw result.FinalException;
        }
        /// <summary>
        /// Execute an action waiting for it to be successfully executed by retrying
        /// Returns answer of type T
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="func"></param>
        /// <returns></returns>
        public virtual T DoAction<T>(Func<IWebElement, T> func)
        {
            var timeoutPolicy = FaultHandling.TimeoutPolicy();
            var waitAndRetryPolicy = GetRetryPolicy();
            var wrap = Policy.Wrap(timeoutPolicy, waitAndRetryPolicy);

            var result = wrap.ExecuteAndCapture(() => func(GetWrappedElement()));
            if (result.Outcome == OutcomeType.Successful)
            {
                Log.Info("Func completed in time");
                return result.Result;
            }

            Log.Warn("Func did not complete in time!");
            throw result.FinalException;
        }


        protected virtual void DoActionQuickPeek(Action<IWebElement> action)
        {
            Log.Info("Starting...");
            var timeoutPolicy = FaultHandling.TimeoutPolicy(5);
            var waitAndRetryPolicy = SingleGetRetryPolicy();
            var wrap = Policy.Wrap(timeoutPolicy, waitAndRetryPolicy);
            var result = wrap.ExecuteAndCapture(() => action(GetWrappedElement()));
            if (result.Outcome == OutcomeType.Successful)
            {
                Log.Info("Action completed in time");
                return;
            }
            Log.Warn($"Action {action} failed");
            throw result.FinalException;
        }

        /// <summary>
        /// Execute an action waiting for it to be successfully executed by retrying
        /// Void return value
        /// </summary>
        protected virtual T DoActionQuickPeek<T>(Func<IWebElement, T> func)
        {
            var timeoutPolicy = FaultHandling.TimeoutPolicy(10);
            var waitAndRetryPolicy = SingleGetRetryPolicy();
            var wrap = Policy.Wrap(timeoutPolicy, waitAndRetryPolicy);
            var value = wrap.Execute(() => func(GetWrappedElement()));
            //LogManager.Flush();
            return value;
        }



        public bool TrySendKeys(string text)
        {
            try
            {
                Log.Info("Checking if I can type...");
                DoActionQuickPeek(i => i.SendKeys(text));
            }
            catch (MultipleElementHitsException e)
            {
                Log.Info(e, $"Checking if I Exist => {e.Message}");
                throw;
            }
            catch (Exception e)
            {
                Log.Info(e, $"Failed to type => {e.Message}");
                return false;
            }
            return true;
        }

        public bool Exists()
        {
            try
            {
                Log.Info("Checking if I Exist...");
                var value = DoActionQuickPeek(i => i.TagName);
                Log.Info($"Checking if I Exist => {value}");
                return value.Trim().Length>0;
            }
            catch (MultipleElementHitsException e)
            {
                Log.Info(e, $"Checking if I Exist (MultipleElementHitsException) => {e.Message}");
                throw;
            }
            catch (Exception e)
            {
                Log.Info(e, $"Checking if I Exist => {e.Message}");
            }
            return false;
        }


        /// <summary>
        /// Get attribute value of element
        /// </summary>
        public string GetAttribute(string attributeName)
        {
            Log.Info($"Fetching attribute '{attributeName}'...");
            var text = DoAction(i => i.GetAttribute(attributeName));
            Log.Info($"Fetching attribute '{attributeName}' => '{text}'");
            return text;
        }

        /// <summary>
        /// Get CSS value of element
        /// </summary>
        public string GetCssValue(string cssValue)
        {
            Log.Info($"Fetching CSS Value '{cssValue}'...");
            var value = DoAction(i => i.GetCssValue(cssValue));
            Log.Info($"Fetching CSS value '{cssValue}' => '{value}'");
            return value;
        }

        /// <summary>
        /// Check if element contains a class, ignoring case
        /// </summary>
        public virtual bool HasClassName(string className)
        {
            Log.Info($"Checking if I have class '{className}'...");
            var classes = GetAttribute("class");
            var classnames = classes.ToLower().Split(' ').ToList();
            Log.Debug($"Classname count : '{classnames.Count}'");
            foreach (var classname in classnames)
            {
                Log.Debug($"Class : '{classname}'");
            }
            var hasClass = classnames.Contains(className.ToLower());
            Log.Info($"Checking if I have class '{className}' => {hasClass}");
            return hasClass;
        }

        public void ActionClick()
        {
            Log.Info("Action clicking...");
            DoAction(i => Actions.Click(i).Build().Perform());
            Log.Info("Action clicking done...");
        }

        public void ActionDoubleClick()
        {
            Log.Info("Action double clicking...");
            DoAction(i => Actions.DoubleClick(i).Build().Perform());
            Log.Info("Action double clicking done...");
        }

        public void ActionSendKeys(string text)
        {
            Log.Info("Action SendKeys...");
            DoAction(i => Actions.SendKeys(i,text).Build().Perform());
            Log.Info("Action SendKeys done...");
        }

        public void ActionMoveTo()
        {
            Log.Info("Action moving...");
            DoAction(i => Actions.MoveToElement(i).Build().Perform());
            Log.Info("Action moving done...");
        }

        public void ActionMoveTo(int offsetX,int offsetY)
        {
            Log.Info($"Action moving with offset {offsetX},{offsetY}...");
            DoAction(i => Actions.MoveToElement(i,offsetX,offsetY).Build().Perform());
            Log.Info("Action moving done...");
        }

        /// <summary>
        /// Wrapper of the Actions class in WebDriver, such as MoveToElement, Click etc
        /// </summary>
        public Actions Actions => WebDriver.Actions;

        /// <summary>
        /// Scroll element into view (javascript)
        /// </summary>
        public virtual void ScrollIntoView()
        {
            Log.Info("Scrolling me into view...");
            DoAction(i => WebDriver.ExecuteJavascript("arguments[0].scrollIntoView(false)", i));
            Log.Info("Scrolling me into view done...");
        }

        /// <summary>
        /// Get Tagname of element
        /// </summary>
        public string TagName
        {
            get
            {
                Log.Info("Fetching my Tag name...");
                var value = DoAction(i => i.TagName);
                Log.Info($"Fetching my Tag name => '{value}'");
                return value;
            }
        }

        /// <summary>
        /// Get text of element
        /// </summary>
        public virtual string Text
        {
            get
            {
                Log.Info("Fetching text...");
                var text = DoAction(i => i.Text);
                Log.Info($"Fetching text => '{text}'");
                return text;
            }
        }

        /// <summary>
        /// Get text of element
        /// </summary>
        public virtual string TrimmedText => Text.Trim();

        /// <summary>
        /// Check if element enabled
        /// </summary>
        public virtual bool Enabled
        {
            get
            {
                Log.Info("Checking if I am enabled...");
                var value = DoAction(i => i.Enabled);
                Log.Info($"Checking if I am enabled  => {value}");
                return value;
            }
        }

        /// <summary>
        /// Check if element selected
        /// </summary>
        public virtual bool Selected
        {
            get
            {
                Log.Info("Checking if I am selected...");
                var value = DoAction(i => i.Selected);
                Log.Info($"Checking if I am selected  => {value}");
                return value;
            }
        }
        public ILocatable AsLocatable
        {
            get
            {
                Log.Info("Getting AsLocatable...");
                var value = DoAction(i => (ILocatable)i);
                Log.Info("Getting AsLocatable done");
                return value;
            }
        }

        /// <summary>
        /// Get element Location
        /// </summary>
        public Point Location
        {
            get
            {
                Log.Info("Getting my location...");
                var value = DoAction(i => i.Location);
                Log.Info($"Getting my location (X,Y) => {value.X},{value.Y}");
                return value;
            }
        }

        /// <summary>
        /// Get element size
        /// </summary>
        public Size Size
        {
            get
            {
                Log.Info("Getting my size...");
                var value = DoAction(i => i.Size);
                Log.Info($"Getting my size (W,H)=> {value.Width},{value.Height}");
                return value;
            }
        }

        /// <summary>
        /// Check if element displayed
        /// </summary>
        public virtual bool Displayed
        {
            get
            {
                if(WebDriver.BrowserType == Driver.Browser.Edge)
                    ScrollIntoView();

                Log.Info("Checking if I am displayed...");
                var value = DoAction(i => i.Displayed);
                Log.Info($"Checking if I am displayed => {value}");
                return value;

            }
        }

        private bool ExecuteFindElements(ISearchContext context,By locator,out ReadOnlyCollection<IWebElement> elements,int timeoutInMs = 5000)
        {
            var myElem = new ReadOnlyCollection<IWebElement>(new List<IWebElement>());
            elements = myElem;

            var task = Task.Factory.StartNew(() =>
            {
                myElem = context.FindElements(locator);
            });

            var t0 = DateTime.Now;
            var done = task.Wait(timeoutInMs);
            var t1 = DateTime.Now - t0;
            if (t1 > TimeSpan.FromSeconds(2))
                Log.Debug($"FindElements took {t1.TotalMilliseconds} ms, locator {locator}");

            if (done)
            {
                Log.Debug("Calling currentSearchContext.FindElements() was OK");
                elements = myElem;
                return true;
            }
           
            Log.Warn("Find Elements never completed");
            return false;
        }

        /// <summary>
        /// Normal FindMeCandidates_ for single element
        /// </summary>
        /// <returns></returns>
        internal ReadOnlyCollection<IWebElement> FindMeCandidates_()
        {
            ReadOnlyCollection<IWebElement> myElem;

            try
            {
                Log.Debug($"Need parent [{Parent}]");
                var currentSearchContext = Parent.SearchContext;
                Log.Debug("Got parent as SearchContext");

                if (currentSearchContext == null)
                {
                    Log.Warn("currentSearchContext == null  !!");
                    throw new Exception("Can not locate me, parent == null,XXX");
                }
                Log.Debug($"Generating webdriver locator from {Locator}");
                var locator = Locator.ToWebdriverLocator();
                Log.Debug($"Generated webdriver locator : {locator}");

                Log.Debug("Calling currentSearchContext.FindElements()");
                try
                {
                    Log.Debug($"currentSearchContext = {currentSearchContext}");
                    
                    var status = ExecuteFindElements(currentSearchContext, locator, out myElem);
                    if (!status)
                    {
                        if (Parent is PageObject o)
                        {
                            Log.Debug("Clearing parent cache due FindElements Tiemout....");
                            o.CachedElement = null;
                        }
                    }
                }
                catch (AggregateException ae)
                {
                    Log.Debug($"Forwarding inner exception....{ae.InnerException}");
                    if (Parent is PageObject o)
                    {
                        Log.Debug("Clearing parent cache due to aggregate exception....");
                        o.CachedElement = null;
                    }
                    if (ae.InnerException != null) throw ae.InnerException;
                    throw;
                }
            }
            catch (StaleElementReferenceException e)
            {
               Log.Warn(e, $"Parent [{Parent}] is stale => I do not exist");
               ((PageObject) Parent).CachedElement = null;
                throw;
            }
            catch (NoSuchElementException)
            {
                Log.Warn($"Parent [{Parent}] not found during lookup of [{Locator}] ");
                throw;
            }
            catch (Exception e)
            {
                Log.Warn(e,$"Unknown error during FindMeCandidates : {e.InnerException}");
                throw;
            }
            return myElem;
        }

        /// <summary>
        /// Find elements that match my predicate (used for IEnumerable elements in PageObject)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="predicate"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        internal ReadOnlyCollection<IWebElement>  FindMeCandidates_<T>(List<Func<T, bool>> predicate, int index = -1) where T:PageObject, new()
        {
            ReadOnlyCollection<IWebElement> meCandidates;

            // Get the elements as normal...
            try
            {
                Log.Debug("Getting me-candidates...");
                meCandidates = FindMeCandidates_();
                Log.Debug($"Got {meCandidates.Count} intial candidates");
            }
            catch(Exception e)
            {
                // Forward the error
                Log.Warn(e, "Predicate : FindmeElements from base call returned an exception, forwarding...");
                Log.Warn(e.Message);

                throw;
            }

            try
            {
                var elementsAsPageObject = meCandidates.Select(e => PageObjectFactory.CreatePageObject<T>(Parent, e)).ToList();

                Log.Debug($"Applying {predicate.Count} predicate(s)...");

                var i = 1;
                foreach (var pred in predicate)
                {
                    var originalCount = elementsAsPageObject.Count;
                    Log.Debug($"Applying predicate # {i}...on {originalCount} elements");
                    elementsAsPageObject = elementsAsPageObject.Where(pred).ToList();
                    Log.Debug($"Applying predicate # {i} left us with {elementsAsPageObject.Count}/{originalCount} candidates");
                    i++;
                }

                if (index >= 0)
                {
                    Log.Debug($"Returning item with index # {index} of available indices {elementsAsPageObject.Count-1}");
                    var mySearchElement = elementsAsPageObject[index].CachedElement;
                    Log.Debug($"Returning element '{mySearchElement}'");
                    var listOfElements = new List<IWebElement> { mySearchElement }.AsReadOnly();
                    return listOfElements;
                }

                var filteredElements = elementsAsPageObject
                    .Select(elem => elem.CachedElement)
                    .ToList()
                    .AsReadOnly();

                return filteredElements;
            }
            catch (Exception e)
            {
                Log.Warn($"Applying predicate on elements failed {e.Message}");
                Log.Warn("Throwing a NoSuchElementException");
                throw new NoSuchElementException();
            }
        }


        internal Func<ReadOnlyCollection<IWebElement>> FindMeCandidates;

        public PageObject()
        {
            Log = LogManager.GetLogger(GetType().FullName);
            FindMeCandidates = FindMeCandidates_;
        }

        public  IWebElement GetWrappedElement()
        {
            try
            {
                Log.Debug($"Locating [{this}] with [{Locator}]");
                if (Locator.UseCache && CachedElement != null)
                {
                    Log.Debug($"Using cached element for [{Locator}]");
                    return CachedElement;
                }

                Log.Debug("Calling PageObject.FindMeCandidates()");
                var meCandidates = FindMeCandidates();
                Log.Debug("Calling PageObject.FindMeCandidates() done");

                var foundElements = meCandidates.Count;

                switch (foundElements)
                {
                    case 1:
                        Log.Debug($"Found [{this}] with [{Locator}]");
                        CachedElement = meCandidates[0];
                        return meCandidates[0];
                    case 0:
                        throw new NoSuchElementException($"Failed to find [{Locator}]");
                    default:
                        Log.Warn($"Non unique hit, found {foundElements} matching elements");
                        foreach (var foundElement in meCandidates)
                        {
                            Log.Warn($@" Elem class = {foundElement.GetAttribute("class")}");
                        }

                        throw new MultipleElementHitsException($"Non unique hit for {Parent} -> {Locator}");
                }
            }
            catch (StaleElementReferenceException)
            {
                Log.Warn("Parent is stale, clearing Parent cached element");
                ((PageObject)Parent).CachedElement = null;
                throw;
            }
            catch (Exception e)
            {
                Log.Warn(e);
                throw;
            }
        }

        protected internal void SetWrappedElement(IWebElement value)
        {
            CachedElement = value;
            Locator.UseCache = true;
        }

        //protected internal virtual Dictionary<Enum, string> ItemDictionary => new Dictionary<Enum, string>();

    }
}
