using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using OpenQA.Selenium;
using System.Collections;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace Framework.PageObjects
{


    public class PageObjectListProxy<T> : IEnumerable<T> where T : PageObject, new()
    {
        protected Logger Log = LogManager.GetCurrentClassLogger();

        internal IParent Parent;
        internal LocatorAttribute Locator { get; set; }
        private readonly By _bylocator;

        public ItemLocatorAttribute ItemLocator { get; set; }

        public Dictionary<Enum,string> ItemDictionary { get; set; }


        public PageObjectListProxy(
               IParent parent,
               LocatorAttribute locatorAttribute)
        {
            Locator = locatorAttribute;
            Parent = parent;
            _bylocator = locatorAttribute.ToWebdriverLocator();
        }

        public List<Func<T, bool>> PageObjectFilter = new List<Func<T, bool>>();

        private bool ExecuteFindElements(ISearchContext context, By locator, out ReadOnlyCollection<IWebElement> elements, int timeoutInMs = 5000)
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
            {
                Log.Warn($"FindElements took {t1.TotalMilliseconds} ms, locator {Locator}");
            }
            else
            {
                Log.Debug($"FindElements took {t1.TotalMilliseconds} ms, locator {Locator}");
            }

            if (done)
            {
                Log.Debug("Calling currentSearchContext.FindElements() was OK");
                elements = myElem;
                return true;
            }

            Log.Warn("Find Elements never completed");
            return false;
        }

        private List<T> FindMeCandidates_()
        {
            ReadOnlyCollection<IWebElement> meCandidates;

            // Get the elements as normal...
            try
            {
                Log.Debug($"Getting me-candidates from predicate, locator : '{Locator}'");
                Log.Debug($"byLocator : '{_bylocator}'");

                var status = ExecuteFindElements(Parent.SearchContext, _bylocator, out meCandidates);
                if (!status)
                {
                    Log.Debug("Return NOthing found for now, FindElements failed....");
                    return new List<T>();
                }
                //meCandidates = ((PageObject)Parent).DoAction(i => i.FindElements(_bylocator));
                Log.Debug($"Got {meCandidates.Count} intial candidates");
            }
            catch (StaleElementReferenceException)
            {
                Log.Warn("Parent is stale, clearing Parent cached element");
                ((PageObject)Parent).CachedElement = null;
                throw;
            }
            catch (Exception e)
            {
                // Forward the error
                Log.Warn(e, "Predicate : FindmeElements from base call returned an exception, forwarding...");
                ((PageObject)Parent).CachedElement = null;
                Log.Warn(e.Message);

                throw;
            }

            try
            {
                Log.Debug($"Converting items to PageObjects of type {typeof(T)}...");

                var filteredElements = meCandidates.Select(e => PageObjectFactory.CreatePageObject<T>(Parent, e)).ToList();

                // Check Filter
                if (PageObjectFilter != null && PageObjectFilter.Any())
                {
                    // Test the predicate
                    try
                    {
                        Log.Debug("Applying the predicate(s)...");
                        filteredElements = PageObjectFilter.Aggregate(filteredElements,
                            (current, filter) => current.Where(filter).ToList());
                        Log.Debug($"Applying predicate(s) ok, {filteredElements.Count}/{meCandidates.Count} matched");
                    }
                    catch (Exception ex)
                    {
                        Log.Warn("Probably stale stuff during filtering, retry...");
                        Log.Warn(ex.Message);
                        Log.Warn(ex.InnerException);
                        Log.Warn(ex.StackTrace);
                        filteredElements = FindMeCandidates_();
                        Log.Debug("FilteredElements a second time completed...");
                    }
                }
                else
                {
                    Log.Debug("No predicate used, returning all elements...");
                }

                return filteredElements;
            }
            catch (Exception e)
            {
                Log.Warn($"Applying predicate on elements failed {e.Message}");
                Log.Warn("Throwing a NoSuchElementException");
                throw new NoSuchElementException();
            }
        }


        private IEnumerable<T> ElementList
        {
            get
            {
                Log.Debug("Starting ElementList...");
                var theList = FindMeCandidates_();
                Log.Debug("Returning ElementList...");
                return theList;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IEnumerator<T> GetEnumerator()
        {
            return ElementList.GetEnumerator();
        }
    }


    public class PageObjectList<T> : IReadOnlyList<T> where T : PageObject, new()
    {
        protected Logger Log = LogManager.GetCurrentClassLogger();
        private IReadOnlyList<T> _readOnlyListImplementation;// = FindMeCandidates_();

        internal IParent Parent;
        internal LocatorAttribute Locator { get; set; }
        
        private readonly By _bylocator;

        public ItemLocatorAttribute ItemLocator { get; set; }

        public Dictionary<Enum, string> ItemDictionary { get; set; }


        public PageObjectList(
               IParent parent,
               LocatorAttribute locatorAttribute)
        {
            Locator = locatorAttribute;
            Parent = parent;
            _bylocator = locatorAttribute.ToWebdriverLocator();
        }

        public List<Func<T, bool>> PageObjectFilter = new List<Func<T, bool>>();

        private bool ExecuteFindElements(ISearchContext context, By locator, out ReadOnlyCollection<IWebElement> elements, int timeoutInMs = 5000)
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
            {
                Log.Warn($"FindElements took {t1.TotalMilliseconds} ms, locator {Locator}");
            }
            else
            {
                Log.Debug($"FindElements took {t1.TotalMilliseconds} ms, locator {Locator}");
            }

            if (done)
            {
                Log.Debug("Calling currentSearchContext.FindElements() was OK");
                elements = myElem;
                return true;
            }

            Log.Warn("Find Elements never completed");
            return false;
        }


        private void FindMeCandidates_()
        {
            ReadOnlyCollection<IWebElement> meCandidates;

            // Get the elements as normal...
            try
            {
                Log.Debug($"Getting me-candidates from predicate, locator : '{Locator}'");
                Log.Debug($"byLocator : '{_bylocator}'");

                var status = ExecuteFindElements(Parent.SearchContext, _bylocator, out meCandidates);

                if (!status)
                {
                    Log.Debug("Return NOthing found for now, FindElements failed....");
                    throw new WebDriverException("Find elements took too long time....");
                }
                Log.Debug($"Got {meCandidates.Count} intial candidates");
            }
            catch (StaleElementReferenceException)
            {
                Log.Warn("Parent is stale, clearing Parent cached element");
                ((PageObject)Parent).CachedElement = null;
                throw;
            }
            catch (Exception e)
            {
                // Forward the error
                Log.Warn(e, "FindmeElements from base call returned an exception, clearing Parent cached element, forwarding...");
                ((PageObject)Parent).CachedElement = null;
                Log.Warn(e.Message);
                throw;
            }

            try
            {
                Log.Debug($"Converting items to PageObjects of type {typeof(T)}...");

                var filteredElements = (IReadOnlyList<T>)meCandidates.Select(e => PageObjectFactory.CreatePageObject<T>(Parent, e)).ToList();

                // Check Filter
                if (PageObjectFilter != null && PageObjectFilter.Any())
                {
                    // Test the predicate
                    try
                    {
                        Log.Debug("Applying the predicate(s)...");
                        filteredElements = PageObjectFilter.Aggregate(filteredElements,
                            (current, filter) => current.Where(filter).ToList());
                        Log.Debug($"Applying predicate(s) ok, {filteredElements.Count}/{meCandidates.Count} matched");
                    }
                    catch (Exception ex)
                    {
                        Log.Warn("Probably stale stuff during filtering, retry...");
                        Log.Warn(ex.Message);
                        Log.Warn(ex.InnerException);
                        Log.Warn(ex.StackTrace);
                        FindMeCandidates_();
                        Log.Debug("FilteredElements a second time completed...");
                    }
                }
                else
                {
                    Log.Debug("No predicate used, returning all elements...");
                }

                _readOnlyListImplementation = filteredElements;
            }
            catch (Exception e)
            {
                Log.Warn($"Applying predicate on elements failed {e.Message}");
                Log.Warn("Throwing a NoSuchElementException");
                throw new NoSuchElementException();
            }
        }

        public IEnumerator<T> GetEnumerator()
        {
            FindMeCandidates_();
            return _readOnlyListImplementation.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable) _readOnlyListImplementation).GetEnumerator();
        }

        public int Count => _readOnlyListImplementation.Count;

        public T this[int index] => _readOnlyListImplementation[index];
    }
}
