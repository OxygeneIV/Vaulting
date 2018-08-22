using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using Framework.PageObjects;
using NLog;
using OpenQA.Selenium;

namespace Framework.IList
{
    public class PageObjectList<T> : IReadOnlyList<T> where T : PageObject, new()
    {
        internal List<Func<T, bool>> PageObjectFilter = new List<Func<T, bool>>();
        private IReadOnlyList<T> _readOnlyListImplementation;
        private By _bylocator;
        protected Logger Log = LogManager.GetCurrentClassLogger();

        private bool ExecuteFindElements(ISearchContext context, By locator, out ReadOnlyCollection<IWebElement> elements, int timeoutInMs = 5000)
        {
            ReadOnlyCollection<IWebElement> myElem = new ReadOnlyCollection<IWebElement>(new List<IWebElement>());
            elements = myElem;

            Task task = Task.Factory.StartNew(() =>
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
                Log.Debug($"Calling currentSearchContext.FindElements() was OK");
                elements = myElem;
                return true;
            }

            Log.Warn($"Find Elements never completed");
            return false;
        }

        private IReadOnlyList<T> FindMeCandidates_()
        {
            ReadOnlyCollection<IWebElement> meCandidates;

            // Get the elements as normal...
            try
            {
                Log.Debug($"Getting me-candidates from predicate, locator : '{Locator}'");
                var status = ExecuteFindElements(Parent.SearchContext, _bylocator, out meCandidates);
                if (!status)
                {
                    Log.Debug($"Return NOthing found for now, FindElements failed....");
                    return new PageObjectList(
                }
                Log.Debug($"Got {meCandidates.Count} intial candidates");
            }
            catch (StaleElementReferenceException)
            {
                Log.Warn("Parent is stale, clearing Parent cached element");
                ((PageObject)Parent)._cachedElement = null;
                throw;
            }
            catch (Exception e)
            {
                // Forward the error
                Log.Warn(e, $"Predicate : FindmeElements from base call returned an exception, forwarding...");
                ((PageObject)Parent)._cachedElement = null;
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
                        Log.Debug($"Applying the predicate(s)...");
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
                        Log.Debug($"FilteredElements a second time completed...");
                    }
                }
                else
                {
                    Log.Debug($"No predicate used, returning all elements...");
                }

                return filteredElements;
            }
            catch (Exception e)
            {
                Log.Warn($"Applying predicate on elements failed {e.Message}");
                Log.Warn($"Throwing a NoSuchElementException");
                throw new NoSuchElementException();
            }
        }

        public PageObjectList(
            IParent parent,
            LocatorAttribute locatorAttribute)
        {
            Locator = locatorAttribute;
            Parent = parent;
            _bylocator = locatorAttribute.ToWebdriverLocator();
        }

        public IParent Parent { get; set; }

        public LocatorAttribute Locator { get; set; }


        //private ReadOnlyCollection<IWebElement> findElements()
        //{
        //    return this.FindMeCandidates_();
        //}

        //private PageObjectList<T> FilteredElements()
        //{
        //    var webelements = findElements();
        //    var filteredElements = webelements.Select(e => PageObjectFactory.CreatePageObject<T>(Parent, e)).ToList();
        //    polist.PageObjectFilter.Add(predicate);
        //    return polist;
        //}


        //public PageObjectList<T> GetElements(Func<T, bool> predicate)
        //{
        //    var polist = GetElements();
        //    polist.PageObjectFilter.Add(predicate);
        //    return polist;
        //}

        //public PageObjectList<T> GetElements()
        //{
        //    var polist = new PageObjectList<T>();
        //    polist.Parent = this.Parent;
        //    return polist;
        //}

        public int Count => _readOnlyListImplementation.Count;

        public T this[int index] => _readOnlyListImplementation[index];
        public IEnumerator<T> GetEnumerator()
        {
            return _readOnlyListImplementation.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable) _readOnlyListImplementation).GetEnumerator();
        }
    }
}
