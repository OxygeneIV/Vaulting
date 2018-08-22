using System;
using System.Collections.Generic;
using Framework.PageObjects;
using Framework.WaitHelpers;
using OpenQA.Selenium.Support.UI;

namespace Viedoc.viedoc.pages.components.elements
{


    public class SelectOption : PageObject
    {
        
    }

    [Locator("select")]
    public class SelectDropdown : PageObject
    {
        public void SelectByValue(string value)
        {
            Wait.UntilOrThrow(() => Displayed && Enabled);
            var selectElement = new SelectElement(GetWrappedElement());
            selectElement.SelectByValue(value);
        }

        public void SelectByValue(int value)
        {
            SelectByValue($"{value}");
        }

        public void SelectByText(string value)
        {
            Wait.UntilOrThrow(() => Displayed && Enabled);
            var selectElement = new SelectElement(GetWrappedElement());
            selectElement.SelectByText(value);
        }

        public bool OptionExists(string value)
        {
            Wait.UntilOrThrow(() => Displayed && Enabled);
            var selectElement = new SelectElement(GetWrappedElement());
            var options = selectElement.Options;

            var count = options.Count;
            int index = -1;

            for (var i = 0; i < count; i++)
            {
                var webElement = options[i];
                if (!webElement.Text.Contains(value)) continue;
                index = i;
                break;
            }

            if (index >= 0)
            {
                return true;
            }

            return false;
        }

        public void SelectBySubstring(string value)
        {
            Wait.UntilOrThrow(() => Displayed && Enabled);
            var selectElement = new SelectElement(GetWrappedElement());
            var options = selectElement.Options;

            var count = options.Count;
            int index = -1;

            for (var i = 0; i < count; i++)
            {
                var webElement = options[i];
                if (!webElement.Text.Contains(value)) continue;
                index = i;
                break;
            }

            if (index >= 0)
            {
                selectElement.SelectByIndex(index);
            }
            else
            {
                throw new Exception($"String {value} not found in dropdown");
            }
        }
    }




    /// <summary>
    /// Default Viedoc Dropdown item, default locating by text
    /// </summary>
    [Locator(How.Css, ".active-result")]
    public class DropdownItem : PageObject
    {
        public override Func<string, LocatorAttribute> LocatorFunc { get; set; } = s => new LocatorAttribute(How.XPath,
            $".//*[contains(@class,'active-result') and text()='{s}']");
    }

    /// <summary>
    /// Default Viedoc Dropdown element
    /// </summary>
    [Locator(".chosen-container",false)]
    public class Dropdown : Dropdown<DropdownItem>
    {
        
    }

    /// <summary>
    /// Generic Dropdown class
    /// </summary>
    /// <typeparam name="T"></typeparam>
    [Locator(".chosen-container", false)]
    public class Dropdown<T> : PageObject where T : DropdownItem, new()
    {
        // Keep track of available items if predefined
        protected internal Dictionary<Enum, string> itemDictionary = new Dictionary<Enum, string>();

        public virtual IEnumerable<T> Items { get; set; }

        [Locator(".chosen-single, .chosen-choices, .select2-selection--single")] protected Link _chosenSingle;

        [Locator(".chosen-single div b")] protected Link _chosenSingleAndroid;

        protected virtual bool IsExpanded()
        {
            Log.Debug("Starting IsExpanded");
            var ret = HasClassName("chosen-with-drop") || _chosenSingle.GetAttribute("aria-expanded") == "true";

            Log.Debug($"IsExpanded , returning => {ret}");
            return ret;
        }


        public InputField SearchField;




        // Trigger dropdown expansion
        protected virtual Action Trigger { get; set; }

        public void ShowOptions()
        {
            Trigger();
        }

        public Dropdown()
        {
                Trigger = () =>
                {
                    Log.Debug("Trigger started");

                    if (WebDriver.IsMobile() && !WebDriver.IsAndroid())
                    {
                        _chosenSingle.Tap();
                      
                    }
                    else
                    {

                        _chosenSingle.Click();
                    }
                    Log.Debug("Trigger ended");
                };
        }

    
        /// <summary>
        /// Return the current selection (trimmed) as text
        /// </summary>
        public string SelectedItem 
        {
            get
            {
                var strItem = _chosenSingle.Text.Trim();
                return strItem;
            }
        }

        public virtual void SelectItem(Enum item)
        {
            string strItem = itemDictionary[item];
            SelectItem(strItem);
        }

        public virtual void SelectItemSkipDropdown(Enum item)
        {
            string strItem = itemDictionary[item];
            SelectItemSkipDropdown(strItem);
        }

        public virtual void SelectByIndex(int index)
        {
            if (!IsExpanded())
            {
                Log.Debug("Selection not expanded yet");
                Trigger();
                Wait.UntilOrThrow(IsExpanded,message:"Wait for expanded Dropdown");
            }
            else
            {
                Log.Debug("Selection already expanded");
            }

            //System.Threading.Thread.Sleep(1000);
            // Todo use GetElements() if locators are overwritten!!

            Items.GetElement(index).Click();
            if (WebDriver.IsMobile())
            {
                WebDriver.HideKeyboard();// ("Done","HideKeyboardStrategy");
            }
        }

        public virtual bool ItemExists(string item)
        {
            Log.Info($"Checking if item {item} exists");
            if (!IsExpanded())
            {
                Trigger();
                Wait.UntilOrThrow(IsExpanded);
            }

           // System.Threading.Thread.Sleep(1000);


            var theItem = Items.GetElement(item);          
            bool ans = theItem.Exists();
            Trigger();
            Wait.UntilOrThrow(()=>!IsExpanded());
            Log.Info($"Checking if item {item} exists returned {ans}");
            return ans;
        }



        public virtual void SelectItem(string item)
        {
            if(!IsExpanded())
            {
                Log.Debug("Selection not expanded yet");
                Trigger();
                if (WebDriver.IsAndroid())
                {
                    System.Threading.Thread.Sleep(1000);
                    WebDriver.HideKeyboard();
                }
                Wait.UntilOrThrow(IsExpanded, message: "Wait for expanded Dropdown");
            }

            //System.Threading.Thread.Sleep(1000);


            var theItem = Items.GetElement(item);

            theItem.ScrollIntoView();
            theItem.Click();

            if (WebDriver.IsMobile())
            {
                WebDriver.HideKeyboard();
            }
        }

        public virtual void SelectItemSkipDropdown(string item)
        {
            if (!IsExpanded())
            {
                Log.Debug("Selection not expanded yet");
                var elem = GetWrappedElement();

                var str = "$(arguments[0]).addClass(\"chosen-container-active chosen-with-drop\")";

                WebDriver.ExecuteJavascript(str,elem);

                if (WebDriver.IsAndroid())
                {
                    System.Threading.Thread.Sleep(1000);

                }
                Wait.UntilOrThrow(IsExpanded, message: "Wait for expanded Dropdown");
            }

            //System.Threading.Thread.Sleep(1000);

            if (SearchField.Displayed)
            {
                SearchField.SendKeys("U"+OpenQA.Selenium.Keys.Tab);
            }
            else
            {
                var theItem = Items.GetElement(item);
                Wait.UntilOrThrow(() => theItem.Displayed);

                //theItem.ScrollIntoView();
                theItem.JavaClick();
            }

            WebDriver.HideKeyboard();

        }

    }
}
