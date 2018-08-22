using System;
using System.Collections.Generic;
using Framework.PageObjects;
using Viedoc.viedoc.pages.components.elements;

namespace Viedoc.viedoc.pages.components.fieldItems
{
    /// <summary>
    /// Viedoc base field item class
    /// </summary>
    public abstract class FieldItem : PageObject
    {
       // protected Dictionary<Enum,string> Enum2String = new Dictionary<Enum, string>();
        internal Dictionary<Enum, string> itemDictionary = new Dictionary<Enum, string>();

        [Locator(How.Css, ".field-validation-error")]
        public virtual Label ErrorLabel { get; set; }

        public string ErrorMessage => ErrorLabel.Text;


        [Locator(".info-box,.info-grey")]
        protected virtual Button infoButton { get; set; }

        public virtual T OpenInfoTooltip<T>() where T : PageObject,new()
        {
            infoButton.Click();
            T tooltip = PageObjectFactory.Init<T>(this);
            return tooltip;
        }
    }
}