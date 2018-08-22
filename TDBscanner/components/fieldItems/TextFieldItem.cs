using System;
using Framework.PageObjects;
using Viedoc.viedoc.pages.components.elements;

namespace Viedoc.viedoc.pages.components.fieldItems
{

    /// <summary>
    ///     Viedoc Form : Dropdown
    /// </summary>
    [Locator(How.ViedocFieldItem, "select")]
    public class DropdownFieldItem<T,U> : FieldItem where T:Dropdown<U> where U:DropdownItem, new()
    {
        public virtual T Dropdown { get; set; }
       

        public void Select(string label)
        {
            Dropdown.SelectItem(label);
        }

        public string SelectedItem => Dropdown.SelectedItem;

        public void SelectByIndex(int index)
        {
            Dropdown.SelectByIndex(index);
        }

        public void Select(Enum item)
        {
            Dropdown.SelectItem(item);
        }
    }


    [Locator(How.ViedocFieldItem, "textarea")]
    public class TextAreaFieldItem<T> : FieldItem where T : TextAreaField
    {
        protected virtual T Field { get; set; }

        public string TextValue => Field.GetAttribute("value");

        /// <summary>
        /// Set text in field
        /// </summary>
        /// <param name="text"></param>
        /// <param name="append"></param>
        public virtual void SetText(string text, bool append = false)
        {
            Field.SetText(text, append);
        }

        public override void Clear()
        {
            Field.Clear();
        }

        /// <summary>
        /// Check if field is empty
        /// </summary>
        /// <returns></returns>
        public bool IsCleared()
        {
            if (WebDriver.IsIe9())
                return Field.Text == String.Empty;

            return TextValue == string.Empty;
        }
    }

    [Locator(How.ViedocFieldItem, "textarea")]
    public class TextAreaFieldItem : FieldItem
    {
        protected virtual TextAreaField Field { get; set; }

        public string TextValue => Field.GetAttribute("value");

        /// <summary>
        /// Set text in field
        /// </summary>
        /// <param name="text"></param>
        /// <param name="append"></param>
        public virtual void SetText(string text, bool append = false)
        {
            Field.SetText(text, append);
        }

        public override void Clear()
        {
            Field.Clear();
        }

        /// <summary>
        /// Check if field is empty
        /// </summary>
        /// <returns></returns>
        public bool IsCleared()
        {
            if (WebDriver.IsIe9())
                return Field.Text == String.Empty;

            return TextValue == string.Empty;
        }
    }

    /// <summary>
    /// Default Viedoc Textfield element
    /// </summary>
    [Locator(How.ViedocFieldItem, "[type=text]")]
    public class TextFieldItem : FieldItem
    {
        public virtual TextField Field { get; set; }

        [Locator(".has-tip")]
        public virtual Button InfoTooltipButton { get; set; }

        [Locator(".label-box-caption label")]
        public virtual Label LabelBoxCaption { get; set; }

        //[Locator(".error")]
        //public override Label ErrorLabel { get; set; }

        [Locator(".ttl-box")]
        public virtual Label Title { get; set; }

        public  string TextValue => Field.GetAttribute("value");

        /// <summary>
        /// Set text in field
        /// </summary>
        /// <param name="text"></param>
        /// <param name="append"></param>
        public virtual void SetText(string text, bool append = false)
        {
            Field.SetText(text,append);
        }

        public void FieldClick()
        {
            Field.Click();
        }

        public virtual void JavaSetText(string text, bool append = false)
        {
            Field.JavaSetAttribute(text);
            WebDriver.ExecuteJavascript("$(arguments[0]).change();", Field.GetWrappedElement());

        }

        public void SetTextOnChange(string text)
        {
            Field.SetText(text);
            WebDriver.ExecuteJavascript("$(arguments[0]).change();", Field.GetWrappedElement());
        }

        /// <summary>
        /// Set Number in field
        /// </summary>
        /// <param name="number"></param>
        /// <param name="append"></param>
        public virtual void SetText(int number, bool append = false)
        {

            Field.SetText(number.ToString(), append);
        }

        public override void Clear()
        {
            Field.Clear();
        }

        /// <summary>
        /// Check if field is empty
        /// </summary>
        /// <returns></returns>
        public bool IsCleared()
        {
            if (WebDriver.IsIe9())
                return Field.Text == String.Empty;

            return TextValue == string.Empty;
        }

        public override bool Enabled => !IsDisabled;

        public bool IsReadOnly() => Field.GetAttribute("disabled") != null;

        public bool IsDisabled => IsReadOnly();

    }
}