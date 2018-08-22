using Framework.PageObjects;
using Viedoc.viedoc.pages.components.elements;

namespace Viedoc.viedoc.pages.forms.FormItems
{
    /// <summary>
    /// Viedoc Form : CheckBoxItem
    /// </summary>
    [Locator(".check-item")]
    public class CheckBoxItem : PageObject
    {
        [Locator(".checkboxArea,.checkboxAreaChecked")]
        protected virtual Label checkboxArea { get; set; }

        public bool IsChecked()
        {
            return checkboxArea.HasClassName("checkboxAreaChecked");
        }


        [Locator("label")]
        protected Label _label;

        public string Label => _label.Text;



        public virtual void Check()
        {
            if (!IsChecked())
            {
                checkboxArea.Click();
            }
        }
        public virtual void UnCheck()
        {
            if (IsChecked())
            {
                checkboxArea.Click();
            }
        }
    }
}