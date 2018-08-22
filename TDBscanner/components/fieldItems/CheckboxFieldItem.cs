using Framework.PageObjects;
using Viedoc.viedoc.pages.forms.FormItems;

namespace Viedoc.viedoc.pages.components.fieldItems
{
    /// <summary>
    /// Default Viedoc Textfield element
    /// </summary>
    [Locator(How.ViedocFieldItem, "[type=checkbox]")]
    public class CheckboxFieldItem : FieldItem
    {
        protected  CheckBoxItem Field { get; set; }

        public virtual void Check()
        {
            Field.Check();
        }
        public virtual void UnCheck()
        {
            Field.UnCheck();
        }

        public virtual bool IsChecked()
        {
            return Field.IsChecked();
        }

        public virtual bool IsReadOnly => HasClassName("hasReadOnly");

        public bool IsDisabled => IsReadOnly;

        public override bool Enabled => !IsReadOnly;

    }
}
