using Framework.PageObjects;
using Viedoc.viedoc.pages.components.elements;

namespace Viedoc.viedoc.pages.components.fieldItems
{
    [Locator(How.ViedocFieldItem, "[type=email]")]
    public class EmailFieldItem : TextFieldItem
    {
        [Locator(How.Sizzle, "input[type='email']")]
        public override TextField Field { get; set; }
    }
}
