using Framework.PageObjects;
using Viedoc.viedoc.pages.components.elements;

namespace Viedoc.viedoc.pages.components.fieldItems
{
    [Locator(How.ViedocFieldItem,"[type=password]")]
    public class PasswordFieldItem : TextFieldItem
    {
       // [Locator(How.Sizzle,"input[name=Password],input[type=password]:not([style*='display: none'])")]
        [Locator(How.Sizzle, "input[name=Password],input[name=ConfirmPassword]")]
        public override TextField Field { get; set; }

    }
}
