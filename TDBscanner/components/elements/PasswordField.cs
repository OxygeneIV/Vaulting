using Framework.PageObjects;

namespace Viedoc.viedoc.pages.components.elements
{
    /// <summary>
    /// Default Viedoc Password field element
    /// </summary>
    [Locator(How.Sizzle, "input[type='password']")]
    public class PasswordField : TextField
    {
    }
}