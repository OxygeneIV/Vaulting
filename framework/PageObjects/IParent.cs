using Framework.WebDriver;
using OpenQA.Selenium;

namespace Framework.PageObjects
{
    /// <summary>
    /// Properties needed to exist for a prent (Driver or othe PageObject)
    /// </summary>
    public interface IParent
    {
        ISearchContext SearchContext { get; }
        Driver WebDriver { get; }

        string WindowHandle { get; set; }

    }
}
