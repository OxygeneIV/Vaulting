using System;
using System.Collections.Generic;
using OpenQA.Selenium;
using SizSelCsZzz;


namespace Framework.PageObjects
{
    /// <summary>
    ///   Provides the lookup methods for the FindsBy attribute (for using in PageObjects)
    /// </summary>
    public enum How
    {
        /// <summary>
        ///   Finds by <see cref="By.Id" />
        /// </summary>
        Id,

        /// <summary>
        ///   Finds by <see cref="By.Name" />
        /// </summary>
        Name,

        /// <summary>
        ///   Finds by <see cref="By.TagName" />
        /// </summary>
        Tag,

        /// <summary>
        ///   Finds by <see cref="By.ClassName" />
        /// </summary>
        Class,

        /// <summary>
        ///   Finds by <see cref="By.CssSelector" />
        /// </summary>
        Css,

        /// <summary>
        ///   Finds by <see cref="By.LinkText" />
        /// </summary>
        LinkText,

        /// <summary>
        ///   Finds by <see cref="By.PartialLinkText" />
        /// </summary>
        PartialLinkText,

        /// <summary>
        ///   Finds by <see cref="By.XPath" />
        /// </summary>
        XPath,

        /// <summary>
        ///   Finds by Sizzle expression or jquery expression
        /// </summary>
        Sizzle,

        /// <summary>
        ///   Finds by Sizzle expression or jquery expression
        /// </summary>
        JQuery,

        /// <summary>
        ///   Finds by matching string in data-bind attribute
        /// </summary>
        DataBind,

        /// <summary>
        ///   Find the Viedoc page specific locator
        /// </summary>
        ViedocPage,

        /// <summary>
        ///   Find the data-attribute element
        /// </summary>
        ViedocTest,

        /// <summary>
        ///   Find the form-item with specific ID
        /// </summary>
        ViedocFormItem,


        ViedocFieldItem,

        Href,
        /// <summary>
        ///   Finds by Formatter expression
        /// </summary>
        Formatter,
    }

    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, Inherited = true)]
    public class ItemLocatorAttribute : Attribute
    {
        /// <summary>
        ///   Gets or sets the how.
        /// </summary>
        public How How { get; set; }

        /// <summary>
        ///   Gets or sets a value indicating whether use cache.
        /// </summary>
        public bool UseCache { get; set; }

        /// <summary>
        ///   Gets or sets the using.
        /// </summary>
        public string Formatter { get; set; }

        public override string ToString()
        {
            return $"How : {How} , Formatter : {Formatter}";
        }

        public ItemLocatorAttribute(How how, string @formatter, bool useCache = true)
        {
            How = how;
            Formatter = @formatter;
            UseCache = useCache;
        }
    }


    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, Inherited = false,AllowMultiple = true)]
    public class DictionaryItemAttribute : Attribute
    {
        /// <summary>
        ///   Gets or sets the how.
        /// </summary>
        public Enum ItemKey { get; set; }

        public string Value { get; set; }

        public DictionaryItemAttribute(object key, string @value)
        {
            ItemKey = (Enum)key;
            Value = @value;
        }
    }


    /// <summary>
    ///   ElementAttribute , define properties of a web element
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Field | AttributeTargets.Property, Inherited = true)]
    public class LocatorAttribute : Attribute
    {
        public static LocatorAttribute Generate(How how, string Using)
        {
            return new LocatorAttribute(how, Using);
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="LocatorAttribute"/> class.
        /// </summary>
        /// <param name="how">
        /// The how.
        /// </param>
        /// <param name="using">
        /// The using.
        /// </param>
        /// <param name="useCache">
        /// The use cache.
        /// </param>
        public LocatorAttribute(How how, string @using, bool useCache = true)
        {
            How = how;
            Using = @using;
            UseCache = useCache;
        }

        public LocatorAttribute(How how1,string @using1, How how2,string @using2,bool useCache = true)
        {
            var loc1 = Generate(how1, using1);
            var loc2 = Generate(how2, using2);

            //How = how;
            //Using = @using;
            UseCache = useCache;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="LocatorAttribute"/> class.
        /// </summary>
        /// <param name="using">
        /// The using.
        /// </param>
        /// <param name="useCache">
        /// The use cache.
        /// </param>
        public LocatorAttribute(string @using, bool useCache = true)
        {
            How = How.Css;
            Using = @using;
            UseCache = useCache;
        }

        /// <summary>
        ///   Gets or sets the how.
        /// </summary>
        public How How { get; set; }

        /// <summary>
        ///   Gets or sets a value indicating whether use cache.
        /// </summary>
        public bool UseCache { get; set; }

        /// <summary>
        ///   Gets or sets the using.
        /// </summary>
        public string Using { get; set; }

        public List<LocatorAttribute> LocatorAttributes { get; set; }

        public override string ToString()
        {
            return $"How : {How} , Using : {Using}";
        }

        /// <summary>
        ///   The to locator.
        /// </summary>
        /// <returns>
        ///   The <see cref="By" />.
        /// </returns>
        /// <exception cref="Exception">
        /// </exception>
        public By ToWebdriverLocator()
        {
            switch (How)
            {
                case How.Id:
                    return By.Id(Using);
                case How.Name:
                    return By.Name(Using);
                case How.Tag:
                    return By.TagName(Using);
                case How.Class:
                    return By.ClassName(Using);
                case How.Css:
                    return By.CssSelector(Using);
                case How.LinkText:
                    return By.LinkText(Using);
                case How.PartialLinkText:
                    return By.PartialLinkText(Using);
                case How.XPath:
                    return By.XPath(Using);
                case How.Sizzle:
                    return BySizzle.CssSelector(Using);
                case How.DataBind:
                    return By.CssSelector($"[data-bind*='{Using}']");
                case How.ViedocFormItem:
                    return BySizzle.CssSelector($".form-item:has('#{Using}')");
                case How.ViedocFieldItem:
                    return BySizzle.CssSelector($".field-item:has('{Using}')");
                case How.ViedocTest:
                    return By.CssSelector($"[data-vidoctest='{Using}']");
                case How.JQuery:
                    return ByJQuery.CssSelector(Using);
                case How.ViedocPage:
                    return By.CssSelector($"body.{Using}");
                case How.Href:
                    return By.CssSelector($"[href*='{Using}']");
            }
            throw new Exception("Unknown Locator method " + How);
        }
    }
}