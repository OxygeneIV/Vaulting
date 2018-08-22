namespace Framework.PageObjects
{
    /// <summary>
    /// Simple Locator for FormItems (fetching by Id by default)
    /// </summary>
    public class Id : LocatorAttribute
    {
        public Id(string @using, bool useCache = true) : base(@using, useCache)
        {
            How = How.ViedocFormItem;
        }

        public Id(How how, string @using, bool useCache = true) : base(how, @using, useCache)
        {
        }
    }
}