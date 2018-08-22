using System;

namespace Tests.Base
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method , Inherited = true)]
    public class BrowserAttribute : Attribute
    {
        public bool Initiate { get; set; }

        public BrowserAttribute(bool initiateBrowser)
        {
            Initiate = initiateBrowser;
        }

        public BrowserAttribute()
        {
        }
    }
}