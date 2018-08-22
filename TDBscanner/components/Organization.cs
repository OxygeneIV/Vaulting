using System;
using System.Collections.Generic;
using Framework.PageObjects;
using Viedoc.viedoc.pages.components.elements;

namespace Viedoc.viedoc.pages.components
{
    [Locator("section.panel")]
    public class Organization : PageObject
    {

        [Locator("a")]
        protected class StudyLink : PageObject
        {
            public override Func<string, LocatorAttribute> LocatorFunc { get; set; } =
                s => new LocatorAttribute(How.Sizzle, $"a[title='{s}']");
        }

        [Locator(How.Sizzle, "header :header")]
        protected Label _name;

        [Locator(How.Sizzle, "[href*='SelectOrganization']:not([title])")]
        protected Button BtnShowStudies;



        public string Name => _name.Text;

        protected IEnumerable<StudyLink> Studies;
    }
}
