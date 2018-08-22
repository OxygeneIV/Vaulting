using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Service;
using OpenQA.Selenium.Remote;

namespace Framework.WebDriver
{
    /// <inheritdoc />
    public class AppiumDriver : AppiumDriver<AppiumWebElement>
    {
        public AppiumDriver(ICommandExecutor commandExecutor, ICapabilities desiredCapabilities) : base(commandExecutor, desiredCapabilities)
        {
        }

        public AppiumDriver(ICapabilities desiredCapabilities) : base(desiredCapabilities)
        {
        }

        public AppiumDriver(ICapabilities desiredCapabilities, TimeSpan commandTimeout) : base(desiredCapabilities, commandTimeout)
        {
        }

        public AppiumDriver(AppiumServiceBuilder builder, ICapabilities desiredCapabilities) : base(builder, desiredCapabilities)
        {
        }

        public AppiumDriver(AppiumServiceBuilder builder, ICapabilities desiredCapabilities, TimeSpan commandTimeout) : base(builder, desiredCapabilities, commandTimeout)
        {
        }

        public AppiumDriver(Uri remoteAddress, ICapabilities desiredCapabilities) : base(remoteAddress, desiredCapabilities)
        {
        }

        public AppiumDriver(AppiumLocalService service, ICapabilities desiredCapabilities) : base(service, desiredCapabilities)
        {
        }

        public AppiumDriver(Uri remoteAddress, ICapabilities desiredCapabilities, TimeSpan commandTimeout) : base(remoteAddress, desiredCapabilities, commandTimeout)
        {
        }

        public AppiumDriver(AppiumLocalService service, ICapabilities desiredCapabilities, TimeSpan commandTimeout) : base(service, desiredCapabilities, commandTimeout)
        {
        }
    }
}