using NUnit.Framework;
using OpenQA.Selenium;

[TestFixture]
public class TestBase
{
    protected IWebDriver Driver;

    [SetUp]
    public void Setup()
    {
        Driver = DriverFactory.CreateDriver(AppConfig.Browser);
        Driver.Manage().Window.Maximize();
    }

    [TearDown]
    public void TearDown()
    {
        Driver?.Quit();
    }
}