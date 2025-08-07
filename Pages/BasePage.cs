using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;

public class BasePage
{
    protected IWebDriver Driver;

    public BasePage(IWebDriver driver)
    {
        Driver = driver;
    }

    protected void WaitForElement(By locator, int timeoutInSeconds = 10)
    {
        WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(timeoutInSeconds));
        wait.Until(d => d.FindElement(locator).Displayed);
    }
}