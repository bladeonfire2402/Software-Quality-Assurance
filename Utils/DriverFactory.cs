using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;

public static class DriverFactory
{
    public static IWebDriver CreateDriver(string browser)
    {
        switch (browser.ToLower())
        {
            case "chrome":
                return new ChromeDriver();
            // Thêm các trình duyệt khác nếu cần (Firefox, Edge, v.v.)
            default:
                throw new ArgumentException("Browser không được hỗ trợ!");
        }
    }
}