using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;


namespace selenium.Pages
{
    public class AdminPage:BasePage
    {
        string errorMess;
        float ketQua;


        public AdminPage(IWebDriver driver) : base(driver) { }
    }
}
