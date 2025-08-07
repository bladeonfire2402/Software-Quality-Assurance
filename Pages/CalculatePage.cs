using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;

public class CalculatePage : BasePage
{
    string errorMess;
    float ketQua;
    // Định nghĩa các locator cho các phần tử trên trang đăng nhập
    private By SoA => By.Id("number1Field");
    private By SoB => By.Id("number2Field");
    private By DropdownButton => By.Id("selectOperationDropdown");
    private By CalculateButton => By.Id("calculateButton");
    private By KetQua => By.Id("numberAnswerField");
    private By IntergerButton => By.Id("integerSelect");
    private By ClearButton => By.Id("clearButton");
    private By ErrorMessage => By.Id("errorMsgField");

    public CalculatePage(IWebDriver driver) : base(driver) { }

    // Các phương thức tương tác với trang
    public void EnterSoA(string soA)
    {
        Driver.FindElement(SoA).SendKeys(soA);
    }
    public void EnterSoB(string soB)
    {
        Driver.FindElement(SoB).SendKeys(soB);
    }
    public void ClickAdd()
    {
        IWebElement dropdownOption = Driver.FindElement(By.XPath("//select[@id='selectOperationDropdown']/option[@value='0']"));
        dropdownOption.Click();
    }
    public void ClickSubtract()
    {
        IWebElement dropdownOption = Driver.FindElement(By.XPath("//select[@id='selectOperationDropdown']/option[@value='1']"));
        dropdownOption.Click();
    }
    public void ClickMultiply()
    {
        IWebElement dropdownOption = Driver.FindElement(By.XPath("//select[@id='selectOperationDropdown']/option[@value='2']"));
        dropdownOption.Click();
    }
    public void ClickDivide()
    {
        IWebElement dropdownOption = Driver.FindElement(By.XPath("//select[@id='selectOperationDropdown']/option[@value='3']"));
        dropdownOption.Click();
    }
    public float GetAnswer()
    {
        return ketQua = float.Parse(Driver.FindElement(KetQua).GetAttribute("value"));
    }
    public void Calculate() 
    {
        IWebElement kq = Driver.FindElement(CalculateButton);
        kq.Click();
    }
    public string getErrorMessage() 
    {
        IWebElement ErrorMessage = Driver.FindElement(By.Id("errorMsgField"));
        if (ErrorMessage.Displayed)
        {
            return errorMess = ErrorMessage.Text;
        }
        return errorMess;
    }
    public void CalculateAdd(string soA, string soB)
    {
        EnterSoA(soA);
        EnterSoB(soB);
        ClickAdd();
        Calculate();
        GetAnswer();
        getErrorMessage();
    }
    public void CalculateSubtract(string soA, string soB)
    {
        EnterSoA(soA);
        EnterSoB(soB);
        ClickSubtract();
        Calculate();
        GetAnswer();
        getErrorMessage();
    }

    public void CalculateMultiphy(string soA, string soB)
    {
        EnterSoA(soA);
        EnterSoB(soB);
        ClickMultiply();
        Calculate();
        GetAnswer();
        getErrorMessage();
    }

    public void CalculateDivide(string soA, string soB)
    {
        EnterSoA(soA);
        EnterSoB(soB);
        ClickDivide();
        Calculate();
        getErrorMessage();
        if (errorMess == null) 
        {
            GetAnswer();
        }
        
    }
}