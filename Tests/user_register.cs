using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OpenQA.Selenium;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using SeleniumExtras.WaitHelpers;

namespace user_auth
{
    public class user_register
    {
        private IWebDriver driver;
        Application excelApp;
        Workbook workbook;
        Worksheet worksheet;

        private IWebElement ipnEmail, ipnPassword, ipnFullName, ipnPhoneNumber,ipnAddress,ipnTerm;

        private static readonly string ExcelFilePath = @"D:\đảm bảo chất lượng phần mềm\selenium\TestData\TestCase.xlsx";

        [SetUp]
        public void Setup()
        {
            driver = new OpenQA.Selenium.Chrome.ChromeDriver();
            driver.Url = "http://localhost/a/McDelivery/mcRegister.php";//Đi vào trang kí
            driver.Navigate();
            Thread.Sleep(1000);

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(ExcelFilePath, ReadOnly: false, Editable: true);
            ExcelPage(6);
            Thread.Sleep(1000);
        }

        [Test]
        [TestCase(12)]
        [TestCase(13)]
        [TestCase(14)]
        [TestCase(15)]
        [TestCase(16)]
        [TestCase(17)]
        [TestCase(18)]
        [TestCase(19)]
        [TestCase(20)]
        [TestCase(21)]
        [TestCase(22)]
        [TestCase(23)]
        [TestCase(24)]
        public void TestRegister(int row)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("email")));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("password")));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("name")));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("phone")));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("address")));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("email")));
            wait.Until(ExpectedConditions.ElementIsVisible(By.TagName("button")));

            ipnEmail = driver.FindElement(By.Id("email"));
            ipnPassword = driver.FindElement(By.Id("password"));
            ipnFullName = driver.FindElement(By.Id("name"));
            ipnPhoneNumber = driver.FindElement(By.Id("phone"));
            ipnAddress = driver.FindElement(By.Id("address"));
            ipnTerm = driver.FindElement(By.Id("terms"));

            // 5 trường trong excel
            var email = handleDataTest(worksheet.Cells[row, 7].Value);
            var pwd = handleDataTest(worksheet.Cells[row, 8].Value);
            var name = handleDataTest(worksheet.Cells[row, 9].Value);
            var phonenumber = handleDataTest(worksheet.Cells[row, 10].Value);
            var address = handleDataTest(worksheet.Cells[row, 11].Value);
            var testId = worksheet.Cells[row, 1].Value;
            var testDes = worksheet.Cells[row, 3].Value;

            Console.WriteLine($"Đây là test case: {testId}");
            Console.WriteLine($"Trường hợp test: {testDes}");


            Console.WriteLine($"Test data đây | Email: {email} | Pwd: {pwd} " +
         $"| Họ tên: {name} | SDT: {phonenumber} | Địa chỉ : {address} ");

         
            //sendkey
            SetKeysForInput(email, pwd, name, phonenumber, address);

            if (row == 17)
            {
                ipnTerm.Click();
                isRequiredAtribute(ipnTerm,row);
            }

            //Kiểm tra khi nhập dữ liệu trống vào trường thì có được 
            checkPwdMin(pwd, ipnPassword, row);
            

            try
            {          
                //Bấm nút đăng kí
                IWebElement regbtn = driver.FindElement(By.TagName("button"));
                if (row == 13 )
                {
                    InputResult(row, "Tạo tài khoản thành công","FAILED");
                }
                regbtn.Click();
            }
            catch(UnhandledAlertException e)
            {
                checkIfRegisterIfExistEmail(row);
            }
            checkIfRegisterSucess(row);
            checkInvalidData(email, pwd, name, phonenumber, address,row);
            checkIfEmailEqualPwd(row);
            checkRegisterWithInvalidPhone(row);
            checkInvalidEmail(email, row);
            checkInvalidPwd(row);
        }
       


        public void SetKeysForInput(string email, string pwd, string name, string phone, string address )
        {
            try
            {
                //Clear nội dung bên trong
                ipnEmail.Clear();
                ipnPassword.Clear();
                ipnFullName.Clear();
                ipnPhoneNumber.Clear();
                ipnAddress.Clear();

                Thread.Sleep(2000);
                ipnEmail.SendKeys(email);
                ipnPassword.SendKeys(pwd);
                ipnFullName.SendKeys(name);
                ipnPhoneNumber.SendKeys(phone);
                ipnAddress.SendKeys(address);
                ipnTerm.Click();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            Thread.Sleep(4000);
        }

        public void checkIfRegisterSucess(int row)
        {
            if (row == 12)
            {
                if (driver.Url == "http://localhost/a/McDelivery/mcDeliver.php")
                {
                    InputResult(row, "Đăng kí thành công và đã chuyển đến trang đăng nhập", "PASSED");
                }
                else
                {
                    InputResult(row, $"Chuyển đến trang {driver.Url}", "FAILED");
                }
            }
        }

        public void checkInvalidPwd(int row)
        {
            if (row == 23)
            {
                if (driver.Url == "http://localhost/a/McDelivery/mcDeliver.php")
                {
                    InputResult(row, $"Hệ thống cho phép tạo tài khoản ", "FAILED");
                }
                else
                {
                    InputResult(row, $"Mật khẩu không hợp lệ ", "PASSED");

                }
            }
        }

        public void checkIfEmailEqualPwd(int row)
        {
            if(row == 20)
            {
                if (driver.Url == "http://localhost/a/McDelivery/mcDeliver.php")
                {
                    InputResult(row, $"Hệ thống cho phép tạo tài khoảng ", "FAILED");
                }
                else
                {
                    InputResult(row, $"Mật khẩu không hợp lệ ", "PASSED");

                }
            }
        }

        public void checkIfRegisterIfExistEmail(int row)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            IAlert alert= wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.AlertIsPresent());

            if (alert != null)
            {
                string alertText = alert.Text;
                InputResult(row, $"Hệ thống thông báo {alertText}", "PASSED");
            }
            else {
                InputResult(row, "Hệ thống vẫn cho phép tạo với email trùng", "FAILED");
            }
        }

        public void checkInvalidData(
          string strEmail, string strPwd, string strName, string strPhone,string strAdress, int row)
        {
            if (strEmail == "")
            {
                isRequiredAtribute(ipnEmail, row);
            }
            else if (strPwd == "")
            {
                isRequiredAtribute(ipnPassword, row);

            }
            else if (strName == "")
            {
                isRequiredAtribute(ipnFullName, row);

            }
            else if (strPhone == "")
            {
                isRequiredAtribute(ipnPhoneNumber, row);
            }
            else if(strAdress == "")
            {
                isRequiredAtribute(ipnAddress, row);
            }
        }

        public void checkPwdMin(string pwd, IWebElement ipnPassword, int row)
        {
            var limited = ipnPassword.GetAttribute("min");

            if (row == 16)
            {
                if (limited != null)
                {
                    int min;
                    if (int.TryParse(limited, out min))
                    {
                        if (pwd.Length <= min)
                        {
                            InputResult(row, "Mật khẩu phải có ít nhất 6 kí tự", "PASSED");
                        }
                        else
                        {
                            InputResult(row, "Hệ thống cho phép nhập hơn 6 kí tự", "FAILED");
                        }
                    }
                    else
                    {
                        InputResult(row, "Mật khẩu phải có ít nhất 6 kí tự", "PASSED");
                    }
                }
                else
                {
                    InputResult(row, "Hệ thống không giới hạn kí tự nhập trường password", "FAILED");
                }
            }
        }
        public void checkAdressMax(string adress, IWebElement ipnAdress, int row)
        {
            var limited = ipnPassword.GetAttribute("max");

            if (row == 24)
            {
                if (limited != null)
                {
                    int max;
                    if (int.TryParse(limited, out max))
                    {
                        if (adress.Length <= max)
                        {
                            InputResult(row, $"Địa chỉ có số kí tự tối đa là {max}", "PASSED");
                        }
                        else
                        {
                            InputResult(row, "Hệ thống không ràng buộc ", "FAILED");
                        }
                    }
                    else
                    {
                        InputResult(row, $"Địa chỉ có số kí tự tối đa là {max}", "PASSED");
                    }
                }
                else
                {
                    InputResult(row, "Hệ thống không giới hạn kí tự nhập trường ", "FAILED");
                }
            }
        }

        public void checkInvalidEmail(string email,int row)
        {
            if (row == 15 )
            {
                if (driver.Url == "http://localhost/a/McDelivery/mcDeliver.php")
                {
                    InputResult(row, $"Tạo tài khoản thành công ", "FAILED");
                }
                else
                {
                    
                    InputResult(row, $"Email không đúng định dạng", "PASSED");
                }
            }
        }

        public void checkRegisterWithInvalidPhone(int row)
        {
            if(row == 18 || row == 22)
            {
                if (driver.Url == "http://localhost/a/McDelivery/mcDeliver.php")
                {
                    InputResult(row, $"Tạo tài khoản thành công ", "FAILED");
                }
                else
                {

                    InputResult(row, $"Số điện thoại không đúng định dạng", "PASSED");
                }
            }
        }

        public void isRequiredAtribute(IWebElement e, int row)
        {
            string requiredAttribute = e.GetAttribute("required");
            if (requiredAttribute == "false")
            {
                InputResult(row, requiredAttribute, "FAILED");
            }
            else
            {
              
                InputResult(row, "Vui lòng nhập vào trường", "PASSED");
            }
        }

        public void ExcelPage(int page)
        {
            worksheet = workbook.Sheets[page];
        }

        public string handleDataTest(string data)
        {
            switch (data)
            {
                case "N/A":
                    return "";
            }
            return data;
        }

        public void InputResult(int row, string result, string status)
        {
            worksheet.Cells[row, 13].Value = result; // Ghi kết quả vào cột (Result)
            worksheet.Cells[row, 14].Value = status; // Ghi kết quả passed or failed
        }


        public void ClickElementIfInteractable(By locator)
        {
            try
            {
                // Đợi phần tử có thể click được (với thời gian tối đa là 10 giây)
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

                // Đợi phần tử trở nên hiển thị và có thể click được
                var element = wait.Until(ExpectedConditions.ElementToBeClickable(locator));

                // Kiểm tra nếu phần tử hiển thị và có thể click được
                if (element.Displayed && element.Enabled)
                {
                    element.Click();
                    Console.WriteLine("Element clicked successfully.");
                }
                else
                {
                    Console.WriteLine("Element is not interactable.");
                }
            }
            catch (NoSuchElementException)
            {
                // Xử lý nếu phần tử không tồn tại
                Console.WriteLine("The element was not found.");
            }
            catch (ElementNotInteractableException)
            {
                // Xử lý nếu phần tử không thể tương tác được
                Console.WriteLine("The element is not interactable.");
            }
            catch (WebDriverTimeoutException)
            {
                // Xử lý nếu thời gian chờ hết mà phần tử không xuất hiện
                Console.WriteLine("Timed out waiting for the element to be clickable.");
            }
            catch (Exception ex)
            {
                // Xử lý các ngoại lệ không xác định
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        [TearDown]
        public void ClearUp()
        {
            try
            {
                workbook.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }
            finally
            {
                workbook.Close();
                excelApp.Quit();
                driver.Quit();

                // Release COM objects
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
