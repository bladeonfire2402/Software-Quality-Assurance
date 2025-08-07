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
using System.Threading.Tasks;

namespace Intergrade
{
    class IntegradeDiscount
    {
        private IWebDriver driverForUser;
        private IWebDriver driverForAdmin;

        Application excelApp;
        Workbook workbook;
        Worksheet worksheet;

        private IWebElement ipnAddress, ipnPhone, ipnEmail, ipnDiscount;
        private IWebElement mailIpn, pwdIpn;

        private string deletedCoupon;//biến để lưu lại mã coupon đã bị xóa
        private string email;//Biến lưu email người dùng

        private static readonly string ExcelFilePath = @"D:\đảm bảo chất lượng phần mềm\selenium\TestData\TestCase.xlsx";

        [SetUp]
        public void SetupAdmin()
        {
            driverForAdmin = new OpenQA.Selenium.Chrome.ChromeDriver();
            driverForAdmin.Url = "http://localhost/a/admin/admin%20dashboard/listmagiamgia.php";
            driverForAdmin.Navigate();
            Thread.Sleep(1000);

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(ExcelFilePath, ReadOnly: false, Editable: true);
            ExcelPage(5);
            Thread.Sleep(1000);
        }

        [SetUp]
        public void SetupUser()
        {
            driverForUser = new OpenQA.Selenium.Chrome.ChromeDriver();
            driverForUser.Url = "http://localhost/a/McDelivery/indexmcDeliver.php";
            driverForUser.Navigate();
            Thread.Sleep(2000);
        }

       

        [Test]
        [TestCase(25)]
        public void IntergrateDiscount(int row)
        {
            //Xóa mã giảm giá bên Admin
            handleRemoveAdmin(row);

            //Kiểm tra xem người dùng còn có thể sử dụng được mã không
            handleCheckUseable(row);
        }

        //Kiểm tra trường hợp mã giảm giá được sử dụng cùng 1 lúc
        [Test]
        public void Intergrate_TestUsersApplyDiscountCodeAtSameTime()
        {
            UserUseCodeAtSameTime();
        }

        static async Task UserUseCodeAtSameTime() {
           
        }

        

        public void loginUser() 
        {

        }

        public void useCode()
        {

        }

        

        //Thực thi xóa mã giảm giá
        public void handleRemoveAdmin(int row)
        {
            var strDiscount = handleDataTest(worksheet.Cells[row, 7].Value);
            var dateRelease = handleDataTest(worksheet.Cells[row, 8].Value);
            var dateExpired = handleDataTest(worksheet.Cells[row, 9].Value);
            var strDescription = handleDataTest(worksheet.Cells[row, 10].Value);

            //Gán để lưu lại mã đã bị xóa
            deletedCoupon = strDiscount;

            Console.WriteLine($"Test data đây | Mã giảm giá: {strDiscount} | Ngày phát hành: {dateRelease} " +
    $"| Ngày hết hạn: {dateExpired} | Mô tả mã giảm giá: {strDescription}");

            try
            {
                IList<IWebElement> trElements = driverForAdmin.FindElements(By.TagName("tr"));

                var deleteItem = driverForAdmin.FindElement(By.TagName("tr"));

                foreach (var tr in trElements)
                {
                    IList<IWebElement> tdElements = tr.FindElements(By.TagName("td"));

                    foreach (var td in tdElements)
                    {
                        if(td.Text == strDiscount)
                        {
                            deleteItem = tr;      
                        }
                    }
                }

                IReadOnlyCollection<IWebElement> links = deleteItem.FindElements(By.TagName("a"));

                IWebElement btnDelete = links.ElementAt(1);

                btnDelete.Click();

                IAlert alert = driverForAdmin.SwitchTo().Alert();

                alert.Accept();

                Thread.Sleep(5000);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Lỗi nè {e.Message}");
            }
        }

        //Kiểm tra mã giảm giá
        public void handleCheckUseable(int row)
        {
            Login();
            addToCheckOut();
            handleCheckOutForm(row);
        }

        public void Login()
        {
            //Lấy dữ liệu để đăng nhập bên trang 6
            ExcelPage(7);
            int row = 12;
            WebDriverWait wait = new WebDriverWait(driverForUser, TimeSpan.FromSeconds(8));

            IList<IWebElement> input = driverForUser.FindElements(By.TagName("input"));

            mailIpn = input[0];
            pwdIpn = input[1];

            var email =worksheet.Cells[row, 7].Value;
            string pwdStr = worksheet.Cells[row, 8].Value.ToString();

            var pwd = pwdStr;

            Console.WriteLine($"Dữ liệu đăng nhập đây khách hàng: {email} và {pwd}");

            mailIpn.SendKeys(email);
            pwdIpn.SendKeys(pwd);

            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("SignIn-btn")));

            IWebElement btnLogin = driverForUser.FindElement(By.ClassName("SignIn-btn"));

            btnLogin.Click();
        }

        public void addToCheckOut()
        {
            Thread.Sleep(1000);
            driverForUser.Url = "http://localhost/a/McDelivery/mcMenu.php";
            driverForUser.Navigate();
            Thread.Sleep(5000);

            WebDriverWait wait = new WebDriverWait(driverForUser, TimeSpan.FromSeconds(8));

            IList<IWebElement> foodItem = driverForUser.FindElements(By.ClassName("product-items"));

            //Chọn mua sản phẩm đầu tiên
            IWebElement orderPick = foodItem[0];

            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("btn-increase")));

            IWebElement btnAdd = orderPick.FindElement(By.ClassName("btn-increase"));

            Console.WriteLine(btnAdd.Text);

            //Thêm vào để thanh toán
            btnAdd.Click();

            Thread.Sleep(3000);

            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("checkout-button")));

            IWebElement btnMoveToCheckOut = driverForUser.FindElement(By.ClassName("checkout-button"));

            btnMoveToCheckOut.Click();

        }

        public void handleCheckOutForm(int row)
        {
            string tempoemail = "phamtrantruc@gmail.com";
   
            Thread.Sleep(4000);
            WebDriverWait wait = new WebDriverWait(driverForUser, TimeSpan.FromSeconds(8));

            wait.Until(ExpectedConditions.ElementIsVisible(By.TagName("input")));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("check_promo")));

            //Tìm tất cả các thẻ input
            IList<IWebElement> input = driverForUser.FindElements(By.TagName("input"));

            //Trường nhập địa chỉ nằm ở input thứ 2
            ipnAddress = input[1];
            ipnPhone = input[2];//input thứ 3
            ipnEmail = input[3];//input thứ 4
            ipnDiscount = input[4];//input thứ 5

            ipnAddress.SendKeys("66/1 Nguyễn tuyển , phường Bình trưng Bruh");
            ipnPhone.SendKeys("0963674165");
            ipnPhone.SendKeys(tempoemail);
            ipnDiscount.SendKeys(deletedCoupon);

            IWebElement btnCheck = driverForUser.FindElement(By.Id("check_promo"));
            btnCheck.Click();

            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("promo_feedback")));
            IWebElement alert = driverForUser.FindElement(By.Id("promo_feedback"));

            CheckDiscountValid(row, alert.Text);

            Thread.Sleep(5000);
        }

        public void CheckDiscountValid(int row,string result)
        {
            if(result== "Mã không hợp lệ hoặc đã hết hạn.")
            {
                InputResult(row, result, "PASSED");
            }
            else
            {
                InputResult(row, result, "FAILED");
            }
        }

        public void SetKeysForInputLogin(string email, string pwd)
        {
            mailIpn.Clear();
            pwdIpn.Clear();

            mailIpn.SendKeys(email);
            pwdIpn.SendKeys(pwd);
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
            ExcelPage(5);
            worksheet.Cells[row, 12].Value = result; // Ghi kết quả vào cột 10 (Result)
            worksheet.Cells[row, 13].Value = status; // Ghi kết quả passed or failed
        }

        public void ClickElementIfInteractable(By locator, IWebDriver driver)
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
                driverForUser.Quit();
                driverForAdmin.Quit();


                // Release COM objects
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
