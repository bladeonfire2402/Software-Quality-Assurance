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



namespace Discount
{
    public class Discount
    {
        private IWebDriver driver;
        Application excelApp;
        Workbook workbook;
        Worksheet worksheet;

        private IWebElement ipnDiscount, ipnReleseDate, ipnExprired,ipnDescription;

        private static readonly string ExcelFilePath = @"D:\đảm bảo chất lượng phần mềm\selenium\TestData\TestCase.xlsx";

        [SetUp]
        public void Setup()
        {
            driver = new OpenQA.Selenium.Chrome.ChromeDriver();
            driver.Url = "http://localhost/a/admin/admin%20dashboard/";
            driver.Navigate();
            Thread.Sleep(1000);

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(ExcelFilePath, ReadOnly: false, Editable: true);
            ExcelPage(5);
            Thread.Sleep(1000);
        }

        [Test]
        public void handleTestDiscount()
        {
            const int startRow = 12;
            const int endRow = 24;
           

            for (int i = startRow; i < endRow + 1; i++)
            {
                if (i >= startRow && i <= startRow + 5)
                {
                    try
                    {
                        Console.WriteLine($"Chạy test case thu {i}");
                        TestCreate(i);

                    }
                    catch (NoSuchDriverException e)
                    {
                        Console.WriteLine($"Lỗi thứ {i} : {e}");
                    }
                }
                else if (i > startRow + 5 && i <= 23)
                {
                    Console.WriteLine($"Chạy test case thu {i}");
                    TestUpdate(i);
                }
                else if (i > 23)
                {
                    Console.WriteLine($"Chạy test case thu {i}");
                    TestDelete(i);
                }
                Console.WriteLine(i);
            }
        }


        public void TestCreate(int row)
        {
            EnterDiscountAddSection();

            Thread.Sleep(1000);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("code_khuyen_mai")));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ngay_phat_hanh")));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ngay_het_han")));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("mo_ta")));

            ipnDiscount = driver.FindElement(By.Id("code_khuyen_mai"));
            ipnReleseDate = driver.FindElement(By.Id("ngay_phat_hanh"));
            ipnExprired = driver.FindElement(By.Id("ngay_het_han"));
            ipnDescription = driver.FindElement(By.Id("mo_ta"));

            var strDiscount = handleDataTest(worksheet.Cells[row, 7].Value);
            var dateRelease = handleDataTest(worksheet.Cells[row, 8].Value);
            var dateExpired = handleDataTest(worksheet.Cells[row, 9].Value);
            var strDescription = handleDataTest(worksheet.Cells[row, 10].Value);

            Console.WriteLine($"Test data đây | Mã giảm giá: {strDiscount} | Ngày phát hành: {dateRelease} " +
          $"| Ngày hết hạn: {dateExpired} | Mô tả mã giảm giá: {strDescription}");

    
          
            //Send keys cho tất cả input 
            SetKeysForInput(ipnDiscount, ipnReleseDate, ipnExprired, ipnDescription, 
                strDiscount, dateRelease, dateExpired, strDescription, row);
            Thread.Sleep(2000);
            
            checkInvalidData(strDiscount, dateRelease, dateExpired, strDescription, row);
            
            ClickElementIfInteractable(By.Id("btn-add-discount"));


            if (row == 16)
            {
                var expectedResult = handleDataTest(worksheet.Cells[row, 11].Value);

                IWebElement body = wait.Until(ExpectedConditions.ElementIsVisible(By.TagName("body")));

                try
                {
                    string status = body.Text;

                    if(expectedResult == status)
                    {
                        InputResult(row, expectedResult, "PASSED");
                    }
                    else
                    {
                        InputResult(row, status, "FAILED");
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Lỗi rồi bruh" + e.Message);
                }
            }

            //Kiểm tra trường hợp nhập ngày không hợp lệ nhưng có cho phép tạo mã không
            if (row == 14)
            {
                checkIfCreateItemWithInvalidDate(strDiscount, row,"create");
            }

            //Kiểm tra xem tạo mã thành công chưa
            if (row == 12)
            {
                checkIfCreateOrUpdateSuccessfull(strDiscount, row, "create");
            }
        }


        public void TestUpdate(int row)
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(8));
            EnterDiscoutListSection();//Truy cập trang list mã giảm giá

            try
            {
                EnterUpdateFormSection();

                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("code_khuyen_mai")));
                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ngay_phat_hanh")));
                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ngay_het_han")));
                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("mo_ta")));

                ipnDiscount = driver.FindElement(By.Id("code_khuyen_mai"));
                ipnReleseDate = driver.FindElement(By.Id("ngay_phat_hanh"));
                ipnExprired = driver.FindElement(By.Id("ngay_het_han"));
                ipnDescription = driver.FindElement(By.Id("mo_ta"));

                var strDiscount = handleDataTest(worksheet.Cells[row, 7].Value);
                var dateRelease = handleDataTest(worksheet.Cells[row, 8].Value);
                var dateExpired = handleDataTest(worksheet.Cells[row, 9].Value);
                var strDescription = handleDataTest(worksheet.Cells[row, 10].Value);

                Console.WriteLine($"Test data đây | Mã giảm giá: {strDiscount} | Ngày phát hành: {dateRelease} " +
        $"| Ngày hết hạn: {dateExpired} | Mô tả mã giảm giá: {strDescription}");

                SetKeysForInput(ipnDiscount, ipnReleseDate, ipnExprired, ipnDescription,
               strDiscount, dateRelease, dateExpired, strDescription, row);
                Thread.Sleep(2000);

                ClickElementIfInteractable(By.Id("btn-update-discount"));

                checkInvalidData(strDiscount, dateRelease, dateExpired, strDescription, row);

                //Kiểm tra trường hợp nhập ngày không hợp lệ nhưng có cho phép tạo mã không
                if (row == 20)
                {
                    checkIfCreateItemWithInvalidDate(strDiscount, row, "update");
                }

                if (row == 22)
                {
                    var expectedResult = handleDataTest(worksheet.Cells[row, 11].Value);

                    IWebElement body = wait.Until(ExpectedConditions.ElementIsVisible(By.TagName("body")));

                    try
                    {
                        string status = body.Text;

                        if (expectedResult == status)
                        {
                            InputResult(row, expectedResult, "PASSED");
                        }
                        else
                        {
                            InputResult(row, status, "FAILED");
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Lỗi rồi bruh" + e.Message);
                    }
                }

                if (row == 18)//Riêng row 18 cần kiểm tra thêm
                {
                    checkIfCreateOrUpdateSuccessfull(strDiscount, row, "update");
                }


                Thread.Sleep(2000);
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine($"Đã xảy ra lỗi: {e.Message}");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Đã xảy ra lỗi {e}");
            }

        }

        public void TestDelete(int row)
        {
            EnterDiscoutListSection();

        
            try
            {
                IList<IWebElement> trElements = driver.FindElements(By.TagName("tr"));

                IWebElement fixElement = trElements[2];//Ở index 2
                Console.WriteLine(fixElement);

                IReadOnlyCollection<IWebElement> links = fixElement.FindElements(By.TagName("a"));
                Console.WriteLine(links.Count);

                IWebElement btnDelete = links.ElementAt(1);

                btnDelete.Click();

                IAlert alert = driver.SwitchTo().Alert();

                alert.Accept();

                InputResult(row, "Đã xóa thành công", "PASSED");

                Thread.Sleep(5000);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Lỗi nè {e.Message}");
            }

        }



        private void checkIfCreateItemWithInvalidDate(string discountId,int row,string isCreate)
        {
            driver.Navigate().GoToUrl("http://localhost/a/admin/admin%20dashboard/listmagiamgia.php");

            try
            {
                IList<IWebElement> trElements = driver.FindElements(By.TagName("tr"));

                foreach (var tr in trElements)
                {
                    var tds = driver.FindElements(By.TagName("td"));

                    foreach (IWebElement td in tds)
                    {
                        if (td.Text == discountId)
                        {
                            if (isCreate == "create")
                            {
                                InputResult(row, "Đã tạo mã giảm giá thành công", "FAILED");
                            }
                            else
                            {
                                InputResult(row, "Đã cập nhật mã giảm giá thành công", "FAILED");
                            }
                        }
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine($"Đã xảy ra lỗi: {e.Message}");
                InputResult(row, e.Message, "FAILED");
            }

        }

        public void checkIfCreateOrUpdateSuccessfull(string discountId, int row, string isCreate)
        {
            //Chuyển đến trang danh sách mã giảm giá
            driver.Navigate().GoToUrl("http://localhost/a/admin/admin%20dashboard/listmagiamgia.php");

            try
            {
                IList<IWebElement> trElements = driver.FindElements(By.TagName("tr"));

                foreach (var tr in trElements)
                {
                    var tds = driver.FindElements(By.TagName("td"));

                    foreach (IWebElement td in tds)
                    {
                        if (td.Text == discountId)
                        {
                            if (isCreate == "create")
                            {
                                InputResult(row, "Đã tạo mã giảm giá thành công", "PASSED");
                            }
                            else
                            {
                                InputResult(row, "Đã cập nhật mã giảm giá thành công", "PASSED");
                            }
                        }
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine($"Đã xảy ra lỗi: {e.Message}");
                InputResult(row, e.Message, "FAILED");
            }
        }


        public void CheckValidity(IWebElement e,int row)
        {
         
            // Sử dụng JavaScript để kiểm tra validity của phần tử
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            var validity = js.ExecuteScript(@"
                return arguments[0].validity.valid;", e);

            if ((bool)validity == false)
            {
                InputResult(row, "Ngày không hợp lệ", "FALSE");
                ClearUp();
            }
        }

        public void checkInvalidData(
           string strDiscount, string strdateRelease, string strdateExpired, string strDescription,int row)
        {
            if (strDiscount == "")
            {
                isRequiredAtribute(ipnDiscount, row);
            }
            else if(strdateRelease == "")
            {
                isRequiredAtribute(ipnReleseDate, row);

            }
            else if (strdateExpired == "")
            {
                isRequiredAtribute(ipnExprired, row);

            }
            else if (strDescription == "")
            {
                isRequiredAtribute(ipnDescription, row);

            }
        }

        public void isRequiredAtribute(IWebElement e, int row)
        {
            string requiredAttribute = e.GetAttribute("required");
            if (requiredAttribute == "false")
            {
                Console.WriteLine(requiredAttribute);
                InputResult(row, requiredAttribute, "FAILED");
            }
            else
            {
                Console.WriteLine(requiredAttribute);
                InputResult(row, "Vui lòng nhập vào trường", "PASSED");
            }
        }



        public void InputResult(int row, string result, string status)
        {
            worksheet.Cells[row, 12].Value = result; // Ghi kết quả vào cột 10 (Result)
            worksheet.Cells[row, 13].Value = status; // Ghi kết quả passed or failed
        }


        public void SetKeysForInput
            (IWebElement discountCode, IWebElement dateRelease, IWebElement dateExpired, IWebElement description,
            string strDiscount, string strdateRelease, string strdateExpired, string strDescription,
            int row)
        {
            try
            {
                //Clear nội dung bên trong
                discountCode.Clear();
                dateRelease.Clear();
                description.Clear();
                dateExpired.Clear();
                Thread.Sleep(3000);
                discountCode.SendKeys(strDiscount);
                dateRelease.SendKeys(strdateRelease);
                dateExpired.SendKeys(strdateExpired);
                description.SendKeys(strDescription);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
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
       


        public void handleIsOverDate(string dateRelease,string dateExpired,int row)
        {
            var culture = new System.Globalization.CultureInfo("en-GB"); // British culture
            DateTime releaseDate = DateTime.Parse(dateRelease, culture);
            DateTime expiredDate = DateTime.Parse(dateExpired, culture);

            if(releaseDate.Year > expiredDate.Year)
            {
                InputResult(row, "Năm không hợp lệ", "FAILED");
                ClearUp();
            }
            else
            {
                if (releaseDate.Month > expiredDate.Year)
                {
                    InputResult(row, "Tháng không hợp lệ", "FAILED");
                    ClearUp();
                }
                else
                {
                    if(releaseDate.Day > expiredDate.Day)
                    {
                        InputResult(row, "Ngày không hợp lệ", "FAILED");
                        ClearUp();
                    }
                }
            }
            
        }


        public void EnterDiscountAddSection()
        {
            ClickElementIfInteractable(By.Id("btn_access_discount"));
            ClickElementIfInteractable(By.Id("btn_add_discount"));
        }

        public void EnterDiscoutListSection()
        {
            ClickElementIfInteractable(By.Id("btn_access_discount"));
            ClickElementIfInteractable(By.Id("btn_list_discount"));
        }

        public void EnterUpdateFormSection()
        {
            try
            {
                IList<IWebElement> trElements = driver.FindElements(By.TagName("tr"));

                IWebElement fixElement = trElements[2];//Ở index 2
                Console.WriteLine(fixElement);

                IReadOnlyCollection<IWebElement> links = fixElement.FindElements(By.TagName("a"));
                Console.WriteLine(links.Count);

                IWebElement btnUpdate = links.ElementAt(0);

                driver.Navigate().GoToUrl(btnUpdate.GetAttribute("href"));
            }
            catch (Exception e)
            {
                Console.WriteLine($"Lỗi nè {e.Message}");
            }
        }

        public void ExcelPage(int page)
        {
            worksheet = workbook.Sheets[page];
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
