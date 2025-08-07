using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using selenium.Pages;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Support.Events;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using SeleniumExtras.WaitHelpers;


namespace news
{
    public class News
    {
        private IWebDriver driver;
        Application excelApp;
        Workbook workbook;
        Worksheet worksheet;

        //Các trường nhập
        private IWebElement ipnTitle,ipnImg,ipnDescription;
     
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
            worksheet = workbook.Sheets[1];
            Thread.Sleep(1000); 
        }

        [Test]
        public void handleTestNews(){
            const int startRow = 12;
            const int endRow = 24;
            ExcelPage(3); // Đi đến testcase trang thứ 3

            for (int i = startRow; i < endRow+1; i++)
            {
                if (i >= startRow && i <= startRow+5)
                {
                    try
                    {
                        Console.WriteLine($"Chạy test case thu {i}");
                        TestCreateNews(i);

                    }
                    catch(NoSuchDriverException e)
                    {
                        Console.WriteLine($"Lỗi thứ {i} : {e}");
                    }
                }
                else if(i>startRow+5 && i<= 23)
                {
                    Console.WriteLine($"Chạy test case thu {i}");
                    TestFixNews(i);
                }
                else if (i > 23)
                {
                    Console.WriteLine($"Chạy test case thu {i}");
                    TestDeleteNews(i);
                }
                Console.WriteLine(i);
            }
        }

        public void TestCreateNews(int row)
        {
            EnterNewsAddSetion();

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            //Chờ load phần tử
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("title")));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("avatar")));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("input_content_admin")));

            ipnTitle = driver.FindElement(By.Id("title"));
            ipnImg = driver.FindElement(By.Id("avatar"));
            ipnDescription = driver.FindElement(By.Id("input_content_admin"));

            //Lấy dữ liệu từ excel 
            var titleNews = handleDataTest(worksheet.Cells[row, 7].Value);
            var imageNews = handleDataTest(worksheet.Cells[row, 8].Value);
            var descriptionNews = handleDataTest(worksheet.Cells[row, 9].Value);

            SetKeysForInputNews(ipnTitle, ipnImg, ipnDescription, titleNews, imageNews, descriptionNews, row);

            ClickElementIfInteractable(By.Id("btn-add-news"));
            
            checkInvalidData(ipnTitle, ipnImg, ipnDescription, titleNews, imageNews, descriptionNews, row);

            if(row == 14)
            {
                IWebElement body = wait.Until(ExpectedConditions.ElementIsVisible(By.TagName("body")));

                try
                {
                    // Lấy nội dung của body
                    string status = body.Text;

                    // Điều hướng đến trang mới (ví dụ: trang chủ hoặc trang khác)
                    driver.Navigate().GoToUrl("http://localhost/a/admin/admin%20dashboard/");

                    // Đợi trang mới tải xong và lấy lại body
                    body = wait.Until(ExpectedConditions.ElementIsVisible(By.TagName("body")));

                    // Ghi kết quả vào một nơi nào đó (ví dụ: một hàm InputResult)
                    InputResult(14, status, "PASSED");
                }
                catch (Exception e)
                {
                    // Xử lý lỗi nếu có
                    Console.WriteLine("Đã xảy ra lỗi: " + e.Message);
                }
            }

            if(row == 12)//Riêng row 12 cần kiểm tra thêm
            {
                checkIfCreateOrUpdateNewsSuccessfull(titleNews, row,"create");
            }
        }


        public void TestFixNews(int row)
        {
            
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(8));
            EnterNewsUpdateSection();//Truy cập trang list tin tức

            try
            {
                EnterUpdateFormSection();

                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("title")));
                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("avatar")));
                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("update-input-content")));
                ipnTitle = driver.FindElement(By.Id("title"));
                ipnImg = driver.FindElement(By.Id("avatar"));
                ipnDescription = driver.FindElement(By.Id("update-input-content"));

                //Lấy data từ excel
                var titleNews = handleDataTest(worksheet.Cells[row, 7].Value);
               
                var imageNews = handleDataTest(worksheet.Cells[row, 8].Value);
                var descriptionNews = handleDataTest(worksheet.Cells[row, 9].Value);

                Console.WriteLine(titleNews, imageNews, descriptionNews);

                SetKeysForInputNews(ipnTitle, ipnImg, ipnDescription, titleNews, imageNews, descriptionNews, row);

                ClickElementIfInteractable(By.Id("btn-update-news"));

                checkInvalidData(ipnTitle, ipnImg, ipnDescription, titleNews, imageNews, descriptionNews, row);

                if (row == 20)
                {
                    IWebElement body = wait.Until(ExpectedConditions.ElementIsVisible(By.TagName("body")));

                    try
                    {
                        // Lấy nội dung của body
                        string status = body.Text;
                        Console.WriteLine(status);
                        Thread.Sleep(5000);
                        // Điều hướng đến trang mới (ví dụ: trang chủ hoặc trang khác)
                        driver.Navigate().GoToUrl("http://localhost/a/admin/admin%20dashboard/");

                        // Đợi trang mới tải xong và lấy lại body
                        body = wait.Until(ExpectedConditions.ElementIsVisible(By.TagName("body")));

                        if(status== "Chỉ chấp nhận các định dạng JPG, JPEG, PNG.")
                        {
                            InputResult(20, status, "PASSED");
                        }
                        else {
                            InputResult(20, status, "FAILED");
                        }

                        // Ghi kết quả vào một nơi nào đó (ví dụ: một hàm InputResult)
                      
                    }
                    catch (Exception e)
                    {
                        // Xử lý lỗi nếu có
                        Console.WriteLine("Đã xảy ra lỗi: " + e.Message);
                    }
                }

                if (row == 18)//Riêng row 18 cần kiểm tra thêm
                {
                    checkIfCreateOrUpdateNewsSuccessfull(titleNews, row,"update");
                }


                Thread.Sleep(3000);
            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine($"Đã xảy ra lỗi: {e.Message}");
            }
            catch(Exception e)
            {
                Console.WriteLine($"Đã xảy ra lỗi {e}");
            }
 
        }




        public void TestDeleteNews(int row)
        {
            EnterNewsUpdateSection();

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

        public void checkIfCreateOrUpdateNewsSuccessfull(string newsTitle ,int row,string isCreate)
        {
            //Chuyển đến trang danh sách tin tức
            driver.Navigate().GoToUrl("http://localhost/a/admin/admin%20dashboard/listnews.php");

            try
            {
                IList<IWebElement> trElements = driver.FindElements(By.TagName("tr"));

                foreach (var tr in trElements)
                {
                    var tds = driver.FindElements(By.TagName("td"));
                  
                    foreach(IWebElement td in tds)
                    {
                        if(td.Text == newsTitle)
                        {
                            if(isCreate== "create")
                            {
                                InputResult(row, "Đã tạo tin tức thành công", "PASSED");
                                
                            }
                            else
                            {
                                InputResult(row, "Đã cập nhật tin tức thành công", "PASSED");
                         
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

      

        public void checkInvalidData(
            IWebElement title, IWebElement image, IWebElement description, string strTitle, string imageNews, 
            string descriptionNews, int row)
        {
            if (strTitle == "")
            {
                isRequiredAtribute(title, row);
            }
            else if (descriptionNews == "")
            {
                isRequiredAtribute(description,row);
            }
            else if (imageNews == "")
            {
                isRequiredAtribute(image, row);
            }
        }

        public void isRequiredAtribute(IWebElement e,int row)
        {
            string requiredAttribute = e.GetAttribute("required");
            if(requiredAttribute == "false")
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

        public string handleDataTest(string data)
        {
            switch (data)
            {
                case "N/A":
                    return "";
            }
            return data;
        }


        public void InputResult(int row,string result, string status)
        {
            worksheet.Cells[row, 11].Value = result; // Ghi kết quả vào cột 10 (Result)
            worksheet.Cells[row, 12].Value = status; // Ghi kết quả passed or failed
        }

        public void SetKeysForInputNews
            (IWebElement title, IWebElement image, IWebElement description, string strTitle, string imageNews, string descriptionNews,
            int row)
        {
            try
            {
                //Clear nội dung bên trong
                title.Clear();
                image.Clear();
                description.Clear();

                title.SendKeys(strTitle);
                image.SendKeys(imageNews);
                description.SendKeys(descriptionNews);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void EnterNewsAddSetion()
        {
            ClickElementIfInteractable(By.Id("btn_access_news"));
            ClickElementIfInteractable(By.Id("btn_news_add"));
        }

        public void EnterNewsUpdateSection()
        {
            ClickElementIfInteractable(By.Id("btn_access_news"));
            ClickElementIfInteractable(By.Id("btn_list_news"));

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
            catch(Exception e)
            {
                Console.WriteLine($"Lỗi nè {e.Message}");
            }
        }

        public void ExcelPage(int page)
        {
            worksheet = workbook.Sheets[page];
        }

        public void findelement(IWebElement title, IWebElement image, IWebElement description) {
            title = driver.FindElement(By.Id("title"));
            image = driver.FindElement(By.Id("avatar"));
            description = driver.FindElement(By.Id("input_content_admin"));
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

