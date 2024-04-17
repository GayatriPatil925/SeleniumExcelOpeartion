using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOpeartion
{
    public class Tests
    {
        public IWebDriver driver;
        public int row, col;
        //use the excel app which is currently available 
        public Excel.Application excelApp;
        public Excel.Workbook workBook;
        public Excel.Worksheet workSheet;
        public string filePath= @"C:\Users\91755\source\repos\ExcelOpeartion\ExcelOpeartion\test.xlsx";
        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            driver.Manage().Window.Maximize();
            driver.Url = "https://www.microsoft.com/en-us/d/surface-laptop-studio-2/8rqr54krf1dz";
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            try
            {
                IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@class='emailSup-swapContent']/following-sibling::button[@aria-label='Close dialog window']")));
                element.Click();
            }
            catch
            {
                Console.WriteLine("Email Signup not present.");
            }
            excelApp =new Excel.Application();
        }
        public void ExcelWrite()
        {
            //to create excel file\workbook
            workBook = excelApp.Workbooks.Add();
            //to select sheet
            workSheet = workBook.Sheets[1];  
            //whenever operation going on we can able to see the excel app
            excelApp.Visible = true;
        }
        public void ExcelRead()
        {
            workBook=excelApp.Workbooks.Open(filePath);
            workSheet = workBook.Sheets[1];
            excelApp.Visible = false;
            row = workSheet.UsedRange.Rows.Count;
            col = workSheet.UsedRange.Columns.Count;
            
        }
        [TearDown]
        public void Teardown()
        {
            workBook.Close(false);
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            driver.Dispose();
            driver.Quit();
        }
        [Test]
        public void ExcelWritingData()
        {
            ExcelWrite();
            //specifying row & column where ImgUrl shoul be written
            workSheet.Cells[1, 1] = "ImgUrl";
            //specifying row & column where Alt Text shoul be written
            workSheet.Cells[1, 2] = "Alt Text";
            //gettiing imgurl and alt text in list 
            IList<IWebElement> imgList = driver.FindElements(By.XPath("//img[contains(@src,'https://') and not(contains(@src,'https://bat'))]"));
            int i = 2;
            foreach (IWebElement img in imgList)
            {
                if (img.GetAttribute("alt") != "")
                {
                    workSheet.Cells[i, 1] = img.GetAttribute("src").ToString();
                    workSheet.Cells[i, 2] = img.GetAttribute("alt").ToString();
                    i++;
                }
            }
            //to save the file
            workBook.SaveAs(filePath);
        }
        [Test]
        public void ExcelReadingData()
        {
            ExcelRead();
            for (int i = 2; i <= row; i++)
            {
                //value 2 = to read the excel data in the form of string
                Console.WriteLine(workSheet.Cells[i, 1].Value2 + " : Alt text is: " + workSheet.Cells[i, 2].Value2);
            }
        }
    }
}