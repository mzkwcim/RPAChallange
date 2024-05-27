using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Threading;

internal class Program
{
    private static void Main(string[] args)
    {
        string filePath = @"C:\users\mzkwcim\Downloads\challenge.xlsx";

        List<CustomStructure> customStructure = GetCustomStructureFromExcel(filePath);

        using (IWebDriver driver = new ChromeDriver())
        {
            driver.Navigate().GoToUrl("https://rpachallenge.com/");
            ClickButton(driver, "//button[text()='Start']");

            foreach (var cs in customStructure)
            {
                EnterData(driver, "Phone Number", cs.PhoneNumber.Trim());
                EnterData(driver, "First Name", cs.FirstName.Trim());
                EnterData(driver, "Last Name", cs.LastName.Trim());
                EnterData(driver, "Email", cs.Email.Trim());
                EnterData(driver, "Company Name", cs.CompanyName.Trim());
                EnterData(driver, "Role in Company", cs.RoleInCompany.Trim());
                EnterData(driver, "Address", cs.Address.Trim());

                ClickButton(driver, "//input[@value='Submit']");
            }
            Console.ReadKey();
        }
    }

    private static List<CustomStructure> GetCustomStructureFromExcel(string filePath)
    {
        List<CustomStructure> customStructure = new List<CustomStructure>();

        using (XLWorkbook workbook = new XLWorkbook(filePath))
        {
            IXLWorksheet worksheet = workbook.Worksheet(1);

            Dictionary<int, string> columnMappings = new Dictionary<int, string>
            {
                { 1, "FirstName" },
                { 2, "LastName" },
                { 3, "CompanyName" },
                { 4, "RoleInCompany" },
                { 5, "Email" },
                { 6, "Address" },
                { 7, "PhoneNumber" }
            };

            foreach (IXLRow row in worksheet.RowsUsed().Skip(1))
            {
                CustomStructure cs = new CustomStructure();

                foreach (IXLCell cell in row.Cells())
                {
                    int columnIndex = cell.Address.ColumnNumber;
                    string columnName = columnMappings[columnIndex];

                    typeof(CustomStructure).GetProperty(columnName)?.SetValue(cs, cell.Value.ToString().Trim());
                }

                customStructure.Add(cs);
            }
        }

        return customStructure;
    }

    private static void ClickButton(IWebDriver driver, string xpath)
    {
        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        IWebElement button = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(xpath)));
        button.Click();
    }

    private static void EnterData(IWebDriver driver, string labelText, string data)
    {
        IWebElement label = driver.FindElement(By.XPath($"//label[contains(text(), '{labelText}')]"));
        IWebElement inputField = label.FindElement(By.XPath("./following-sibling::input"));
        inputField.SendKeys(data);
    }
}

class CustomStructure
{
    public string FirstName { get; set; } = "";
    public string LastName { get; set; } = "";
    public string CompanyName { get; set; } = "";
    public string RoleInCompany { get; set; } = "";
    public string Email { get; set; } = "";
    public string Address { get; set; } = "";
    public string PhoneNumber { get; set; } = "";
}
