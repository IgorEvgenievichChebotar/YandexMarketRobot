using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

IWebDriver driver = new ChromeDriver();

try
{
    // Переход на страницу Яндекс Маркета
    driver.Url = "https://market.yandex.ru/";

    // Поиск поля для ввода запроса и кнопки для отправки запроса
    IWebElement searchInput = driver.FindElement(By.Name("text"));
    IWebElement searchButton = driver.FindElement(By.CssSelector("button.mini-suggest__button"));

    // Ввод запроса и нажатие кнопки поиска
    searchInput.SendKeys("Носки с дедом морозом");
    searchButton.Click();

    // Поиск первых 3 товаров и получение их информации
    var names = driver.FindElements(By.CssSelector("[data-baobab-name='title']")).Take(3).ToList();
    var prices = driver.FindElements(By.CssSelector("[data-auto='mainPrice']")).Take(3).ToList();
    var links = driver.FindElements(By.CssSelector("[data-baobab-name='title']")).Take(3).ToList();

    // Подготовка к созданию Excel-файла
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    var package = new ExcelPackage();
    var worksheet = package.Workbook.Worksheets.Add("Результаты поиска носков");

    // Заполнение Excel-файла информацией о товарах
    for (var i = 0; i < names.Count; i++)
    {
        worksheet.Cells[i + 1, 1].Value = names[i].Text;
        worksheet.Cells[i + 1, 2].Value = prices[i].Text;
        worksheet.Cells[i + 1, 3].Value = links[i].GetAttribute("href");
    }

    // Сохранение Excel-файла
    package.SaveAs(new FileInfo(@"результаты поиска носков.xlsx"));
}
finally
{
    // Закрытие браузера
    driver.Quit();
}