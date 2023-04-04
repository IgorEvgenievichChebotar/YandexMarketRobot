using System.IO;
using System.Linq;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var excel = new ExcelPackage();
var driver = new ChromeDriver();

try
{
    // Переход на страницу Яндекс Маркета
    driver.Navigate().GoToUrl("https://market.yandex.ru/");

    // Поиск поля для ввода запроса и кнопки для отправки запроса
    var searchInput = driver.FindElement(By.Name("text"));
    var searchButton = driver.FindElement(By.CssSelector("button.mini-suggest__button"));

    // Ввод запроса и нажатие кнопки поиска
    searchInput.SendKeys("Носки с дедом морозом");
    searchButton.Click();

    // Поиск первых 3 товаров и получение их информации
    var names = driver.FindElements(By.CssSelector("[data-baobab-name='title']")).Take(3).ToList();
    var prices = driver.FindElements(By.CssSelector("[data-auto='mainPrice']")).Take(3).ToList();
    var links = driver.FindElements(By.CssSelector("[data-baobab-name='title']")).Take(3).ToList();

    // Подготовка к созданию Excel-файла
    var worksheet = excel.Workbook.Worksheets.Add("Результаты поиска носков");

    // Заполнение Excel-файла информацией о товарах
    for (var i = 0; i < names.Count; i++)
    {
        worksheet.Cells[i + 1, 1].Value = names[i].Text;
        worksheet.Cells[i + 1, 2].Value = prices[i].Text;
        worksheet.Cells[i + 1, 3].Value = links[i].GetAttribute("href");
    }

    // Сохранение Excel-файла
    excel.SaveAs(new FileInfo(@"результаты поиска носков.xlsx"));
}
finally
{
    // Закрытие браузера и экселя
    driver.Quit();
    excel.Dispose();
}