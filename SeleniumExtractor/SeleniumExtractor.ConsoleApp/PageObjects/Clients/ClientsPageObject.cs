using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using SeleniumExtractor.ConsoleApp.Configurations;
using System;
using System.Collections.Generic;

namespace SeleniumExtractor.ConsoleApp.PageObjects.Clients
{
    public class ClientsPageObject
    {
        private SeleniumConfigurations _configurations;
        private IWebDriver _driver;

        public ClientsPageObject(SeleniumConfigurations configurations)
        {
            _configurations = configurations;

            ChromeOptions chromeOptions = new ChromeOptions();

            if (_configurations.Headless)
            {
                chromeOptions.AddArguments("--window-size=1920,1080");
                chromeOptions.AddArguments("--disable-gpu");
                chromeOptions.AddArguments("--disable-extensions");
                chromeOptions.AddArguments("--proxy-server='direct://'");
                chromeOptions.AddArguments("--proxy-bypass-list=*");
                chromeOptions.AddArguments("--start-maximized");
                chromeOptions.AddArguments("--headless");
            }
            _driver = new ChromeDriver(_configurations.DriverChromePath, chromeOptions);
        }

        public void LoadPage()
        {
            _driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(_configurations.Timeout);
            _driver.Navigate().GoToUrl(_configurations.ClientsPageUrl);
        }

        public void ClosePage()
        {
            _driver.Quit();
            _driver = null;
        }

        public IEnumerable<Client> GetClients()
        {
            var clients = new List<Client>();

            var clientRowElements = _driver.FindElement(By.Id("clientsTable"))
                .FindElement(By.TagName("tbody"))
                .FindElements(By.TagName("tr"));

            foreach (var rowElement in clientRowElements)
            {
                var rowData = rowElement.FindElements(By.TagName("td"));
                var name = rowData[0].Text;
                var email = rowData[1].Text;
                var birthdate = rowData[2].Text;
                clients.Add(new Client
                {
                    Name = name,
                    Email = email,
                    Birthdate = birthdate
                });
            }

            return clients;
        }
    }
}
