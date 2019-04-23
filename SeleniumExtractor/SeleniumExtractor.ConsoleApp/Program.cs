using Microsoft.Extensions.Configuration;
using SeleniumExtractor.ConsoleApp.Configurations;
using SeleniumExtractor.ConsoleApp.PageObjects.Clients;
using SeleniumExtractor.ConsoleApp.Services;
using System.Collections.Generic;
using System.IO;

namespace SeleniumExtractor.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json");

            var configuration = builder.Build();

            var seleniumConfigurations = new SeleniumConfigurations();
            var excelConfigurations = new ExcelConfigurations();
            configuration.GetSection("SeleniumConfigurations").Bind(seleniumConfigurations);
            configuration.GetSection("ExcelConfigurations").Bind(excelConfigurations);

            var clients = GetClients(seleniumConfigurations);

            ExportClientsToExcel(excelConfigurations, clients);
        }

        private static IEnumerable<Client> GetClients(SeleniumConfigurations seleniumConfigurations)
        {
            var clientsPage = new ClientsPageObject(seleniumConfigurations);
            clientsPage.LoadPage();

            var clients = clientsPage.GetClients();

            //clientsPage.ClosePage();

            return clients;
        }

        private static void ExportClientsToExcel(ExcelConfigurations configurations, IEnumerable<Client> clients)
        {
            var excelExporter = new ExcelExporter(configurations);
            excelExporter.ExportClientsList(clients);
        }
    }
}
