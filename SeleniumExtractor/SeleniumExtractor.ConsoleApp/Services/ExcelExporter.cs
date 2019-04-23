using ClosedXML.Excel;
using SeleniumExtractor.ConsoleApp.Configurations;
using SeleniumExtractor.ConsoleApp.PageObjects.Clients;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumExtractor.ConsoleApp.Services
{
    public class ExcelExporter
    {
        private ExcelConfigurations _configurations;

        public ExcelExporter(ExcelConfigurations configurations)
        {
            _configurations = configurations;
        }

        public string ExportClientsList(IEnumerable<Client> clients)
        {
            if (!Directory.Exists(_configurations.ClientsExportPath))
                Directory.CreateDirectory(_configurations.ClientsExportPath);

            var destinationFile = _configurations.ClientsExportPath + $"Clients {DateTime.Now.ToString("yyy-MM-dd HH_mm_ss")}.xlsx";
            File.Copy(_configurations.ClientsExportTemplate, destinationFile);

            using (var workbook = new XLWorkbook(destinationFile))
            {
                var worksheet = workbook.Worksheets.Worksheet("Clients");
                var initialCell = worksheet.Cell("A2");

                for (int i = 0; i < clients.Count(); i++)
                {
                    var client = clients.ElementAt(i);
                    initialCell.CellBelow(i).Value = client.Name;
                    initialCell.CellBelow(i).CellRight(1).Value = client.Email;
                    initialCell.CellBelow(i).CellRight(2).Value = client.Birthdate;
                }

                workbook.Save();
            }

            return destinationFile;
        }
    }
}
