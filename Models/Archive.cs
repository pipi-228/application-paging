using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;  // Псевдоним для устранения конфликта Range
using Microsoft.Data.SqlClient;                // Замена System.Data.SqlClient
using Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using System.Windows;

namespace Lab7.Models
{
    /// <summary>
    /// Статическое хранилище архива (передаётся между страницами)
    /// </summary>
    public static class Archive
    {
        public static List<Item> Items { get; set; } = new List<Item>();

        // ─── XML ─────────────────────────────────────────────────────────────
        public static void LoadFromXml(string path)
        {
            Items = new List<Item>();
            XDocument doc = XDocument.Load(path);
            foreach (XElement el in doc.Root.Elements("item"))
            {
                Items.Add(new Item
                {
                    Name = (string)el.Element("name"),
                    InventoryNumber = (int)el.Element("inventoryNumber"),
                    LabNumber = (int)el.Element("labNumber"),
                    PurchaseYear = (int)el.Element("purchaseYear"),
                    PurchaseMonth = (int)el.Element("purchaseMonth"),
                    Cost = (decimal)(double)el.Element("cost"),
                    ServiceLife = (int)el.Element("serviceLife"),
                    Quantity = (int)el.Element("quantity")
                });
            }
        }

        public static void SaveToXml(string path)
        {
            var doc = new XDocument(
                new XElement("archive",
                    Items.Select(i => new XElement("item",
                        new XElement("name", i.Name),
                        new XElement("inventoryNumber", i.InventoryNumber),
                        new XElement("labNumber", i.LabNumber),
                        new XElement("purchaseYear", i.PurchaseYear),
                        new XElement("purchaseMonth", i.PurchaseMonth),
                        new XElement("cost", i.Cost),
                        new XElement("serviceLife", i.ServiceLife),
                        new XElement("quantity", i.Quantity)
                    ))
                )
            );
            doc.Save(path);
        }

        // ─── Excel ───────────────────────────────────────────────────────────
        // Столбцы: A-Name, B-InventoryNumber, C-LabNumber,
        //          D-PurchaseYear, E-PurchaseMonth, F-Cost, G-ServiceLife, H-Quantity
        // ─── Excel ───────────────────────────────────────────────────────────
        // Столбцы: A-Name, B-InventoryNumber, C-LabNumber,
        //          D-PurchaseYear, E-PurchaseMonth, F-Cost, G-ServiceLife, H-Quantity
        public static void LoadFromExcel(string path)
        {
            Items = new List<Item>();

            using (var workbook = new XLWorkbook(path))
            {
                var worksheet = workbook.Worksheet(1);

                // Находим последнюю использованную строку
                var lastRow = worksheet.LastRowUsed();

                if (lastRow == null)
                {
                    MessageBox.Show("Файл Excel не содержит данных!");
                    return;
                }

                int lastRowNumber = lastRow.RowNumber();

                // Читаем со 2 строки до последней включительно
                for (int row = 2; row <= lastRowNumber; row++)
                {
                    var nameCell = worksheet.Cell(row, 1);
                    if (string.IsNullOrWhiteSpace(nameCell.GetString()))
                        continue;

                    try
                    {
                        Item item = new Item
                        {
                            Name = nameCell.GetString(),
                            InventoryNumber = worksheet.Cell(row, 2).GetValue<int>(),
                            LabNumber = worksheet.Cell(row, 3).GetValue<int>(),
                            PurchaseYear = worksheet.Cell(row, 4).GetValue<int>(),
                            PurchaseMonth = worksheet.Cell(row, 5).GetValue<int>(),
                            Cost = worksheet.Cell(row, 6).GetValue<decimal>(),
                            ServiceLife = worksheet.Cell(row, 7).GetValue<int>(),
                            Quantity = worksheet.Cell(row, 8).GetValue<int>()
                        };
                        Items.Add(item);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка в строке {row}: {ex.Message}\n" +
                                       $"Проверьте формат данных (возможно, пустая ячейка с числовым форматом)");
                        // Пропускаем проблемную строку
                        continue;
                    }
                }
            }

            MessageBox.Show($"Загружено {Items.Count} записей из Excel файла!");
        }

        // ─── SQL Server ───────────────────────────────────────────────────────
        // Таблица: items(name, inventoryNumber, labNumber, purchaseYear,
        //                purchaseMonth, cost, serviceLife, quantity)
        public static void LoadFromDatabase(string connectionString)
        {
            Items = new List<Item>();
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("SELECT * FROM items", conn);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Items.Add(new Item
                    {
                        Name = reader.GetString(0),
                        InventoryNumber = reader.GetInt32(1),
                        LabNumber = reader.GetInt32(2),
                        PurchaseYear = reader.GetInt32(3),
                        PurchaseMonth = reader.GetInt32(4),
                        Cost = reader.GetDecimal(5),
                        ServiceLife = reader.GetInt32(6),
                        Quantity = reader.GetInt32(7)
                    });
                }
            }
        }

        // ─── Тестовые данные ─────────────────────────────────────────────────
        public static void LoadSampleData()
        {
            Items = new List<Item>
            {
                new Item { Name="Компьютер",        InventoryNumber=1001, LabNumber=1, PurchaseYear=2015, PurchaseMonth=3,  Cost=25000, ServiceLife=10, Quantity=5 },
                new Item { Name="Монитор",           InventoryNumber=1002, LabNumber=1, PurchaseYear=2015, PurchaseMonth=3,  Cost=8000,  ServiceLife=8,  Quantity=5 },
                new Item { Name="Принтер",           InventoryNumber=1003, LabNumber=2, PurchaseYear=2018, PurchaseMonth=9,  Cost=12000, ServiceLife=7,  Quantity=2 },
                new Item { Name="Проектор",          InventoryNumber=1004, LabNumber=3, PurchaseYear=2017, PurchaseMonth=1,  Cost=35000, ServiceLife=12, Quantity=1 },
                new Item { Name="Маршрутизатор",     InventoryNumber=1005, LabNumber=2, PurchaseYear=2019, PurchaseMonth=6,  Cost=5000,  ServiceLife=5,  Quantity=3 },
                new Item { Name="Осциллограф",       InventoryNumber=1006, LabNumber=4, PurchaseYear=2010, PurchaseMonth=11, Cost=18000, ServiceLife=15, Quantity=2 },
                new Item { Name="Сервер",            InventoryNumber=1007, LabNumber=1, PurchaseYear=2020, PurchaseMonth=4,  Cost=90000, ServiceLife=10, Quantity=1 },
                new Item { Name="Веб-камера",        InventoryNumber=1008, LabNumber=3, PurchaseYear=2021, PurchaseMonth=8,  Cost=3000,  ServiceLife=4,  Quantity=6 },
                new Item { Name="Стенд",             InventoryNumber=1009, LabNumber=4, PurchaseYear=2008, PurchaseMonth=2,  Cost=45000, ServiceLife=20, Quantity=1 },
                new Item { Name="Клавиатура",        InventoryNumber=1010, LabNumber=1, PurchaseYear=2022, PurchaseMonth=1,  Cost=1500,  ServiceLife=3,  Quantity=10 },
                new Item { Name="ИБП",               InventoryNumber=1011, LabNumber=2, PurchaseYear=2016, PurchaseMonth=7,  Cost=6000,  ServiceLife=9,  Quantity=4 },
                new Item { Name="Коммутатор",        InventoryNumber=1012, LabNumber=2, PurchaseYear=2014, PurchaseMonth=5,  Cost=9000,  ServiceLife=11, Quantity=2 },
            };
        }
    }
}