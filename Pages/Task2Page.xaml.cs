using System;
using System.Linq;
using System.Windows.Controls;
using Lab7.Models;

namespace Lab7.Pages
{
    // Обёртка с порядковым номером для отображения в гриде
    public class RankedItem : Item
    {
        public int Rank { get; set; }
    }

    public partial class Task2Page : Page
    {
        public Task2Page()
        {
            InitializeComponent();
        }

        private void Page_Initialized(object sender, EventArgs e)
        {
            var top5 = Archive.Items
                .OrderByDescending(i => i.ServiceLife)
                .Take(5)
                .Select((item, idx) => new RankedItem
                {
                    Rank            = idx + 1,
                    Name            = item.Name,
                    InventoryNumber = item.InventoryNumber,
                    LabNumber       = item.LabNumber,
                    PurchaseYear    = item.PurchaseYear,
                    PurchaseMonth   = item.PurchaseMonth,
                    Cost            = item.Cost,
                    ServiceLife     = item.ServiceLife,
                    Quantity        = item.Quantity
                })
                .ToList();

            resultGrid.ItemsSource = top5;
        }
    }
}
