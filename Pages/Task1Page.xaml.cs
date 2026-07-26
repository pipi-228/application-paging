using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Lab7.Models;

namespace Lab7.Pages
{
    public partial class Task1Page : Page
    {
        public Task1Page()
        {
            InitializeComponent();
        }

        private void Page_Initialized(object sender, EventArgs e)
        {
            tbYear.Text = DateTime.Now.Year.ToString();
            RunQuery();
        }

        private void Show_Click(object sender, RoutedEventArgs e)
        {
            RunQuery();
        }

        private void RunQuery()
        {
            if (!int.TryParse(tbYear.Text, out int year))
            {
                MessageBox.Show("Введите корректный год (целое число).");
                return;
            }

            // Предмет подлежит списанию в year, если год_покупки + срок_службы == year
            var result = Archive.Items
                .Where(i => i.WriteOffYear == year)
                .OrderBy(i => i.Name)
                .ToList();

            resultGrid.ItemsSource = result;
            countLabel.Text = $"Найдено: {result.Count} шт.";
        }
    }
}
