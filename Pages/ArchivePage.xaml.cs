using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using Lab7.Models;

namespace Lab7.Pages
{
    public partial class ArchivePage : Page
    {
        private bool _editMode = false;
        private int _editIndex = -1;  // индекс редактируемой записи (-1 = новая)

        public ArchivePage()
        {
            InitializeComponent();
        }

        private void Page_Initialized(object sender, EventArgs e)
        {
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            dataGrid.ItemsSource = null;
            dataGrid.ItemsSource = Archive.Items;
        }

        // ── Переключение режима ───────────────────────────────────────────────
        private void rbView_Checked(object sender, RoutedEventArgs e)
        {
            _editMode = false;
            if (editPanel != null) editPanel.Visibility = Visibility.Collapsed;
            if (dataGrid  != null) dataGrid.IsReadOnly  = true;
        }

        private void rbEdit_Checked(object sender, RoutedEventArgs e)
        {
            _editMode = true;
            if (editPanel != null) editPanel.Visibility = Visibility.Visible;
        }

        // ── Добавить запись ───────────────────────────────────────────────────
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            if (!_editMode) { MessageBox.Show("Переключитесь в режим «Редактирование»."); return; }
            _editIndex = -1;
            ClearForm();
            editPanel.Visibility = Visibility.Visible;
        }

        // ── Удалить запись ────────────────────────────────────────────────────
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (!_editMode) { MessageBox.Show("Переключитесь в режим «Редактирование»."); return; }
            if (dataGrid.SelectedItem is Item item)
            {
                if (MessageBox.Show($"Удалить «{item.Name}»?", "Подтверждение",
                        MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    Archive.Items.Remove(item);
                    RefreshGrid();
                }
            }
            else MessageBox.Show("Выберите запись для удаления.");
        }

        // ── Сохранить (применить изменения из формы) ─────────────────────────
        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (!_editMode) return;
            if (!TryParseForm(out Item item)) return;

            if (_editIndex >= 0 && _editIndex < Archive.Items.Count)
                Archive.Items[_editIndex] = item;
            else
                Archive.Items.Add(item);

            RefreshGrid();
            ClearForm();
            _editIndex = -1;
        }

        // Двойной клик — загрузить запись в форму для редактирования
        private void dataGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (!_editMode) return;
            if (dataGrid.SelectedItem is Item item)
            {
                _editIndex = Archive.Items.IndexOf(item);
                tbName.Text    = item.Name;
                tbInvNum.Text  = item.InventoryNumber.ToString();
                tbLab.Text     = item.LabNumber.ToString();
                tbYear.Text    = item.PurchaseYear.ToString();
                tbMonth.Text   = item.PurchaseMonth.ToString();
                tbCost.Text    = item.Cost.ToString();
                tbService.Text = item.ServiceLife.ToString();
                tbQty.Text     = item.Quantity.ToString();
            }
        }

        // ── Задания ───────────────────────────────────────────────────────────
        private void Task1_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Uri("Pages/Task1Page.xaml", UriKind.Relative));
        }

        private void Task2_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Uri("Pages/Task2Page.xaml", UriKind.Relative));
        }

        // ── Вспомогательные ──────────────────────────────────────────────────
        private void ClearForm()
        {
            tbName.Text = tbInvNum.Text = tbLab.Text = tbYear.Text =
            tbMonth.Text = tbCost.Text = tbService.Text = tbQty.Text = "";
        }

        private bool TryParseForm(out Item item)
        {
            item = null;
            if (!int.TryParse(tbInvNum.Text, out int invNum) || invNum < 1000 || invNum > 9999)
            { MessageBox.Show("Инвентарный номер: 4 цифры (1000–9999)."); return false; }
            if (!int.TryParse(tbLab.Text, out int lab))
            { MessageBox.Show("Номер лаборатории: целое число."); return false; }
            if (!int.TryParse(tbYear.Text, out int year) || year < 1900 || year > 2100)
            { MessageBox.Show("Год приобретения: от 1900 до 2100."); return false; }
            if (!int.TryParse(tbMonth.Text, out int month) || month < 1 || month > 12)
            { MessageBox.Show("Месяц: от 1 до 12."); return false; }
            if (!decimal.TryParse(tbCost.Text, out decimal cost) || cost < 0)
            { MessageBox.Show("Стоимость: неотрицательное число."); return false; }
            if (!int.TryParse(tbService.Text, out int service) || service < 1)
            { MessageBox.Show("Срок службы: целое положительное число."); return false; }
            if (!int.TryParse(tbQty.Text, out int qty) || qty < 0)
            { MessageBox.Show("Количество: неотрицательное целое."); return false; }
            if (string.IsNullOrWhiteSpace(tbName.Text))
            { MessageBox.Show("Введите наименование предмета."); return false; }

            item = new Item
            {
                Name            = tbName.Text.Trim(),
                InventoryNumber = invNum,
                LabNumber       = lab,
                PurchaseYear    = year,
                PurchaseMonth   = month,
                Cost            = cost,
                ServiceLife     = service,
                Quantity        = qty
            };
            return true;
        }
    }
}
