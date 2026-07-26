using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using Lab7.Models;
using Microsoft.Win32;

namespace Lab7.Pages
{
    public partial class MainPage : Page
    {
        public MainPage()
        {
            InitializeComponent();
            Loaded += (s, e) => UpdateStatus();
        }

        private void UpdateStatus()
        {
            statusText.Text = Archive.Items.Count > 0
                ? $"Загружено записей: {Archive.Items.Count}"
                : "Архив не загружен.";
        }

        // ── Загрузка ─────────────────────────────────────────────────────────
        private void LoadXml_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "XML files|*.xml", Title = "Открыть XML архив" };
            if (dlg.ShowDialog() == true)
            {
                try   { Archive.LoadFromXml(dlg.FileName); UpdateStatus(); NavigateToArchive(); }
                catch (Exception ex) { MessageBox.Show("Ошибка загрузки XML:\n" + ex.Message); }
            }
        }

        private void LoadExcel_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Excel files|*.xlsx;*.xls", Title = "Открыть Excel архив" };
            if (dlg.ShowDialog() == true)
            {
                try   { Archive.LoadFromExcel(dlg.FileName); UpdateStatus(); NavigateToArchive(); }
                catch (Exception ex) { MessageBox.Show("Ошибка загрузки Excel:\n" + ex.Message); }
            }
        }

        private void LoadDb_Click(object sender, RoutedEventArgs e)
        {
            string connStr = $@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename={
                Environment.CurrentDirectory}\lab7.mdf;Integrated Security=True";
            try   { Archive.LoadFromDatabase(connStr); UpdateStatus(); NavigateToArchive(); }
            catch (Exception ex) { MessageBox.Show("Ошибка подключения к БД:\n" + ex.Message); }
        }

        private void LoadSample_Click(object sender, RoutedEventArgs e)
        {
            Archive.LoadSampleData();
            UpdateStatus();
            NavigateToArchive();
        }

        private void NavigateToArchive()
        {
            NavigationService.Navigate(new Uri("Pages/ArchivePage.xaml", UriKind.Relative));
        }

        // ── Меню ─────────────────────────────────────────────────────────────
        private void About_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Материальные ценности кафедры\n\n", "О программе");
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
    }
}
