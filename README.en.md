[Русский](README.md) | **English**

# Material Assets Registry — WPF Paging Application

A C# WPF application built on the page-based navigation model (`Page` + `NavigationWindow`) for tracking material assets — computers, monitors, printers, and other equipment assigned to a department. The app loads the asset archive from Excel, XML, or a database, lets you view and edit records, and generates filtered reports.

## Features

- **Load the archive** from one of three user-selectable sources: Excel (`.xlsx`), XML (`baza.xml`), or a database.
- **View the archive** — display all records in a table.
- **Add** new records to the archive.
- **Delete** existing records.
- **Edit** specific record fields: quantity, cost, service life.
- **Write-off list** — generate the list of items whose service life has expired by a given year.
- **Top 5 longest-serving items** — generate the list of the five items with the longest service life.
- If an item's name consists of two words, they are automatically joined with an underscore (e.g. "Hard_Drive").
- Navigation between pages works both via hyperlinks and programmatically.

## Data model

Each archive record contains:

| Field | Description |
|---|---|
| Name | Item name |
| Inventory number | 4-digit number |
| Lab number | Room the item is assigned to |
| Purchase year | Year of purchase |
| Purchase month | Month of purchase |
| Cost | Cost in monetary units |
| Service life | Service life, in years |
| Quantity | Number of units of the item |

A sample dataset is included in `baza.xml`.

## Architecture (WPF page-based model)

Instead of the usual `Window` as the top-level container, the app uses `System.Windows.Controls.Page`. Each action (loading the archive, viewing it, adding/deleting/editing records, generating reports) is implemented as a separate page.

- **`MainWindow`** — derives from `NavigationWindow` (a window with built-in "back"/"forward" buttons) and hosts the pages.
- **Hyperlinks** (`Hyperlink` with the `NavigateUri` property) are used to move between pages directly from XAML markup.
- **Programmatic navigation** is implemented via `NavigationService.GetNavigationService(this)` and the `Navigate()` method — used, for example, after completing an action (saving a record, deleting one, etc.) to automatically send the user to the appropriate page.

## Repository contents

| File / folder | Purpose |
|---|---|
| `Pages/` | The application's pages (`Page`) — archive loading, viewing, adding/deleting/editing, generating reports |
| `Models/` | Data model classes (archive record) |
| `MainWindow.xaml` / `MainWindow.xaml.cs` | The main window (`NavigationWindow`), hosting the pages |
| `App.xaml` / `App.xaml.cs` | WPF application entry point and setup |
| `AssemblyInfo.cs` | Assembly metadata |
| `Lab7.csproj` | Project file (.NET, WPF); includes a commented-out `Microsoft.Office.Interop.Excel` reference for Excel loading |
| `WpfApp1.csproj` / `WpfApp1.sln` | Visual Studio solution files |
| `baza.xml` | Sample archive in XML format |
| `Microsoft Excel.xlsx` | Sample archive in Excel format |

## Requirements

- Windows
- .NET 6 (`net6.0-windows`) with WPF support (the ".NET desktop development" workload in Visual Studio)
- For Excel loading via COM interop — Microsoft Excel installed, plus uncommenting the `Microsoft.Office.Interop.Excel` reference in `Lab7.csproj` (disabled by default)

## Running the project

1. Open `WpfApp1.sln` in Visual Studio.
2. Build and run the project.
3. From the main menu, choose the data source to load the archive from: Excel, XML (`baza.xml`), or a database.
4. Once the archive is loaded, navigate between pages via the menu/hyperlinks to view the data, add/delete records, edit fields, or generate the reports (write-off list, top 5 by service life).

## Technologies

- C#, WPF (`System.Windows.Controls.Page`, `NavigationWindow`, `Frame`, `NavigationService`)
- XML handling (`System.Xml`)
- Optionally, `Microsoft.Office.Interop.Excel` for loading from Excel
