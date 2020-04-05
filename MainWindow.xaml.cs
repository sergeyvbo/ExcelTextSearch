using Microsoft.Win32;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using OfficeOpenXml;

namespace ExcelTextSearch
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void ButtonOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                FileName = "Select folder",
                ValidateNames = false,
                CheckFileExists = false,
                CheckPathExists = true
            };
            var result = dlg.ShowDialog();
            if (result == true)
            {
                TextBoxStartPath.Text = System.IO.Path.GetDirectoryName(dlg.FileName);
                UpdateTreeView(TextBoxStartPath.Text);
            }

        }

        private void UpdateTreeView(string path)
        {
            TreeViewPath.Items.Clear();
            var treeViewItem = new TreeViewItem();
            treeViewItem.Header = path;
            TreeViewPath.Items.Add(treeViewItem);
            LoadFiles(path, treeViewItem);
            LoadSubfoldersRec(path, treeViewItem);
        }

        private void LoadSubfoldersRec(string path, TreeViewItem parent)
        {
            string[] subfolders = Directory.GetDirectories(path);
            foreach (string subfolder in subfolders)
            {
                var direcoryInfo = new DirectoryInfo(subfolder);
                var treeViewItem = new TreeViewItem();
                treeViewItem.Header = direcoryInfo.Name;
                treeViewItem.Tag = direcoryInfo.FullName;
                parent.Items.Add(treeViewItem);
                LoadFiles(subfolder, treeViewItem);
                LoadSubfoldersRec(subfolder, treeViewItem);

            }

        }

        private void LoadFiles(string path, TreeViewItem parent)
        {
            string[] files = Directory.GetFiles(path);
            foreach (string file in files)
            {
                var fileInfo = new FileInfo(file);
                var treeViewItem = new TreeViewItem();
                treeViewItem.Header = fileInfo.Name;
                treeViewItem.Tag = fileInfo.FullName;
                parent.Items.Add(treeViewItem);
            }
        }

        private void ButtonSearch_Click(object sender, RoutedEventArgs e)
        {
            string[] xlsFiles = Directory.GetFiles(TextBoxStartPath.Text, "*.xlsx", SearchOption.AllDirectories);
            foreach (var filename in xlsFiles)
            {
                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(filename)))
                {
                    var myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                    var totalRows = myWorksheet.Dimension.End.Row;
                    var totalColumns = myWorksheet.Dimension.End.Column;

                    for (var rowNum = 1; rowNum <= totalRows; rowNum++) //select starting row here
                    {
                        var row = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value?.ToString())
                            .Where(c => c != null && c.Contains(TextBoxSearch.Text));
                        if (!row.Any()) continue;
                        listBoxSearchResult.Items.Add(filename);
                        break;

                    }
                }
            }
        }

        private void listBoxSearchResult_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (listBoxSearchResult.SelectedItem.ToString() == string.Empty) return;
            System.Diagnostics.Process.Start(listBoxSearchResult.SelectedItem.ToString());
        }
    }
}
