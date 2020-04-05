using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

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
            var app = new Excel.Application();
            foreach (string file in xlsFiles)
            {
                Excel.Workbook book = app.Workbooks.Open(file);
                foreach (Excel.Worksheet worksheet in book.Worksheets)
                {
                    var range = worksheet.UsedRange;
                    var result = range.Find(TextBoxSearch.Text);
                    if (result != null)
                    {
                        listBoxSearchResult.Items.Add(file);
                        break;
                    }
                }
            }
        }
    }
}
