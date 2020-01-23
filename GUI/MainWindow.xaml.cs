using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ExcelReaderConsole;
using ExcelReaderConsole.Models;
using GUI.Models;

namespace GUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ExcelReader excelReader;
        private DocumentsStorage ds;

        private readonly MainWindowModel model;

        public MainWindow()
        {
            InitializeComponent();
            model = new MainWindowModel();
            DataContext = model;
        }

        private void InitColumnToTemplateGridItem(IEnumerable<string> newColumnNames)
        {
            TemplateDataGrid.Columns.Clear();
            foreach (string name in newColumnNames)
            {
                TemplateDataGrid.Columns.Add(new DataGridTextColumn()
                {
                    Binding = new Binding($"TemplateDataGridRows[{name}]"),
                    Header = name
                });
            }
        }

        private void MenuItemLoadTemplate_OnClick(object sender, RoutedEventArgs e)
        {
            excelReader = new ExcelReader(AppSettings.Instance.GetExcelTemplateFilePath());
            ds = new DocumentsStorage();
            excelReader.ReadData(ds);

            var attributes= ds.GetUsedDocumentAttributes();
            model.TemplateDataGridModel.TemplateDataGridColumns.AddRange(attributes.Select(item => item.Name));
            InitColumnToTemplateGridItem(model.TemplateDataGridModel.TemplateDataGridColumns);
            var documents = ds.GetDocuments();
            foreach (var document in documents)
            {
                var tgi = new TemplateGridItem();
                model.TemplateDataGridModel.TemplateGridItems.Add(tgi);
                foreach (var attribute in attributes)
                {
                    var documentAttrValue = document.GetValue(attribute.Identifier).Value;
                    tgi.TemplateDataGridRows.Add(attribute.Name, documentAttrValue);
                }
            }
            TemplateDataGrid.ItemsSource = model.TemplateDataGridModel.TemplateGridItems;
        }

        private void MenuItemSettings_OnClick(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void MenuItemExit_OnClick(object sender, RoutedEventArgs e)
        {
            var messageBoxResult = MessageBox.Show("Are you sure want to close application?", "Exit",
                MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (messageBoxResult == MessageBoxResult.Yes)
            {
                this.Close();
            }
        }
    }
}
