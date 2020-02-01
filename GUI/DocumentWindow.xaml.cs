using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using ExcelReaderConsole.Models;
using GUI.Models;

namespace GUI
{
    /// <summary>
    /// Interaction logic for DocumentWindow.xaml
    /// </summary>
    public partial class DocumentWindow : Window
    {
        public DocumentWindow(DocumentWindowModel model)
        {
            InitializeComponent();
            DataContext = model;
            //listViewLogs.ItemsSource = model.LogMessages;
        }
    }
}
