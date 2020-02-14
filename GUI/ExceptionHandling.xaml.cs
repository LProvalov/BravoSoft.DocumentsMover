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
using GUI.Models;

namespace GUI
{
    /// <summary>
    /// Interaction logic for ExceptionHandling.xaml
    /// </summary>
    public partial class ExceptionHandling : Window
    {
        private static ExceptionHandling _instance;
        private ExceptionHandlingWindowModel model = new ExceptionHandlingWindowModel();

        public ExceptionHandlingWindowModel Model
        {
            get { return model; }
        }

        public static ExceptionHandling Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ExceptionHandling();
                }
                return _instance;
            }
        }

        public ExceptionHandling()
        {
            InitializeComponent();
            this.DataContext = model;
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Hidden;
        }

        private void FillStringBuilderByException(StringBuilder sb, Exception ex)
        {
            if (sb == null || ex == null) return;
            sb.AppendLine(ex.Message);
            if (ex.InnerException != null)
            {
                FillStringBuilderByException(sb, ex.InnerException);
            }
        }
        private void ListViewItem_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var item = sender as ListViewItem;
            if (item != null && item.IsSelected)
            {
                StringBuilder sb = new StringBuilder();
                if (item.DataContext is Exception)
                {
                    Exception ex = item.DataContext as Exception;
                    FillStringBuilderByException(sb, ex);
                }
                textBlockExceptionDescription.Text = sb.ToString();
            }
        }
    }
}
