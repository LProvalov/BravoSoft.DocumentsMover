using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using GUI.Extensions;

namespace GUI.Models
{
    public class MainWindowModel
    {
        public TemplateDataGridModel TemplateDataGridModel { get; } = new TemplateDataGridModel();
    }
}
