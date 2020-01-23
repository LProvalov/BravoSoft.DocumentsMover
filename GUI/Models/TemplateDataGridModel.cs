using System;
using System.Collections.Generic;
using System.ComponentModel;
using GUI.Extensions;

namespace GUI.Models
{
    public class TemplateDataGridModel
    {
        public readonly ObservableRangeCollection<string> TemplateDataGridColumns = new ObservableRangeCollection<string>();
        public readonly List<TemplateGridItem> TemplateGridItems = new List<TemplateGridItem>();
    }
}
