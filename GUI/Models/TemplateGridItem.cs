using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace GUI.Models
{
    public class TemplateGridItem
    {
        public Dictionary<string, object> TemplateDataGridRows { get; set; } = new Dictionary<string, object>();
    }
}
