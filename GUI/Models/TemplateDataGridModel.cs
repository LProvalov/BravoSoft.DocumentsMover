using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using GUI.Extensions;

namespace GUI.Models
{
    public class TemplateDataGridModel
    {
        public static class AdditionalAttributes
        {
            public static readonly string TextFileAttribute = "Текстовый файл";
            public static readonly string ScanCopyAttribute = "Загрузить как скан-копию";
            public static readonly string TextPDFAttribute = "Текстовый PDF";
            public static readonly string AttachmentsAttribute = "Вложения";

        }

        public readonly ObservableRangeCollection<string> TemplateDataGridColumns = new ObservableRangeCollection<string>();
        public readonly List<TemplateGridItem> TemplateGridItems = new List<TemplateGridItem>();
        public bool IsWarning { get; set; } = true;
    }
}
