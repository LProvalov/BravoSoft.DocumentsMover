using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;

namespace GUI.Models
{
    public class TemplateGridItem
    {
        public Dictionary<string, object> TemplateDataGridRows { get; set; } = new Dictionary<string, object>();
        public bool TextFileExists { get; set; } = false;
        public bool TextPDFFileExists { get; set; } = false;
        public bool ScanFileExists { get; set; } = false;
        public bool AttachmentFiles { get; set; } = false;
        public bool CopiedTextFile { get; set; } = false;
        public bool CopiedPDFFile { get; set; } = false;
        public bool CopiedScanFile { get; set; } = false;
        public bool CopiedAttachmentFiles { get; set; } = false;

    }
}
