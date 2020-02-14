using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReaderConsole.Models
{
    public class DocumentEdition : Document
    {
        public string MainEditionIdentifier { get; private set; }
        public DocumentEdition(string identifier, string mainEditionIdentifier, IEnumerable<DocumentAttribute> attributes) : base(identifier, attributes)
        {
            MainEditionIdentifier = mainEditionIdentifier;
        }
    }
}
