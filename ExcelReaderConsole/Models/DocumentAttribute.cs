using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReaderConsole.Models
{
    public class DocumentAttribute
    {
        public DocumentAttribute()
        {
            identifier = string.Empty;
            name = string.Empty;
        }

        public DocumentAttribute(string identifier, string name)
        {
            this.identifier = identifier;
            this.name = name;
        }

        private string identifier;
        private string name;

        public string Identifier {
            get
            {
                return this.identifier;
            }
            set
            {
                this.identifier = value;
            }
        }

        public string Name
        {
            get
            {
                return this.name;
            }
            set
            {
                this.name = value;
            }
        }
    }
}
