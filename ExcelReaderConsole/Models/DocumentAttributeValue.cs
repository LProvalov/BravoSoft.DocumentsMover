using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReaderConsole.Models
{
    public class DocumentAttributeValue
    {
        private object value = null;
        private Type valueType;

        public DocumentAttributeValue(object value, Type valueType)
        {
            if (value != null && (value.GetType() == valueType))
            {
                this.value = value;
                this.valueType = valueType;
            }
        }

        public Type ValueType { get { return this.valueType; } }
        public object Value { get { return this.value; } }
    }
}
