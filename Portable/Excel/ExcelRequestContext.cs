using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeExtension;

namespace Microsoft.ExcelServices
{
    public class ExcelRequestContext: ClientRequestContext
    {
        private Workbook m_workbook;
        public ExcelRequestContext(string url)
            : base(url)
        {
            m_workbook = new Workbook(this, ObjectPathFactory._CreateGlobalObjectObjectPath(this));
            this._RootObject = m_workbook;            
        }

        public Workbook Workbook
        {
            get
            {
                return m_workbook;
            }
        }
    }
}
