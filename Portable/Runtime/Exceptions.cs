using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeExtension
{
    public class ClientRuntimeException: Exception
    {
        private string m_code;
        private string m_location;

        public ClientRuntimeException(string code, string message, string location)
            : base(message)
        {
            this.m_code = code;
            this.m_location = location;
        }

        public string Code
        {
            get
            {
                return m_code;
            }
        }

        public string Location
        {
            get
            {
                return m_location;
            }
        }
        
    }
}
