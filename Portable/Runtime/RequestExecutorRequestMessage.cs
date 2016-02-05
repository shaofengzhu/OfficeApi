using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeExtension
{
    public sealed class RequestExecutorRequestMessage
    {
        private Dictionary<string, string> m_headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string Url
        {
            get;
            set;
        }

        public IDictionary<string, string> Headers
        {
            get
            {
                return m_headers;
            }
        }

        public string Body
        {
            get;
            set;
        }
    }
}
