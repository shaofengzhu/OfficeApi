using System;
using System.Collections.Generic;
using System.Net;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeExtension
{
    public class RequestExecutorResponseMessage
    {
        private Dictionary<string, string> m_headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public HttpStatusCode StatusCode
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
