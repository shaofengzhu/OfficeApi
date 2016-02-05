using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeExtension
{
    public class HttpRequestExecutor: IRequestExecutor
    {
        public HttpRequestExecutor()
        {

        }
        public Task<RequestExecutorResponseMessage> Execute(RequestExecutorRequestMessage request)
        {
            return null;
        }
    }
}
