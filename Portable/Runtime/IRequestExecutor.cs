using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeExtension
{
    public interface IRequestExecutor
    {
        Task<RequestExecutorResponseMessage> Execute(RequestExecutorRequestMessage request);
    }
}
