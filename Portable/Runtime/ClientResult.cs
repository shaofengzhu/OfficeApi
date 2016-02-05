using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OfficeExtension
{
    public class ClientResult<T>: IResultHandler
    {
        private T m_value;

        public T Value
        {
            get
            {
                return m_value;
            }
        }

        public void _HandleResult(JToken value)
        {
            this.m_value = value.Value<T>();
        }
    }
}
