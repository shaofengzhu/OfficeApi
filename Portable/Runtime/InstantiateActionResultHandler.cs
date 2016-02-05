using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OfficeExtension
{
    public class InstantiateActionResultHandler: IResultHandler
    {

        private ClientObject m_clientObject;

        public InstantiateActionResultHandler(ClientObject clientObject)
        {
            this.m_clientObject = clientObject;
        }


        public void _HandleResult(JToken value)
        {
			Utility.FixObjectPathIfNecessary(this.m_clientObject, value);
			if (value != null &&
				!OfficeExtension.Utility.IsNullOrUndefined(value[Constants.ReferenceId]))
            {
                ITrackedObject trackedObject = this.m_clientObject as ITrackedObject;
                if (trackedObject != null)
                {
                    trackedObject._ReferenceId = value[Constants.ReferenceId].Value<string>();
                }
            }
        }
    }
}
