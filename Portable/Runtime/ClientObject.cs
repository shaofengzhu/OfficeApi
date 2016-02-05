using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OfficeExtension
{
    public class ClientObject : IResultHandler
    {
        private ClientRequestContext m_context;
        private ObjectPath m_objectPath;

        public ClientObject(ClientRequestContext context, ObjectPath objectPath)
        {
            Utility.CheckArgumentNull(context, "context");
            this.m_context = context;
            this.m_objectPath = objectPath;
            if (this.m_objectPath != null)
            {
                // If object is being created during a normal API flow (and NOT as part of processing load results),
                // create an instantiation action and call keepReference, if applicable
                if (!context.ProcessingResult)
                {
                    ActionFactory.CreateInstantiateAction(context, this);
                }
            }
        }

        public ClientRequestContext Context
        {
            get
            {
                return this.m_context;
            }
        }

        public ObjectPath _ObjectPath
        {
            get
            {
                return this.m_objectPath;
            }
            set
            {
                this.m_objectPath = value;
            }
        }


        public virtual void _HandleResult(JToken value)
        {
        }
    }
}
