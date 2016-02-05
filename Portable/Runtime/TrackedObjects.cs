using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeExtension
{
    public class TrackedObjects
    {
        private ClientRequestContext m_context;

        // Objects that need to clean up after, if in auto-cleanup mode
        private Dictionary<int, ClientObject> _autoCleanupList = new Dictionary<int, ClientObject>();


        internal TrackedObjects(ClientRequestContext context)
        {
            this.m_context = context;
        }


        public void Add(ClientObject clientObject)
        {
            ITrackedObject trackedObj = clientObject as ITrackedObject;
            if (trackedObj != null)
            {
                if (string.IsNullOrEmpty(trackedObj._ReferenceId))
                {
                    trackedObj._KeepReference();
                    ActionFactory.CreateInstantiateAction(this.m_context, clientObject);
                }
            }
        }

        public void Add(IEnumerable<ClientObject> clientObjects)
        {
            if (clientObjects != null)
            {
                foreach (ClientObject clientObject in clientObjects)
                {
                    this.Add(clientObject);
                }
            }
        }

        public void Remove(ClientObject clientObject)
        {
            if (this.m_context._RootObject == null)
            {
                return;
            }

            ITrackedObject trackedObject = clientObject as ITrackedObject;
            if (trackedObject == null)
            {
                return;
            }

            string referenceId = trackedObject._ReferenceId;
            if (string.IsNullOrEmpty(referenceId))
            {
                return;
            }

            ActionFactory.CreateMethodAction(this.m_context, this.m_context._RootObject, "_RemoveReference", OperationType.Read, new object[] { referenceId });
        }

        public void Remove(IEnumerable<ClientObject> clientObjects)
        {
            if (clientObjects != null)
            {
                foreach (ClientObject clientObject in clientObjects)
                {
                    Remove(clientObject);
                }
            }
        }
    }
}
