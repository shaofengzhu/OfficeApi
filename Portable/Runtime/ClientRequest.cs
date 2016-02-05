using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
namespace OfficeExtension
{
    public class ClientRequest
    {
        private ClientRequestFlags m_flags;
        private ClientRequestContext m_context;
        private List<Action> m_actions;
        private Dictionary<int, IResultHandler> m_actionResultHandler;
        private Dictionary<int, ObjectPath> m_referencedObjectPaths;
        private Dictionary<int, string> m_traceInfos;


        internal ClientRequest(ClientRequestContext context)
        {
            this.m_context = context;
            this.m_actions = new List<Action>();
            this.m_actionResultHandler = new Dictionary<int, IResultHandler>();
            this.m_referencedObjectPaths = new Dictionary<int, ObjectPath>();
            this.m_flags = ClientRequestFlags.None;
            this.m_traceInfos = new Dictionary<int, string>();
        }

        public ClientRequestFlags Flags
        {
            get
            {
                return this.m_flags;
            }
        }

        public Dictionary<int, string> TraceInfos
        {
            get
            {
                return this.m_traceInfos;
            }
        }


        public void AddAction(Action action)
        {
            if (action.IsWriteOperation)
            {
                this.m_flags = this.m_flags | ClientRequestFlags.WriteOperation;
            }
            this.m_actions.Add(action);
        }

        public bool HasActions
        {
            get
            {
                return this.m_actions.Count > 0;
            }
        }


        public void AddTrace(int actionId, string message)
        {
            this.m_traceInfos[actionId] = message;
        }


        public void AddReferencedObjectPath(ObjectPath objectPath)
        {
            if (this.m_referencedObjectPaths.ContainsKey(objectPath.ObjectPathInfo.Id))
            {
                return;
            }

            if (!objectPath.IsValid)
            {
                throw Utility.CreateInvalidObjectPathException(objectPath);
            }

            while (objectPath != null)
            {
                if (objectPath.IsWriteOperation)
                {
                    this.m_flags = this.m_flags | ClientRequestFlags.WriteOperation;
                }

                this.m_referencedObjectPaths[objectPath.ObjectPathInfo.Id] = objectPath;

                if (objectPath.ObjectPathInfo.ObjectPathType == ObjectPathType.Method)
                {
                    this.AddReferencedObjectPaths(objectPath.ArgumentObjectPaths);
                }

                objectPath = objectPath.ParentObjectPath;
            }
        }


        internal void AddReferencedObjectPaths(IEnumerable<ObjectPath> objectPaths)
        {
            if (objectPaths != null)
            {
                foreach (ObjectPath objectPath in objectPaths)
                {
                    this.AddReferencedObjectPath(objectPath);
                }
            }
        }


        public void AddActionResultHandler(Action action, IResultHandler resultHandler)
        {
            this.m_actionResultHandler[action.ActionInfo.Id] = resultHandler;
        }


        internal RequestMessageBody BuildRequestMessageBody()
        {
            Dictionary<int, ObjectPathInfo> objectPaths = new Dictionary<int, ObjectPathInfo>();
            foreach (var pair in this.m_referencedObjectPaths)
            {
                objectPaths[pair.Key] = pair.Value.ObjectPathInfo;
            }

            List<ActionInfo> actions = new List<ActionInfo>();
            foreach (Action action in this.m_actions)
            {
                actions.Add(action.ActionInfo);
            }

            RequestMessageBody ret = new RequestMessageBody();
            ret.Actions = actions;
            ret.ObjectPaths = objectPaths;

            return ret;
        }

        internal void ProcessResponse(JToken json)
        {
            if (json != null && 
                json.Type == JTokenType.Object && 
                json["Results"] != null)
            {
                JArray results = json["Results"] as JArray;
                for (var i = 0; i < results.Count; i++)
                {
                    JToken actionResult = results[i];
                    int actionId = actionResult.Value<int>("ActionId");
                    var handler = this.m_actionResultHandler[actionId];
                    if (handler != null)
                    {
                        JToken actionValue = actionResult["Value"];
                        handler._HandleResult(actionValue);
                    }
                }
            }
        }


        internal void InvalidatePendingInvalidObjectPaths()
        {
            foreach (var pair in this.m_referencedObjectPaths)
            {
                if (pair.Value.IsInvalidAfterRequest)
                {
                    pair.Value.IsValid = false;
                }
            }
        }
    }
}
