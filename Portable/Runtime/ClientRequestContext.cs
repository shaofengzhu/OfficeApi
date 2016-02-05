using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
namespace OfficeExtension
{
    public class ClientRequestContext
    {
        private int m_nextId;
		private ClientRequest m_pendingRequest;
		private string m_url;
		private TrackedObjects m_trackedObjects;

		private IRequestExecutor m_requestExecutor;
		private ClientObject m_rootObject;
		private bool m_processingResult;


        public ClientRequestContext(string url)
        {
            this.m_url = url;
            if (string.IsNullOrEmpty(this.m_url))
            {
                this.m_url = Constants.LocalDocument;
            }

            this.m_processingResult = false;
            this.m_requestExecutor = new HttpRequestExecutor();
        }

        public ClientRequest _PendingRequest
        {
            get
            {
                if (this.m_pendingRequest == null)
                {
                    this.m_pendingRequest = new ClientRequest(this);
                }

                return this.m_pendingRequest;
            }
        }

        internal bool ProcessingResult
        {
            get
            {
                return this.m_processingResult;
            }
        }

        public TrackedObjects TrackedObjects
        {
            get
            {
                if (this.m_trackedObjects == null)
                {
                    this.m_trackedObjects = new TrackedObjects(this);
                }
                return this.m_trackedObjects;
            }
        }

        public ClientObject _RootObject
        {
            get
            {
                return m_rootObject;
            }
            protected set
            {
                m_rootObject = value;
            }
        }

        public void load(ClientObject clientObj, LoadOption loadOption)
        {
			Utility.ValidateContext(this, clientObj);

            QueryInfo queryOption = new QueryInfo();

            if (!string.IsNullOrEmpty(loadOption.Select))
            {
                queryOption.Select = this.ParseSelectExpand(loadOption.Select);
            }

            if (!string.IsNullOrEmpty(loadOption.Expand))
            {
                queryOption.Expand = this.ParseSelectExpand(loadOption.Expand);
            }

            queryOption.Skip = loadOption.Skip;
            queryOption.Top = loadOption.Top;

            var action = ActionFactory.CreateQueryAction(this, clientObj, queryOption);
			this._PendingRequest.AddActionResultHandler(action, clientObj);
        }


        public void trace(string message)
        {
			ActionFactory.CreateTraceAction(this, message);
        }

        private string[] ParseSelectExpand(string select)
        {
            List<string> ret = new List<string>();
            foreach (string str in select.Split(','))
            {
                string tmp = str.Trim();
                if (tmp.Length > 0)
                {
                    ret.Add(tmp);
                }
            }

            return ret.ToArray();
		}

		public async void Sync()
        { 
			ClientRequest req = this.m_pendingRequest;
            if (req == null)
            {
                return;
            }

			// If there are no actions to dispatch, short-circuit without sending an empty request to the server
			if (!req.HasActions)
            {
                return;
            }

			this.m_pendingRequest = null;
            RequestMessageBody msgBody = req.BuildRequestMessageBody();
            ClientRequestFlags requestFlags = req.Flags;

            RequestExecutorRequestMessage requestExecutorRequestMessage = new RequestExecutorRequestMessage();
            requestExecutorRequestMessage.Url = this.m_url;
            requestExecutorRequestMessage.Body = Utility.ToJsonString(msgBody);

            req.InvalidatePendingInvalidObjectPaths();

            RequestExecutorResponseMessage response = await this.m_requestExecutor.Execute(requestExecutorRequestMessage);
            JToken json = Utility.ToJsonObject(response.Body);
            this.m_processingResult = true;
            try
            {
                req.ProcessResponse(json);
            }
            finally
            {
                this.m_processingResult = false;
            }
        }

        public int _NextId()
        {
			return ++this.m_nextId;
        }
    }
}
