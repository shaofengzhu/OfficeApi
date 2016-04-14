using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
namespace OfficeExtension
{
    public class ObjectPath
    {
        private ObjectPathInfo m_objectPathInfo;
		private bool m_isWriteOperation;
		private ObjectPath m_parentObjectPath;
		private List<ObjectPath> m_argumentObjectPaths;
		private bool m_isCollection;
		private bool m_isInvalidAfterRequest;
		private bool m_isValid;


        public ObjectPath(ObjectPathInfo objectPathInfo, ObjectPath parentObjectPath, bool isCollection, bool isInvalidAfterRequest)
        {
            this.m_objectPathInfo = objectPathInfo;
            this.m_parentObjectPath = parentObjectPath;
            this.m_isWriteOperation = false;
            this.m_isCollection = isCollection;
            this.m_isInvalidAfterRequest = isInvalidAfterRequest;
            this.m_isValid = true;
        }

        public ObjectPathInfo ObjectPathInfo
        {
            get
            {
                return this.m_objectPathInfo;
            }
		}

        public bool IsWriteOperation
        {
            get
            {
                return this.m_isWriteOperation;
            }
            set
            {
                this.m_isWriteOperation = value;
            }
        }
        

        public bool IsCollection
        {
            get
            {
                return this.m_isCollection;
            }
        }

        public bool IsInvalidAfterRequest
        {
            get
            {
                return this.m_isInvalidAfterRequest;
            }
        }

        public ObjectPath ParentObjectPath
        { 
            get
            {
                return this.m_parentObjectPath;
            }
        }

        public List<ObjectPath> ArgumentObjectPaths
        {
            get
            {
                return this.m_argumentObjectPaths;
            }
            set
            {
                this.m_argumentObjectPaths = value;
            }
        }

        public bool IsValid
        {
            get
            {
                return this.m_isValid;
            }
            internal set
            {
                this.m_isValid = value;
            }
        }

        public void UpdateUsingObjectData(JObject value)
        {
			JToken jsonReferenceId = value[Constants.ReferenceId];
			if (!Utility._IsNullOrUndefined(jsonReferenceId))
			{
				string referenceId = jsonReferenceId.Value<string>();
				if (!string.IsNullOrEmpty(referenceId))
				{

					this.m_isInvalidAfterRequest = false;
					this.m_isValid = true;
					this.m_objectPathInfo.ObjectPathType = ObjectPathType.ReferenceId;
					this.m_objectPathInfo.Name = referenceId;
					this.m_objectPathInfo.ArgumentInfo = new ArgumentInfo();
					this.m_parentObjectPath = null;
					this.m_argumentObjectPaths = null;
					return;
				}
			}

			if (this.ParentObjectPath != null && this.ParentObjectPath.IsCollection)
            {
                JToken jsonId = value[Constants.Id];
                if (Utility._IsNullOrUndefined(jsonId))
                {
					jsonId = value[Constants.IdPrivate];
                }

                if (!Utility._IsNullOrUndefined(jsonId))
                {
                    this.m_isInvalidAfterRequest = false;
                    this.m_isValid = true;
                    this.m_objectPathInfo.ObjectPathType = ObjectPathType.Indexer;
                    this.m_objectPathInfo.Name = "";
                    this.m_objectPathInfo.ArgumentInfo = new ArgumentInfo();
                    this.m_objectPathInfo.ArgumentInfo.Arguments = new object[] { jsonId.ToObject<object>() };
                    this.m_argumentObjectPaths = null;
                    return;
                }
            }
        }
    }
}
