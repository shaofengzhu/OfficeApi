using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
namespace OfficeExtension
{
    public static class ObjectPathFactory
    {
        public static ObjectPath CreateGlobalObjectObjectPath(ClientRequestContext context)
        {
            ObjectPathInfo objectPathInfo = new ObjectPathInfo()
            {
                Id = context._NextId(),
                ObjectPathType = ObjectPathType.GlobalObject,
                Name = ""
            };

			return new ObjectPath(objectPathInfo, null, false /*isCollection*/, false /*isInvalidAfterRequest*/);
        }

        public static ObjectPath CreateNewObjectObjectPath(ClientRequestContext context, string typeName, bool isCollection)
        {
            ObjectPathInfo objectPathInfo = new ObjectPathInfo()
            {
                Id = context._NextId(),
                ObjectPathType = ObjectPathType.NewObject,
                Name = typeName
            };

			return new ObjectPath(objectPathInfo, null, isCollection, false /*isInvalidAfterRequest*/);
		}

		public static ObjectPath CreatePropertyObjectPath(ClientRequestContext context, ClientObject parent, string propertyName, bool isCollection, bool isInvalidAfterRequest)
        {
            ObjectPathInfo objectPathInfo = new ObjectPathInfo()
				{
					Id = context._NextId(),
					ObjectPathType = ObjectPathType.Property,
					Name = propertyName,
					ParentObjectPathId = parent._ObjectPath.ObjectPathInfo.Id,
				};

            return new ObjectPath(objectPathInfo, parent._ObjectPath, isCollection, isInvalidAfterRequest);
		}

		public static ObjectPath CreateIndexerObjectPath(ClientRequestContext context, ClientObject parent, object[] args)
        { 
			ObjectPathInfo objectPathInfo = new ObjectPathInfo()
				{
					Id = context._NextId(),
					ObjectPathType = ObjectPathType.Indexer,
					Name = "",
					ParentObjectPathId = parent._ObjectPath.ObjectPathInfo.Id,
					ArgumentInfo = new ArgumentInfo()
				};

			objectPathInfo.ArgumentInfo.Arguments = args;
			return new ObjectPath(objectPathInfo, parent._ObjectPath, false /*isCollection*/, false /*isInvalidAfterRequest*/);
		}

		public static ObjectPath CreateIndexerObjectPathUsingParentPath(ClientRequestContext context, ObjectPath parentObjectPath, object[] args)
        {
            ObjectPathInfo objectPathInfo = new ObjectPathInfo()
				{
					Id = context._NextId(),
					ObjectPathType = ObjectPathType.Indexer,
					Name = "",
					ParentObjectPathId = parentObjectPath.ObjectPathInfo.Id,
					ArgumentInfo = new ArgumentInfo()
				};
			objectPathInfo.ArgumentInfo.Arguments = args;
			return new ObjectPath(objectPathInfo, parentObjectPath, false /*isCollection*/, false /*isInvalidAfterRequest*/);
		}

		public static ObjectPath CreateMethodObjectPath(ClientRequestContext context, ClientObject parent, string methodName, OperationType operationType, object[] args, bool isCollection, bool isInvalidAfterRequest)
        {
			ObjectPathInfo objectPathInfo = new ObjectPathInfo()
				{
					Id = context._NextId(),
					ObjectPathType = ObjectPathType.Method,
					Name = methodName,
					ParentObjectPathId = parent._ObjectPath.ObjectPathInfo.Id,
					ArgumentInfo = new ArgumentInfo()
				};
			List<ObjectPath> argumentObjectPaths = Utility.SetMethodArguments(context, objectPathInfo.ArgumentInfo, args);
			ObjectPath ret = new ObjectPath(objectPathInfo, parent._ObjectPath, isCollection, isInvalidAfterRequest);
            ret.ArgumentObjectPaths = argumentObjectPaths;
			ret.IsWriteOperation = (operationType != OperationType.Read);
			return ret;
		}

		public static ObjectPath CreateChildItemObjectPathUsingIndexerOrGetItemAt(bool hasIndexerMethod, ClientRequestContext context, ClientObject parent,  JObject childItem, int index)
        {
			var id = childItem[Constants.Id];
			if (Utility.IsNullOrUndefined(id))
            {
				id = childItem[Constants.IdPrivate];
			}

			if (hasIndexerMethod && !Utility.IsNullOrUndefined(id))
            {
				return ObjectPathFactory.CreateChildItemObjectPathUsingIndexer(context, parent, childItem);
			}
			else
            {
				return ObjectPathFactory.CreateChildItemObjectPathUsingGetItemAt(context, parent, childItem, index);
			}
		}

		public static ObjectPath CreateChildItemObjectPathUsingIndexer(ClientRequestContext context, ClientObject parent, JObject childItem)
        {
			var id = childItem[Constants.Id];
			if (Utility.IsNullOrUndefined(id)) {
				id = childItem[Constants.IdPrivate];
			}

			ObjectPathInfo objectPathInfo = new ObjectPathInfo()
				{
					Id = context._NextId(),
					ObjectPathType = ObjectPathType.Indexer,
					Name = "",
					ParentObjectPathId = parent._ObjectPath.ObjectPathInfo.Id,
					ArgumentInfo = new ArgumentInfo()
				};
			objectPathInfo.ArgumentInfo.Arguments = new object[] { id };
			return new ObjectPath(objectPathInfo, parent._ObjectPath, false /*isCollection*/, false /*isInvalidAfterRequest*/);
		}

		public static ObjectPath CreateChildItemObjectPathUsingGetItemAt(ClientRequestContext context, ClientObject parent, JObject childItem, int index)
        {
			JToken indexFromServer = childItem[Constants.Index];
			if (!Utility.IsNullOrUndefined(indexFromServer))
            {
                index = indexFromServer.Value<int>();
			}

            ObjectPathInfo objectPathInfo = new ObjectPathInfo()
				{
					Id = context._NextId(),
					ObjectPathType = ObjectPathType.Method,
					Name = Constants.GetItemAt,
					ParentObjectPathId = parent._ObjectPath.ObjectPathInfo.Id,
					ArgumentInfo = new ArgumentInfo()
				};
			objectPathInfo.ArgumentInfo.Arguments = new object[] { index };
			return new ObjectPath(objectPathInfo, parent._ObjectPath, false /*isCollection*/, false /*isInvalidAfterRequest*/);
		}
	}
}
