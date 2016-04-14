using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;


namespace OfficeExtension
{
    public static class Utility
    {
        public static void CheckArgumentNull(object value, string name)
        {
			if (value == null)
            {
                throw new ArgumentNullException(name);
			}
		}
        public static bool _IsUndefined(JToken value)
        {
            if (value == null ||
                value.Type == JTokenType.Undefined)
            {
                return true;
            }

            return false;
        }

        public static bool _IsNullOrUndefined(JToken value)
        {
            if (value == null ||
                value.Type == JTokenType.None ||
                value.Type == JTokenType.Undefined ||
                value.Type == JTokenType.Null)
            {
                return true;
            }

            return false;
        }

		internal static string CombineUrl(string baseUrl, string relativeUrl)
		{
			if (!baseUrl.EndsWith("/", StringComparison.Ordinal))
			{
				baseUrl += "/";
			}

			if (relativeUrl.StartsWith("/", StringComparison.Ordinal))
			{
				relativeUrl = relativeUrl.Substring(1);
			}

			return baseUrl + relativeUrl;
		}

		internal static List<ObjectPath> SetMethodArguments(ClientRequestContext context, ArgumentInfo argumentInfo, object[] args)
        {
			if (args == null)
            {
				return null;
			}

			var referencedObjectPaths = new List<ObjectPath>();
            List<object> referencedObjectPathIds = new List<object>();
            bool hasOne = Utility.CollectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
            argumentInfo.Arguments = args;
			if (hasOne)
            {
				argumentInfo.ReferencedObjectPathIds = referencedObjectPathIds.ToArray();
				return referencedObjectPaths;
			}

			return null;
		}

		private static bool CollectObjectPathInfos(
            ClientRequestContext context,
            object[] args,
            List<ObjectPath> referencedObjectPaths,
            List<object> referencedObjectPathIds)
        {
			bool hasOne = false;
			for (int i = 0; i < args.Length; i++)
            {
				if (args[i] is ClientObject) {
                    ClientObject clientObject = (ClientObject)args[i];
					Utility.ValidateContext(context, clientObject);
					args[i] = clientObject._ObjectPath.ObjectPathInfo.Id;
					referencedObjectPathIds.Add(clientObject._ObjectPath.ObjectPathInfo.Id);
					referencedObjectPaths.Add(clientObject._ObjectPath);
					hasOne = true;
				}
				else if (args[i] != null && args[i].GetType().IsArray)
                {
					var childArrayObjectPathIds = new List<object>();
                    var childArrayHasOne = Utility.CollectObjectPathInfos(context, (object[])args[i], referencedObjectPaths, childArrayObjectPathIds);
					if (childArrayHasOne)
                    {
						referencedObjectPathIds.Add(childArrayObjectPathIds.ToArray());
						hasOne = true;
					}
					else
                    {
						referencedObjectPathIds.Add(0);
					}
				}
				else
                {
					referencedObjectPathIds.Add(0);
				}
			}

			return hasOne;
		}

		public static void _FixObjectPathIfNecessary(ClientObject clientObject, JToken value)
        {
			if (clientObject != null && 
                clientObject._ObjectPath != null && 
                value != null &&
                value.Type == JTokenType.Object)
            {
				clientObject._ObjectPath.UpdateUsingObjectData((JObject)value);
			}
		}

		internal static void ValidateObjectPath(ClientObject clientObject)
        {
			ObjectPath objectPath  = clientObject._ObjectPath;
			while (objectPath != null)
            {
				if (!objectPath.IsValid)
                {
                    throw CreateInvalidObjectPathException(objectPath);
                }

                objectPath = objectPath.ParentObjectPath;
			}
		}

		internal static void ValidateReferencedObjectPaths(IEnumerable<ObjectPath> objectPaths)
        {
			if (objectPaths != null)
            {
				foreach (ObjectPath item in objectPaths)
                {
                    ObjectPath objectPath = item;
					while (objectPath != null)
                    {
						if (!objectPath.IsValid)
                        {
                            throw CreateInvalidObjectPathException(objectPath);
						}

						objectPath = objectPath.ParentObjectPath;
					}
				}
			}
		}

		internal static void ValidateContext(ClientRequestContext context, ClientObject obj)
        {
			if (obj != null && obj.Context != context)
            {
                throw CreateRuntimeError(
                    ErrorCodes.GeneralException,
                    _GetResourceString(ResourceStrings.InvalidObjectPath),
                    null);
			}
        }
        internal static Exception CreateRuntimeError(string code, string message)
        {
            return CreateRuntimeError(code, message, null);
        }

        internal static Exception CreateRuntimeError(string code, string message, string location)
        {
            ClientRuntimeException ex = new ClientRuntimeException(code, message, location);
            return ex;
		}

        internal static Exception CreateInvalidObjectPathException(ObjectPath objectPath)
        {
            string pathExpression = Utility.GetObjectPathExpression(objectPath);
            return CreateRuntimeError(
                ErrorCodes.GeneralException,
                _GetResourceString(ResourceStrings.InvalidObjectPath, pathExpression));
        }

        public static string _GetResourceString(string resourceId, params object[] args)
        {
            return resourceId;
        }

		public static void _ThrowIfNotLoaded(ClientObject clientObject, string propertyName, object fieldValue)
        {
			if (!clientObject.LoadedPropertyNames.Contains(propertyName))
			{
				throw CreateRuntimeError(
					ErrorCodes.GeneralException,
					_GetResourceString(ResourceStrings.PropertyNotLoaded, propertyName));
			}
		}

		internal static string GetObjectPathExpression(ObjectPath objectPath)
        {
			string ret = "";
			while (objectPath != null) {
				switch (objectPath.ObjectPathInfo.ObjectPathType) {
					case ObjectPathType.GlobalObject:
						ret = "";
						break;
					case ObjectPathType.NewObject:
						ret = "new()" + (ret.Length > 0 ? "." : "") + ret;
						break;
					case ObjectPathType.Method:
						ret = Utility.NormalizeName(objectPath.ObjectPathInfo.Name) + "()" + (ret.Length > 0 ? "." : "") + ret;
						break;
					case ObjectPathType.Property:
						ret = Utility.NormalizeName(objectPath.ObjectPathInfo.Name) + (ret.Length > 0 ? "." : "") + ret;
						break;
					case ObjectPathType.Indexer:
						ret = "getItem()" + (ret.Length > 0 ? "." : "") + ret;
						break;
					case ObjectPathType.ReferenceId:
						ret = "_reference()" + (ret.Length > 0 ? "." : "") + ret;
						break;
				}

				objectPath = objectPath.ParentObjectPath;
			}

			return ret;
		}

		public static void _AddActionResultHandler(ClientObject clientObj, Action action, IResultHandler resultHandler)
        {
			clientObj.Context._PendingRequest.AddActionResultHandler(action, resultHandler);
		}

        public static void _Load(ClientObject clientObject, LoadOption loadOption)
        {
            clientObject.Context.Load(clientObject, loadOption);
        }

		private static string NormalizeName(string name)
        {
			return name.Substring(0, 1).ToLowerInvariant() + name.Substring(1);
		}

        internal static string ToJsonString(object obj)
        {
            Newtonsoft.Json.JsonSerializerSettings settings = new Newtonsoft.Json.JsonSerializerSettings();
            settings.NullValueHandling = Newtonsoft.Json.NullValueHandling.Ignore;
            return Newtonsoft.Json.JsonConvert.SerializeObject(obj, settings);
        }

        internal static JToken ToJsonObject(string value)
        {
            JToken ret = Newtonsoft.Json.Linq.JToken.Parse(value);
            return ret;
        }

		internal static Exception ConvertJsComErrorToException(JToken jsonError)
		{
			string code = jsonError[Constants.Code].ToObject<string>();
			string location = string.Empty;
			if (jsonError[Constants.Location] != null)
			{
				location = jsonError[Constants.Location].ToObject<string>();
			}

			string message = string.Empty;
			if (jsonError[Constants.Message] != null)
			{
				message = jsonError[Constants.Message].ToObject<string>();
			}

			if (string.IsNullOrEmpty(message))
			{
				message = code;
			}

			return new Exception(message);
		}

		internal static Exception ConvertODataErrorToException(JToken jsonError)
		{
			string code = jsonError[Constants.ODataCode].ToObject<string>();

			string message = string.Empty;
			if (jsonError[Constants.ODataMessage] != null)
			{
				message = jsonError[Constants.ODataMessage].ToObject<string>();
			}

			if (string.IsNullOrEmpty(message))
			{
				message = code;
			}

			return new Exception(message);
		}
	}
}
