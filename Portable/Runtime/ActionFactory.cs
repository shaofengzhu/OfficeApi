using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeExtension
{
    public static class ActionFactory
    {
        public static Action _CreateSetPropertyAction(ClientRequestContext context, ClientObject parent, string propertyName, object value)
        {
			Utility.ValidateObjectPath(parent);
			ActionInfo actionInfo = new ActionInfo()
				{
					Id = context._NextId(),
					ActionType = ActionType.SetProperty,
					Name = propertyName,
					ObjectPathId = parent._ObjectPath.ObjectPathInfo.Id,
					ArgumentInfo = new ArgumentInfo()
				};
            object[] args = new object[] { value };
			var referencedArgumentObjectPaths = Utility.SetMethodArguments(context, actionInfo.ArgumentInfo, args);
            Utility.ValidateReferencedObjectPaths(referencedArgumentObjectPaths);
			var ret = new Action(actionInfo, true);
            context._PendingRequest.AddAction(ret);
			context._PendingRequest.AddReferencedObjectPath(parent._ObjectPath);
			context._PendingRequest.AddReferencedObjectPaths(referencedArgumentObjectPaths);

			return ret;
		}

        public static Action _CreateMethodAction(ClientRequestContext context, ClientObject parent, string methodName, OperationType operationType, object[] args)
        {
			Utility.ValidateObjectPath(parent);
			ActionInfo actionInfo = new ActionInfo()
				{
					Id = context._NextId(),
					ActionType = ActionType.Method,
					Name = methodName,
					ObjectPathId = parent._ObjectPath.ObjectPathInfo.Id,
					ArgumentInfo = new ArgumentInfo()
				};
			var referencedArgumentObjectPaths = Utility.SetMethodArguments(context, actionInfo.ArgumentInfo, args);
            Utility.ValidateReferencedObjectPaths(referencedArgumentObjectPaths);
			bool isWriteOperation = operationType != OperationType.Read;
            var ret = new Action(actionInfo, isWriteOperation);
            context._PendingRequest.AddAction(ret);
			context._PendingRequest.AddReferencedObjectPath(parent._ObjectPath);
			context._PendingRequest.AddReferencedObjectPaths(referencedArgumentObjectPaths);
			return ret;
		}

		internal static Action CreateQueryAction(ClientRequestContext context, ClientObject parent, QueryInfo queryOption)
        {
			Utility.ValidateObjectPath(parent);
            ActionInfo actionInfo = new ActionInfo()
				{
					Id = context._NextId(),
					ActionType = ActionType.Query,
					Name = "",
					ObjectPathId = parent._ObjectPath.ObjectPathInfo.Id,
				};
			actionInfo.QueryInfo = queryOption;
			Action ret = new Action(actionInfo, false);
            context._PendingRequest.AddAction(ret);
			context._PendingRequest.AddReferencedObjectPath(parent._ObjectPath);
			return ret;
		}

		internal static Action CreateInstantiateAction(ClientRequestContext context, ClientObject obj)
        {
			Utility.ValidateObjectPath(obj);
			ActionInfo actionInfo = new ActionInfo()
				{
					Id = context._NextId(),
					ActionType = ActionType.Instantiate,
					Name = "",
					ObjectPathId = obj._ObjectPath.ObjectPathInfo.Id
				};
			var ret = new Action(actionInfo, false);
            context._PendingRequest.AddAction(ret);
			context._PendingRequest.AddReferencedObjectPath(obj._ObjectPath);
			context._PendingRequest.AddActionResultHandler(ret, new InstantiateActionResultHandler(obj));
			return ret;
		}

		internal static Action CreateTraceAction(ClientRequestContext context, string message)
        {
			ActionInfo actionInfo = new ActionInfo()
				{
					Id = context._NextId(),
					ActionType = ActionType.Trace,
					Name = "Trace",
					ObjectPathId = 0
				};
			Action ret = new Action(actionInfo, false);
            context._PendingRequest.AddAction(ret);
			context._PendingRequest.AddTrace(actionInfo.Id, message);
			return ret;
		}
	}
}
