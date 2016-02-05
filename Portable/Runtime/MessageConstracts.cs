using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

// This file contains the message contract between the Javascript and the unmanaged
// code implementation.
// We intentionally keep all of them in one file to make it easy to understand

namespace OfficeExtension
{
    public enum RichApiRequestMessageIndex
    {
        CustomData = 0,
        Method = 1,
        PathAndQuery = 2,
        Headers = 3, // It's string[] and its value is [name1, value1, name2, value2, ...]
        Body = 4,
        AppPermission = 5,
        RequestFlags = 6,
    }

    public enum RichApiResponseMessageIndex
    {
        StatusCode = 0,
        Headers = 1,
        Body = 2,
    }

    public enum ActionType
    {
        Instantiate = 1,
        Query = 2,
        Method = 3,
        SetProperty = 4,
        Trace = 5,
    }

    public enum ObjectPathType
    {
        GlobalObject = 1,
        NewObject = 2,
        Method = 3,
        Property = 4,
        Indexer = 5,
        ReferenceId = 6,
    }

    public class ArgumentInfo
    {
        public object[] Arguments;
        // If it's an ClientObject, which corresponding to the IDispatch, the argument value is the ObjectPathId;
        public object[] ReferencedObjectPathIds;
	}

    public class QueryInfo
    {
        public string[] Select;
        public string[] Expand;
        public int? Skip;
        public int? Top;
	}

    public class ActionInfo
    {
        public int Id;
		public ActionType ActionType;
		public string Name;
		public int ObjectPathId;
		public ArgumentInfo  ArgumentInfo;
		public QueryInfo QueryInfo;
	}

    public class ActionResult
    {
        public int ActionId;
		public object Value;
	}

    public class ObjectPathInfo
    {
        public int Id;
		public ObjectPathType ObjectPathType;
		public string Name;
		public int? ParentObjectPathId;
		public ArgumentInfo ArgumentInfo;
	}

    public class RequestMessageBody
    {
        public List<ActionInfo> Actions;
        public Dictionary<int, ObjectPathInfo> ObjectPaths;
    }

	public class ErrorInfo
    {
        public string Code;
		public string Message;
		public string Location;
	}

    public class ResponseMessageBody
    {
        public ErrorInfo Error;
        public List<ActionResult> Results;
        public List<int> TraceIds;
	}
}
