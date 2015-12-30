import sys
import json
import enum
class Constants:
    getItemAt = "GetItemAt"
    id = "Id"
    idPrivate = "_Id"
    index = "_Index"
    items = "_Items"
    iterativeExecutor = "IterativeExecutor"
    localDocument = "http://document.localhost/"
    localDocumentApiPrefix = "http://document.localhost/_api/"
    referenceId = "_ReferenceId";

class ClientRuntimeContext:
    def __init__(self, url):
        self._url = url
        self._pendingRequest = None
        self._nextId = 1

    @property
    def url(self):
        return self._url

    @property
    def pendingRequest(self):
        if (self._pendingRequest is None):
            self._pendingRequest = ClientRequest(self)
        return self._pendingRequest
    
    @property
    def _nextId(self):
        ret = self._nextId
        self._nextId = self._nextId + 1
        return ret


class ClientRequest:
    _actions = []
    _context = None
    def __init__(self, context):
        self._context = context

    @property
    def context(self):
        return self._context

class ClientObject:
    _context = None
    def __init__(self, context):
        self._context = context
    
    @property
    def context(self):
        return self._context

class Action:
    pass

class ActionFactory:
    pass

class RichApiRequestMessageIndex(enum.IntEnum):
    CustomData = 0
    Method = 1
    PathAndQuery = 2
    Headers = 3
    Body = 4
    AppPermission = 5
    RequestFlags = 6

class RichApiResponseMessageIndex(enum.IntEnum):
    StatusCode = 0
    Headers = 1
    Body = 2

class ActionType(enum.IntEnum):
    Instantiate = 1
    Query = 2
    Method = 3
    SetProperty = 4
    Trace = 5

class OperationType(enum.IntEnum):
    Default = 0
    Read = 1

class ObjectPathType(enum.IntEnum):
    GlobalObject = 1
    NewObject = 2
    Method = 3
    Property = 4
    Indexer = 5
    ReferenceId = 6

class ClientRequestFlags(enum.IntEnum):
    NoneValue = 0
    WriteOperation = 1

class ArgumentInfo:
    Arguments = None
    ReferencedObjectPathIds = None

class QueryInfo:
	Select = None
	Expand = None
	Skip = None
	Top = None


class ActionInfo:
	Id = 0
	ActionType = None
	Name = None
	ObjectPathId = 0
	ArgumentInfo = None
	QueryInfo = None


class ActionResultInfo:
	ActionId = 0
	Value = None

class ObjectPathInfo:
	Id = 0
	ObjectPathType = None
	Name = None
	ParentObjectPathId = 0
	ArgumentInfo = None

class RequestMessageBodyInfo:
    Actions = None
    ObjectPaths = None


class Action:
    def __init__(self, actionInfo, isWriteOperation):
        self._actionInfo = actionInfo
        self._isWriteOperation = isWriteOperation
    
    @property
    def actionInfo(self):
        return self._actionInfo

    @property
    def isWriteOperation(self):
        return self._isWriteOperation

class ActionFactory:
    @staticmethod
    def createSetPropertyAction(context, parent, propertyName, value):
        Utility.validateObjectPath(parent)
        actionInfo = ActionInfo()
        actionInfo.Id = context._nextId()
        actionInfo.ActionType = ActionType.SetProperty
        actionInfo.Name = propertyName,
        actionInfo.ObjectPathId = parent._objectPath.objectPathInfo.Id,
        actionInfo.ArgumentInfo = {}		
        args = [value]
        referencedArgumentObjectPaths = Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
        Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
        ret = Action(actionInfo, true);
        context._pendingRequest.addAction(ret);
        context._pendingRequest.addReferencedObjectPath(parent._objectPath);
        context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
        return ret;

    @staticmethod
    def createMethodAction(context, parent, methodName, operationType, args):
        Utility.validateObjectPath(parent);
        actionInfo = ActionInfo()
        actionInfo.Id = context._nextId()
        actionInfo.ActionType = ActionType.Method
        actionInfo.Name = methodName
        actionInfo.ObjectPathId = parent._objectPath.objectPathInfo.Id
        actionInfo.ArgumentInfo = {}
        referencedArgumentObjectPaths = Utility.setMethodArguments(context, actionInfo.ArgumentInfo, args);
        Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
        isWriteOperation = operationType != OperationType.Read;
        ret = Action(actionInfo, isWriteOperation);
        context._pendingRequest.addAction(ret);
        context._pendingRequest.addReferencedObjectPath(parent._objectPath);
        context._pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
        return ret;

    @staticmethod
    def createQueryAction(context, parent, queryInfo): 
        Utility.validateObjectPath(parent);
        actionInfo = ActionInfo()
        actionInfo.Id = context._nextId(),
        actionInfo.ActionType = ActionType.Query,
        actionInfo.Name = ""
        actionInfo.ObjectPathId = parent._objectPath.objectPathInfo.Id,
        actionInfo.QueryInfo = queryInfo;
        ret = Action(actionInfo, false)
        context._pendingRequest.addAction(ret);
        context._pendingRequest.addReferencedObjectPath(parent._objectPath);
        return ret;

    @staticmethod
    def createInstantiateAction(context, obj):
        Utility.validateObjectPath(obj)
        actionInfo = ActionInfo()
        actionInfo.Id = context._nextId()
        actionInfo.ActionType = ActionType.Instantiate
        actionInfo.Name = ""
        actionInfo.ObjectPathId = obj._objectPath.objectPathInfo.Id
        ret = Action(actionInfo, false);
        context._pendingRequest.addAction(ret);
        context._pendingRequest.addReferencedObjectPath(obj._objectPath);
        handler = InstantiateActionResultHandler(obj)
        context._pendingRequest.addActionResultHandler(ret, handler);
        return ret;

    @staticmethod
    def createTraceAction(context, message):
        actionInfo = ActionInfo()
        actionInfo.Id = context._nextId()
        actionInfo.ActionType = ActionType.Trace
        actionInfo.Name = "Trace"
        actionInfo.ObjectPathId = 0
        ret = Action(actionInfo, false)
        context._pendingRequest.addAction(ret)
        context._pendingRequest.addTrace(actionInfo.Id, message)
        return ret

class ObjectPath:
    def __init(self, objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest):
        self._objectPathInfo = objectPathInfo
        self._parentObjectPath = parentObjectPath
        self._isWriteOperation = false
        self._isCollection = isCollection
        self._isInvalidAfterRequest = isInvalidAfterRequest
        self._isValid = true
        self._argumentObjectPaths = None

    @property
    def objectPathInfo(self):
        return self._objectPathInfo

    @property
    def isWriteOperation(self):
        return self._isWriteOperation

    @isWriteOperation.setter
    def isWriteOperation(self, value):
        self._isWriteOperation = value

    @property
    def isCollection(self):
        return self._isCollection;

    @property
    def isInvalidAfterRequest(self):
        return self._isInvalidAfterRequest

    @property
    def parentObjectPath(self):
        return self._parentObjectPath

    @property
    def argumentObjectPaths(self):
        return self._argumentObjectPaths

    @argumentObjectPaths.setter
    def argumentObjectPaths(self, value):
        self._argumentObjectPaths = value


    @property
    def isValid(self):
        return self._isValid;

    @isValid.setter
    def isValid(self, value):
        self._isValid = value;


    def updateUsingObjectData(self, value):
        referenceId = value.get(Constants.referenceId, None)
        if not Utility.isNullOrEmptyString(referenceId):
            self._isInvalidAfterRequest = false
            self._isValid = true
            self._objectPathInfo.ObjectPathType = ObjectPathType.ReferenceId
            self._objectPathInfo.Name = referenceId
            self._objectPathInfo.ArgumentInfo = {}
            self._parentObjectPath = None
            self._argumentObjectPaths = None
            return

        if self.parentObjectPath and self.parentObjectPath.isCollection:
            id = value.get(Constants.id, None)
            if Utility.isNullOrUndefined(id):
                id = value.get(Constants.idPrivate, None)

            if not Utility.isNullOrUndefined(id):
                self._isInvalidAfterRequest = false
                self._isValid = true
                self._objectPathInfo.ObjectPathType = ObjectPathType.Indexer
                self._objectPathInfo.Name = ""
                self._objectPathInfo.ArgumentInfo = {}
                self._objectPathInfo.ArgumentInfo.Arguments = [id]
                self._argumentObjectPaths = None
                return

class ObjectPathFactory:
    @staticmethod
    def createGlobalObjectObjectPath(context: ClientRequestContext):
        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.GlobalObject
        objectPathInfo.Name = ""
        return ObjectPath(objectPathInfo,
                          None, 
                          false,    # isCollection
                          false     # isInvalidAfterRequest
                          )

    @staticmethod
    def createNewObjectObjectPath(context, typeName, isCollection):
        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.NewObject
        objectPathInfo.Name = typeName
        return ObjectPath(objectPathInfo, 
                          None, 
                          isCollection, 
                          false     # isInvalidAfterRequest
                          )

    @staticmethod
    def createPropertyObjectPath(context, parent, propertyName, isCollection, isInvalidAfterRequest):
        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.Property
        objectPathInfo.Name = propertyName
        objectPathInfo.ParentObjectPathId = parent._objectPath.objectPathInfo.Id
        return ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest)
    

    @staticmethod
    def createIndexerObjectPath(context, parent, args):
        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.Indexer
        objectPathInfo.Name = ""
        objectPathInfo.ParentObjectPathId = parent._objectPath.objectPathInfo.Id
        objectPathInfo.ArgumentInfo = {}
        objectPathInfo.ArgumentInfo.Arguments = args
        return ObjectPath(objectPathInfo, 
                          parent._objectPath, 
                          false,    # isCollection
                          false     # isInvalidAfterRequest
                          )
    

    @staticmethod
    def createIndexerObjectPathUsingParentPath(context, parentObjectPath, args):
        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.Indexer
        objectPathInfo.Name = ""
        objectPathInfo.ParentObjectPathId = parentObjectPath.objectPathInfo.Id,
        objectPathInfo.ArgumentInfo = {}
        objectPathInfo.ArgumentInfo.Arguments = args;
        return ObjectPath(objectPathInfo, 
                          parentObjectPath, 
                          false,    # isCollection
                          false     # isInvalidAfterRequest
                          )
    
    @staticmethod
    def createMethodObjectPath(context, parentObject, methodName, operationType, args, isCollection, isInvalidAfterRequest):
        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.Method
        objectPathInfo.Name = methodName
        objectPathInfo.ParentObjectPathId = parentObject._objectPath.objectPathInfo.Id,
        objectPathInfo.ArgumentInfo = {}

        argumentObjectPaths = Utility.setMethodArguments(context, objectPathInfo.ArgumentInfo, args);
        ret = ObjectPath(objectPathInfo, parent._objectPath, isCollection, isInvalidAfterRequest);
        ret.argumentObjectPaths = argumentObjectPaths;
        ret.isWriteOperation = (operationType != OperationType.Read);
        return ret;
    
    @staticmethod
    def createChildItemObjectPathUsingIndexerOrGetItemAt(hasIndexerMethod, context, parentObject, childItem, index):
        id = childItem.get(Constants.id, None)
        if Utility.isNullOrUndefined(id):
            id = childItem.get(Constants.idPrivate)

        if hasIndexerMethod and not Utility.isNullOrUndefined(id):
            return ObjectPathFactory.createChildItemObjectPathUsingIndexer(context, parentObject, childItem)
        else:
            return ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(context, parentObject, childItem, index)

    @staticmethod
    def createChildItemObjectPathUsingIndexer(context, parentObject, childItem):
        id = childItem.get(Constants.id)
        if Utility.isNullOrUndefined(id):
            id = childItem.get(Constants.idPrivate)

        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id =context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.Indexer
        objectPathInfo.Name = ""
        objectPathInfo.ParentObjectPathId = parentObject._objectPath.objectPathInfo.Id,
        objectPathInfo.ArgumentInfo = {}
        objectPathInfo.ArgumentInfo.Arguments = [id];
        return ObjectPath(objectPathInfo, parent._objectPath, 
                          false, # isCollection
                          false # isInvalidAfterRequest
                          )
    
    @staticmethod
    def createChildItemObjectPathUsingGetItemAt(context, parentObject, childItem, index):
        indexFromServer = childItem.get(Constants.index);
        if indexFromServer:
            index = indexFromServer;

        objectPathInfo = ObjectPathInfo()
        objectPathInfo.Id = context._nextId()
        objectPathInfo.ObjectPathType = ObjectPathType.Method
        objectPathInfo.Name = Constants.getItemAt
        objectPathInfo.ParentObjectPathId = parent._objectPath.objectPathInfo.Id
        objectPathInfo.ArgumentInfo = {}
        objectPathInfo.ArgumentInfo.Arguments = [index];
        return ObjectPath(objectPathInfo, parent._objectPath, 
                          false, # isCollection
                          false # isInvalidAfterRequest
                          )
