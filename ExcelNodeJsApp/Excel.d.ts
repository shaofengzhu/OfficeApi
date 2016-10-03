declare module OfficeExtension {
    class Action {
        private m_actionInfo;
        private m_isWriteOperation;
        constructor(actionInfo: ActionInfo, isWriteOperation: boolean);
        actionInfo: ActionInfo;
        isWriteOperation: boolean;
    }
}
declare module OfficeExtension {
    class ActionFactory {
        static createSetPropertyAction(context: ClientRequestContext, parent: ClientObject, propertyName: string, value: any): Action;
        static createMethodAction(context: ClientRequestContext, parent: ClientObject, methodName: string, operationType: OperationType, args: Array<any>): Action;
        static createQueryAction(context: ClientRequestContext, parent: ClientObject, queryOption: QueryInfo): Action;
        static createRecursiveQueryAction(context: ClientRequestContext, parent: ClientObject, query: RecursiveQueryInfo): Action;
        static createInstantiateAction(context: ClientRequestContext, obj: ClientObject): Action;
        static createTraceAction(context: ClientRequestContext, message: string, addTraceMessage: boolean): Action;
    }
}
declare module OfficeExtension {
    class ClientObject implements IResultHandler {
        private m_context;
        private m_objectPath;
        private m_isNull;
        constructor(context: ClientRequestContext, objectPath: ObjectPath);
        context: ClientRequestContext;
        _objectPath: ObjectPath;
        isNull: boolean;
        _isNull: boolean;
        _handleResult(value: any): void;
        _handleIdResult(value: any): void;
    }
}
declare module OfficeExtension {
    class ClientRequest {
        private m_flags;
        private m_context;
        private m_actions;
        private m_actionResultHandler;
        private m_referencedObjectPaths;
        private m_traceInfos;
        private m_responseTraceIds;
        private m_responseTraceMessages;
        private m_pendingProcessEventHandlers;
        private m_pendingEventHandlerActions;
        constructor(context: ClientRequestContext);
        flags: ClientRequestFlags;
        traceInfos: {
            [index: number]: string;
        };
        _responseTraceMessages: Array<string>;
        _responseTraceIds: {
            [index: number]: number;
        };
        _setResponseTraceIds(value: Array<number>): void;
        addAction(action: Action): void;
        hasActions: boolean;
        addTrace(actionId: number, message: string): void;
        addReferencedObjectPath(objectPath: ObjectPath): void;
        addReferencedObjectPaths(objectPaths: Array<ObjectPath>): void;
        addActionResultHandler(action: Action, resultHandler: IResultHandler): void;
        buildRequestMessageBody(): RequestMessageBody;
        processResponse(msg: ResponseMessageBody): void;
        invalidatePendingInvalidObjectPaths(): void;
        _addPendingEventHandlerAction(eventHandlers: EventHandlers<any>, action: _Internal.EventHandlerAction): void;
        _pendingProcessEventHandlers: Array<EventHandlers<any>>;
        _getPendingEventHandlerActions(eventHandlers: EventHandlers<any>): Array<_Internal.EventHandlerAction>;
    }
}
declare module OfficeExtension {
    interface LoadOption {
        select?: string | string[];
        expand?: string | string[];
        top?: number;
        skip?: number;
    }
    function _setRequestExecutorFactory(reqExecFactory: () => IRequestExecutor): void;
    interface RequestUrlAndHeaderInfo {
        url: string;
        headers?: {
            [name: string]: string;
        };
    }
    class ClientRequestContext {
        private m_nextId;
        private m_pendingRequest;
        private m_url;
        private m_trackedObjects;
        private m_requestHeaders;
        static defaultRequestUrlAndHeaders: RequestUrlAndHeaderInfo;
        _autoCleanup: boolean;
        _onRunFinishedNotifiers: Array<() => void>;
        _requestExecutor: IRequestExecutor;
        _rootObject: ClientObject;
        _customData: string;
        _processingResult: boolean;
        constructor(url?: string);
        _pendingRequest: ClientRequest;
        trackedObjects: TrackedObjects;
        requestHeaders: {
            [name: string]: string;
        };
        load(clientObj: ClientObject, option?: string | string[] | LoadOption): void;
        private static parseQueryOption(option);
        loadRecursive(clientObj: ClientObject, options: {
            [typeName: string]: string | string[] | LoadOption;
        }, maxDepth?: number): void;
        trace(message: string): void;
        private syncPrivate();
        private processRequestExecutorResponseMessage(req, response);
        private processPendingEventHandlers(req);
        private createProcessOneEventHandlersFunc(eventHandlers, req);
        sync<T>(passThroughValue?: T): IPromise<T>;
        static _run<TResult, TContext extends ClientRequestContext>(ctxInitializer: () => TContext, batch: (ctx: TContext) => OfficeExtension.IPromise<any>, numCleanupAttempts?: number, retryDelay?: number, onCleanupSuccess?: (counter: number) => void, onCleanupFailure?: (counter: number) => void): IPromise<TResult>;
        static _runBatch<TResult, TContext extends ClientRequestContext>(functionName: string, receivedRunArgs: IArguments, ctxInitializer: (requestInfo?: RequestUrlAndHeaderInfo) => TContext, numCleanupAttempts?: number, retryDelay?: number, onCleanupSuccess?: (counter: number) => void, onCleanupFailure?: (counter: number) => void): IPromise<TResult>;
        static createErrorPromise(functionName: any, code?: string): IPromise<{}>;
        private static _runCommon<TResult, TContext>(functionName, requestInfo, ctxRetriever, batch, numCleanupAttempts, retryDelay, onCleanupSuccess?, onCleanupFailure?);
        _nextId(): number;
    }
}
declare module OfficeExtension {
    enum ClientRequestFlags {
        None = 0,
        WriteOperation = 1,
    }
}
declare module OfficeExtension {
    enum ClientResultProcessingType {
        none = 0,
        date = 1,
    }
    class ClientResult<T> implements IResultHandler {
        private m_value;
        private m_isLoaded;
        private m_type;
        constructor(type?: ClientResultProcessingType);
        value: T;
        _handleResult(value: any): void;
    }
}
declare module OfficeExtension {
    class Constants {
        static getItemAt: string;
        static id: string;
        static idPrivate: string;
        static index: string;
        static items: string;
        static iterativeExecutor: string;
        static localDocument: string;
        static localDocumentApiPrefix: string;
        static referenceId: string;
        static isTracked: string;
        static sourceLibHeader: string;
        static requestInfoHeader: string;
    }
}
declare module OfficeExtension {
    class EmbedRequestExecutor implements IRequestExecutor {
        private static DispidExecuteRichApiRequestMethod;
        private static SourceLibHeaderValue;
        executeAsync(customData: string, requestFlags: number, requestMessage: IRequestExecutorRequestMessage): IPromise<IRequestExecutorResponseMessage>;
        private static _transformMessageArrayIntoParams(msgArray);
    }
}
declare module OfficeExtension {
    module _Internal {
        class Error {
            name: string;
            message: string;
            stack: string;
            constructor(message?: string);
        }
        class RuntimeError extends Error {
            code: string;
            traceMessages: Array<string>;
            debugInfo: {
                errorLocation?: string;
            };
            constructor(code: string, message: string, traceMessages: Array<string>, debugInfo: {
                errorLocation?: string;
            });
            toString(): string;
        }
    }
    class Error {
        name: string;
        message: string;
        stack: string;
        code: string;
        traceMessages: Array<string>;
        debugInfo: {
            errorLocation?: string;
        };
        constructor(message?: string);
    }
}
declare module OfficeExtension {
    class ErrorCodes {
        static accessDenied: string;
        static generalException: string;
        static activityLimitReached: string;
        static invalidObjectPath: string;
        static propertyNotLoaded: string;
        static valueNotLoaded: string;
        static invalidRequestContext: string;
        static invalidArgument: string;
        static runMustReturnPromise: string;
        static cannotRegisterEvent: string;
        static apiNotFound: string;
        static connectionFailure: string;
    }
}
declare module OfficeExtension {
    interface EventInfo<T> {
        registerFunc: (callback: (args: any) => void) => IPromise<any>;
        unregisterFunc: (callback: (args: any) => void) => IPromise<any>;
        eventArgsTransformFunc: (args: any) => IPromise<T>;
    }
}
declare module OfficeExtension {
    module _Internal {
        enum EventHandlerActionType {
            add = 0,
            remove = 1,
            removeAll = 2,
        }
        interface EventHandlerAction {
            id: number;
            operation: EventHandlerActionType;
            handler: (args: any) => IPromise<any>;
        }
    }
}
declare module OfficeExtension {
    class EventHandlers<T> {
        private m_context;
        private m_name;
        private m_handlers;
        private m_registered;
        private m_eventInfo;
        private m_callback;
        private m_id;
        constructor(context: ClientRequestContext, parentObject: ClientObject, name: string, eventInfo: EventInfo<T>);
        _registered: boolean;
        _id: number;
        _handlers: Array<(args: T) => IPromise<any>>;
        add(handler: (args: T) => IPromise<any>): EventHandlerResult<T>;
        remove(handler: (args: T) => IPromise<any>): void;
        removeAll(): void;
        _processRegistration(req: ClientRequest): IPromise<any>;
        private fireEvent(args);
        private createFireOneEventHandlerFunc(handler, args);
    }
}
declare module OfficeExtension {
    class EventHandlerResult<T> {
        private m_context;
        private m_allHandlers;
        private m_handler;
        constructor(context: ClientRequestContext, handlers: EventHandlers<T>, handler: (args: T) => IPromise<any>);
        remove(): void;
    }
}
declare module OfficeExtension {
    interface HttpRequestInfo {
        method: string;
        url: string;
        headers: {
            [name: string]: string;
        };
        body: string;
    }
    interface HttpResponseInfo {
        statusCode: number;
        headers: {
            [name: string]: string;
        };
        body: string;
    }
    function setCustomHttpRequestExecuteFunc(func: (request: HttpRequestInfo) => IPromise<HttpResponseInfo>): void;
    class HttpRequestExecutor implements IRequestExecutor {
        private static SourceLibHeaderValue;
        executeAsync(customData: string, requestFlags: number, requestMessage: IRequestExecutorRequestMessage): IPromise<IRequestExecutorResponseMessage>;
    }
}
declare module OfficeExtension {
    class InstantiateActionResultHandler implements IResultHandler {
        private m_clientObject;
        constructor(clientObject: ClientObject);
        _handleResult(value: any): void;
    }
}
declare module OfficeExtension {
    interface IPromise<R> {
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => U): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => void): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => U): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => void): IPromise<U>;
        catch<U>(onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        catch<U>(onRejected?: (error: any) => U): IPromise<U>;
        catch<U>(onRejected?: (error: any) => void): IPromise<U>;
    }
    class Promise<R> implements IPromise<R> {
        constructor(func: (resolve, reject) => void);
        static all<U>(promises: OfficeExtension.IPromise<U>[]): IPromise<U[]>;
        static resolve<U>(value: U): IPromise<U>;
        static reject<U>(error: any): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => U): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => void): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => U): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => void): IPromise<U>;
        catch<U>(onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        catch<U>(onRejected?: (error: any) => U): IPromise<U>;
        catch<U>(onRejected?: (error: any) => void): IPromise<U>;
    }
}
declare module OfficeExtension {
    interface IRequestExecutorRequestMessage {
        Url: string;
        Headers: {
            [name: string]: string;
        };
        Body: RequestMessageBody;
    }
    interface IRequestExecutorResponseMessage {
        ErrorCode: string;
        ErrorMessage: string;
        Headers: {
            [name: string]: string;
        };
        Body: ResponseMessageBody;
    }
    interface IRequestExecutor {
        executeAsync(customData: string, requestFlags: number, requestMessage: IRequestExecutorRequestMessage): IPromise<IRequestExecutorResponseMessage>;
    }
}
declare module OfficeExtension {
    interface IResultHandler {
        _handleResult(value: any): void;
    }
}
declare module OfficeExtension {
    enum RichApiRequestMessageIndex {
        CustomData = 0,
        Method = 1,
        PathAndQuery = 2,
        Headers = 3,
        Body = 4,
        AppPermission = 5,
        RequestFlags = 6,
    }
    enum RichApiResponseMessageIndex {
        StatusCode = 0,
        Headers = 1,
        Body = 2,
    }
    enum ActionType {
        Instantiate = 1,
        Query = 2,
        Method = 3,
        SetProperty = 4,
        Trace = 5,
        RecursiveQuery = 6,
    }
    enum ObjectPathType {
        GlobalObject = 1,
        NewObject = 2,
        Method = 3,
        Property = 4,
        Indexer = 5,
        ReferenceId = 6,
        NullObject = 7,
    }
    interface ArgumentInfo {
        Arguments?: Array<any>;
        ReferencedObjectPathIds?: Array<number | Array<number>>;
    }
    interface QueryInfo {
        Select?: Array<string>;
        Expand?: Array<string>;
        Skip?: number;
        Top?: number;
    }
    interface RecursiveQueryInfo {
        MaxDepth?: number;
        Queries: {
            [typeName: string]: QueryInfo;
        };
    }
    interface ActionInfo {
        Id: number;
        ActionType: ActionType;
        Name: string;
        ObjectPathId: number;
        ArgumentInfo?: ArgumentInfo;
        QueryInfo?: QueryInfo;
        RecursiveQueryInfo?: RecursiveQueryInfo;
    }
    interface ActionResult {
        ActionId: number;
        Value: any;
    }
    interface ObjectPathInfo {
        Id: number;
        ObjectPathType: ObjectPathType;
        Name: string;
        ParentObjectPathId?: number;
        ArgumentInfo?: ArgumentInfo;
    }
    interface RequestMessageBody {
        Actions: Array<ActionInfo>;
        ObjectPaths: {
            [Id: number]: ObjectPathInfo;
        };
    }
    interface ErrorInfo {
        Code: string;
        Message: string;
        Location: string;
    }
    interface ResponseMessageBody {
        Error: ErrorInfo;
        Results: Array<ActionResult>;
        TraceIds: Array<number>;
    }
}
declare module OfficeExtension {
    class ObjectPath {
        private m_objectPathInfo;
        private m_isWriteOperation;
        private m_parentObjectPath;
        private m_argumentObjectPaths;
        private m_isCollection;
        private m_isInvalidAfterRequest;
        private m_isValid;
        private m_getByIdMethodName;
        constructor(objectPathInfo: ObjectPathInfo, parentObjectPath: ObjectPath, isCollection: boolean, isInvalidAfterRequest: boolean);
        objectPathInfo: ObjectPathInfo;
        isWriteOperation: boolean;
        isCollection: boolean;
        isInvalidAfterRequest: boolean;
        parentObjectPath: ObjectPath;
        argumentObjectPaths: Array<ObjectPath>;
        isValid: boolean;
        getByIdMethodName: string;
        _updateAsNullObject(): void;
        updateUsingObjectData(value: Object): void;
    }
}
declare module OfficeExtension {
    class ObjectPathFactory {
        static createGlobalObjectObjectPath(context: ClientRequestContext): ObjectPath;
        static createNewObjectObjectPath(context: ClientRequestContext, typeName: string, isCollection?: boolean): ObjectPath;
        static createPropertyObjectPath(context: ClientRequestContext, parent: ClientObject, propertyName: string, isCollection?: boolean, isInvalidAfterRequest?: boolean): ObjectPath;
        static createIndexerObjectPath(context: ClientRequestContext, parent: ClientObject, args: Array<any>): ObjectPath;
        static createIndexerObjectPathUsingParentPath(context: ClientRequestContext, parentObjectPath: ObjectPath, args: Array<any>): ObjectPath;
        static createMethodObjectPath(context: ClientRequestContext, parent: ClientObject, methodName: string, operationType: OperationType, args: Array<any>, isCollection?: boolean, isInvalidAfterRequest?: boolean, getByIdMethodName?: string): ObjectPath;
        static createChildItemObjectPathUsingIndexerOrGetItemAt(hasIndexerMethod: boolean, context: ClientRequestContext, parent: ClientObject, childItem: Object, index: number): ObjectPath;
        static createChildItemObjectPathUsingIndexer(context: ClientRequestContext, parent: ClientObject, childItem: Object): ObjectPath;
        static createChildItemObjectPathUsingGetItemAt(context: ClientRequestContext, parent: ClientObject, childItem: Object, index: number): ObjectPath;
    }
}
declare module OfficeExtension {
    class OfficeJsRequestExecutor implements IRequestExecutor {
        private static SourceLibHeaderValue;
        executeAsync(customData: string, requestFlags: number, requestMessage: IRequestExecutorRequestMessage): IPromise<IRequestExecutorResponseMessage>;
    }
}
declare module OfficeExtension {
    class OfficeXHRSettings {
        oldxhr: () => IXMLHttpRequest;
        executeRichApiRequestAsync: (message: Array<any>, callback: (result: OSF.DDA.RichApi.ExecuteRichApiRequestResult) => void) => void;
    }
    function resetXHRFactory(oldFactory: () => IXMLHttpRequest): () => IXMLHttpRequest;
    function officeXHRFactory(): OfficeXHR;
    class OfficeXHR implements IXMLHttpRequest {
        private static UNSENT;
        private static OPENED;
        private static DONE;
        static settings: OfficeXHRSettings;
        private m_innerXhr;
        private m_method;
        private m_url;
        private m_allResponseHeaders;
        private m_responseHeaders;
        private m_requestHeaders;
        onreadystatechange: () => void;
        readyState: number;
        status: number;
        statusText: string;
        response: any;
        responseText: string;
        responseType: string;
        open(method: string, url: string): void;
        abort(): void;
        send(body: string): void;
        setRequestHeader(header: string, value: string): void;
        getResponseHeader(header: string): string;
        getAllResponseHeaders(): string;
        overrideMimeType(mimeType: string): void;
        private innerXhrOnreadystatechage();
        private officeContextRequestCallback(result);
        private setAllResponseHeaders(allResponseHeaders);
    }
}
declare module OfficeExtension {
    module _Internal.PromiseImpl {
        function Init(): void;
    }
}
declare function IsEdgeLessThan14(): boolean;
declare module OfficeExtension {
    enum OperationType {
        Default = 0,
        Read = 1,
    }
}
declare module OfficeExtension {
    class TrackedObjects {
        private m_context;
        private _autoCleanupList;
        constructor(context: ClientRequestContext);
        add(object: ClientObject): void;
        add(objects: ClientObject[]): void;
        _autoAdd(object: ClientObject): void;
        private _addCommon(object, isExplicitlyAdded);
        remove(object: ClientObject): void;
        remove(objects: ClientObject[]): void;
        private _removeCommon(object);
        _retrieveAndClearAutoCleanupList(): {
            [id: number]: ClientObject;
        };
    }
}
declare module OfficeExtension {
    class ResourceStrings {
        static invalidObjectPath: string;
        static propertyNotLoaded: string;
        static valueNotLoaded: string;
        static invalidRequestContext: string;
        static invalidArgument: string;
        static runMustReturnPromise: string;
        static cannotRegisterEvent: string;
        static connectionFailureWithStatus: string;
        static connectionFailureWithDetails: string;
    }
}
declare module OfficeExtension {
    class RichApiMessageUtility {
        private static OfficeJsErrorCode_ooeNoCapability;
        private static OfficeJsErrorCode_ooeActivityLimitReached;
        static buildMessageArrayForIRequestExecutor(customData: string, requestFlags: number, requestMessage: IRequestExecutorRequestMessage, sourceLibHeaderValue: string): Array<any>;
        static buildResponseOnSuccess(responseBody: string, responseHeaders: {
            [name: string]: string;
        }): IRequestExecutorResponseMessage;
        static buildResponseOnError(errorCode: number, message: string): IRequestExecutorResponseMessage;
        static buildRequestMessageSafeArray(customData: string, requestFlags: number, method: string, path: string, headers: {
            [name: string]: string;
        }, body: string): Array<any>;
        static getResponseBody(result: OSF.DDA.RichApi.ExecuteRichApiRequestResult): string;
        static getResponseHeaders(result: OSF.DDA.RichApi.ExecuteRichApiRequestResult): {
            [name: string]: string;
        };
        static getResponseBodyFromSafeArray(data: Array<any>): string;
        static getResponseHeadersFromSafeArray(data: Array<any>): {
            [name: string]: string;
        };
        static getResponseStatusCode(result: OSF.DDA.RichApi.ExecuteRichApiRequestResult): number;
        static getResponseStatusCodeFromSafeArray(data: Array<any>): number;
    }
}
declare module OfficeExtension {
    interface ITestSettings {
    }
    var _testSettings: ITestSettings;
}
declare module OfficeExtension {
    class Utility {
        static checkArgumentNull(value: any, name: string): void;
        static isNullOrUndefined(value: any): boolean;
        static isUndefined(value: any): boolean;
        static isNullOrEmptyString(value: string): boolean;
        static isPlainJsonObject(value: any): boolean;
        static trim(str: string): string;
        static caseInsensitiveCompareString(str1: string, str2: string): boolean;
        static adjustToDateTime(value: any): any;
        static isReadonlyRestRequest(method: string): boolean;
        static setMethodArguments(context: ClientRequestContext, argumentInfo: ArgumentInfo, args: Array<any>): Array<ObjectPath>;
        private static collectObjectPathInfos(context, args, referencedObjectPaths, referencedObjectPathIds);
        static fixObjectPathIfNecessary(clientObject: ClientObject, value: Object): void;
        static validateObjectPath(clientObject: ClientObject): void;
        static validateReferencedObjectPaths(objectPaths: Array<ObjectPath>): void;
        static validateContext(context: ClientRequestContext, obj: ClientObject): void;
        private static _logEnabled;
        static log(message: string): void;
        static load(clientObj: ClientObject, option?: string | LoadOption): void;
        static _parseSelectExpand(select: string): Array<string>;
        private static s_underscoreCharCode;
        static throwError(resourceId: string, arg?: string, errorLocation?: string): void;
        static createRuntimeError(code: string, message: string, location: string): Error;
        static createInvalidArgumentException(name: string, errorLocation?: string): Error;
        static _getResourceString(resourceId: string, arg?: string | Array<string>): string;
        static _formatString(format: string, arrArg: Array<string>): string;
        static throwIfNotLoaded(propertyName: string, fieldValue: any, entityName?: string, isNull?: boolean): void;
        static getObjectPathExpression(objectPath: ObjectPath): string;
        static _createPromiseFromResult<T>(value: T): IPromise<T>;
        static _createTimeoutPromise(timeout: number): IPromise<any>;
        static promisify<T>(action: (callback) => void): OfficeExtension.IPromise<T>;
        static _addActionResultHandler(clientObj: ClientObject, action: Action, resultHandler: IResultHandler): void;
        static _handleNavigationPropertyResults(clientObj: ClientObject, objectValue: any, propertyNames: Array<string>): void;
        private static normalizeName(name);
        static _isLocalDocumentUrl(url: string): boolean;
        static _parseHttpResponseHeaders(allResponseHeaders: string): {
            [name: string]: string;
        };
    }
}

declare module Excel {
    interface ThreeArrowsSet {
        [index: number]: Icon;
        redDownArrow: Icon;
        yellowSideArrow: Icon;
        greenUpArrow: Icon;
    }
    interface ThreeArrowsGraySet {
        [index: number]: Icon;
        grayDownArrow: Icon;
        graySideArrow: Icon;
        grayUpArrow: Icon;
    }
    interface ThreeFlagsSet {
        [index: number]: Icon;
        redFlag: Icon;
        yellowFlag: Icon;
        greenFlag: Icon;
    }
    interface ThreeTrafficLights1Set {
        [index: number]: Icon;
        redCircleWithBorder: Icon;
        yellowCircle: Icon;
        greenCircle: Icon;
    }
    interface ThreeTrafficLights2Set {
        [index: number]: Icon;
        redTrafficLight: Icon;
        yellowTrafficLight: Icon;
        greenTrafficLight: Icon;
    }
    interface ThreeSignsSet {
        [index: number]: Icon;
        redDiamond: Icon;
        yellowTriangle: Icon;
        greenCircle: Icon;
    }
    interface ThreeSymbolsSet {
        [index: number]: Icon;
        redCrossSymbol: Icon;
        yellowExclamationSymbol: Icon;
        greenCheckSymbol: Icon;
    }
    interface ThreeSymbols2Set {
        [index: number]: Icon;
        redCross: Icon;
        yellowExclamation: Icon;
        greenCheck: Icon;
    }
    interface FourArrowsSet {
        [index: number]: Icon;
        redDownArrow: Icon;
        yellowDownInclineArrow: Icon;
        yellowUpInclineArrow: Icon;
        greenUpArrow: Icon;
    }
    interface FourArrowsGraySet {
        [index: number]: Icon;
        grayDownArrow: Icon;
        grayDownInclineArrow: Icon;
        grayUpInclineArrow: Icon;
        grayUpArrow: Icon;
    }
    interface FourRedToBlackSet {
        [index: number]: Icon;
        blackCircle: Icon;
        grayCircle: Icon;
        pinkCircle: Icon;
        redCircle: Icon;
    }
    interface FourRatingSet {
        [index: number]: Icon;
        oneBar: Icon;
        twoBars: Icon;
        threeBars: Icon;
        fourBars: Icon;
    }
    interface FourTrafficLightsSet {
        [index: number]: Icon;
        blackCircleWithBorder: Icon;
        redCircleWithBorder: Icon;
        yellowCircle: Icon;
        greenCircle: Icon;
    }
    interface FiveArrowsSet {
        [index: number]: Icon;
        redDownArrow: Icon;
        yellowDownInclineArrow: Icon;
        yellowSideArrow: Icon;
        yellowUpInclineArrow: Icon;
        greenUpArrow: Icon;
    }
    interface FiveArrowsGraySet {
        [index: number]: Icon;
        grayDownArrow: Icon;
        grayDownInclineArrow: Icon;
        graySideArrow: Icon;
        grayUpInclineArrow: Icon;
        grayUpArrow: Icon;
    }
    interface FiveRatingSet {
        [index: number]: Icon;
        noBars: Icon;
        oneBar: Icon;
        twoBars: Icon;
        threeBars: Icon;
        fourBars: Icon;
    }
    interface FiveQuartersSet {
        [index: number]: Icon;
        whiteCircleAllWhiteQuarters: Icon;
        circleWithThreeWhiteQuarters: Icon;
        circleWithTwoWhiteQuarters: Icon;
        circleWithOneWhiteQuarter: Icon;
        blackCircle: Icon;
    }
    interface ThreeStarsSet {
        [index: number]: Icon;
        silverStar: Icon;
        halfGoldStar: Icon;
        goldStar: Icon;
    }
    interface ThreeTrianglesSet {
        [index: number]: Icon;
        redDownTriangle: Icon;
        yellowDash: Icon;
        greenUpTriangle: Icon;
    }
    interface FiveBoxesSet {
        [index: number]: Icon;
        noFilledBoxes: Icon;
        oneFilledBox: Icon;
        twoFilledBoxes: Icon;
        threeFilledBoxes: Icon;
        fourFilledBoxes: Icon;
    }
    interface IconCollections {
        threeArrows: ThreeArrowsSet;
        threeArrowsGray: ThreeArrowsGraySet;
        threeFlags: ThreeFlagsSet;
        threeTrafficLights1: ThreeTrafficLights1Set;
        threeTrafficLights2: ThreeTrafficLights2Set;
        threeSigns: ThreeSignsSet;
        threeSymbols: ThreeSymbolsSet;
        threeSymbols2: ThreeSymbols2Set;
        fourArrows: FourArrowsSet;
        fourArrowsGray: FourArrowsGraySet;
        fourRedToBlack: FourRedToBlackSet;
        fourRating: FourRatingSet;
        fourTrafficLights: FourTrafficLightsSet;
        fiveArrows: FiveArrowsSet;
        fiveArrowsGray: FiveArrowsGraySet;
        fiveRating: FiveRatingSet;
        fiveQuarters: FiveQuartersSet;
        threeStars: ThreeStarsSet;
        threeTriangles: ThreeTrianglesSet;
        fiveBoxes: FiveBoxesSet;
    }
    var icons: IconCollections;
    /**
     *
     * Provides information about the binding that raised the SelectionChanged event.
     *
     * [Api set: ExcelApi 1.2]
     */
    interface BindingSelectionChangedEventArgs {
        /**
         *
         * Gets the Binding object that represents the binding that raised the SelectionChanged event.
         *
         * [Api set: ExcelApi 1.2]
         */
        binding: Excel.Binding;
        /**
         *
         * Gets the number of columns selected.
         *
         * [Api set: ExcelApi 1.2]
         */
        columnCount: number;
        /**
         *
         * Gets the number of rows selected.
         *
         * [Api set: ExcelApi 1.2]
         */
        rowCount: number;
        /**
         *
         * Gets the index of the first column of the selection (zero-based).
         *
         * [Api set: ExcelApi 1.2]
         */
        startColumn: number;
        /**
         *
         * Gets the index of the first row of the selection (zero-based).
         *
         * [Api set: ExcelApi 1.2]
         */
        startRow: number;
    }
    /**
     *
     * Provides information about the binding that raised the DataChanged event.
     *
     * [Api set: ExcelApi 1.1]
     */
    interface BindingDataChangedEventArgs {
        /**
         *
         * Gets the Binding object that represents the binding that raised the DataChanged event.
         *
         * [Api set: ExcelApi 1.2]
         */
        binding: Excel.Binding;
    }
    /**
     *
     * Provides information about the document that raised the SelectionChanged event.
     *
     * [Api set: ExcelApi 1.1]
     */
    interface SelectionChangedEventArgs {
        /**
         *
         * Gets the workbook object that raised the SelectionChanged event.
         *
         * [Api set: ExcelApi 1.2]
         */
        workbook: Excel.Workbook;
    }
    /**
     *
     * Provides information about the setting that raised the SettingsChanged event
     *
     * [Api set: ExcelApi 1.3]
     */
    interface SettingsChangedEventArgs {
        /**
         *
         * Gets the Setting object that represents the binding that raised the SettingsChanged event
         *
         * [Api set: ExcelApi 1.3]
         */
        settings: Excel.SettingCollection;
    }
    /**
     *
     * Represents the Excel application that manages the workbook.
     *
     * [Api set: ExcelApi 1.1]
     */
    class Application extends OfficeExtension.ClientObject {
        private m_calculationMode;
        /**
         *
         * Returns the calculation mode used in the workbook. See Excel.CalculationMode for details. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        calculationMode: string;
        /**
         *
         * Recalculate all currently opened workbooks in Excel.
         *
         * @param calculationType Specifies the calculation type to use. See Excel.CalculationType for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        calculate(calculationType: string): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.Application;
    }
    /**
     *
     * Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc.
     *
     * [Api set: ExcelApi 1.1]
     */
    class Workbook extends OfficeExtension.ClientObject {
        private m_application;
        private m_bindings;
        private m_functions;
        private m_names;
        private m_pivotTables;
        private m_settings;
        private m_tables;
        private m_worksheets;
        private m__V1Api;
        private m_selectionChanged;
        /**
         *
         * Represents Excel application instance that contains this workbook. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        application: Excel.Application;
        /**
         *
         * Represents a collection of bindings that are part of the workbook. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        bindings: Excel.BindingCollection;
        /**
         *
         * Represents Excel application instance that contains this workbook. Read-only.
         *
         * [Api set: ExcelApi 1.2]
         */
        functions: Excel.Functions;
        /**
         *
         * Represents a collection of workbook scoped named items (named ranges and constants). Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        names: Excel.NamedItemCollection;
        /**
         *
         * Represents a collection of PivotTables associated with the workbook. Read-only.
         *
         * [Api set: ExcelApi 1.3]
         */
        pivotTables: Excel.PivotTableCollection;
        /**
         *
         * Represents a collection of Settings associated with the workbook. Read-only.
         *
         * [Api set: ExcelApi 1.3]
         */
        settings: Excel.SettingCollection;
        /**
         *
         * Represents a collection of tables associated with the workbook. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        tables: Excel.TableCollection;
        /**
         *
         * Represents a collection of worksheets associated with the workbook. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        worksheets: Excel.WorksheetCollection;
        _V1Api: Excel._V1Api;
        /**
         *
         * Gets the currently selected range from the workbook.
         *
         * [Api set: ExcelApi 1.1]
         */
        getSelectedRange(): Excel.Range;
        _GetObjectByReferenceId(bstrReferenceId: string): OfficeExtension.ClientResult<any>;
        _GetObjectTypeNameByReferenceId(bstrReferenceId: string): OfficeExtension.ClientResult<string>;
        _GetReferenceCount(): OfficeExtension.ClientResult<number>;
        _RemoveAllReferences(): void;
        _RemoveReference(bstrReferenceId: string): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.Workbook;
        /**
         *
         * Occurs when the selection in the document is changed.
         *
         * [Api set: ExcelApi 1.2]
         */
        onSelectionChanged: OfficeExtension.EventHandlers<Excel.SelectionChangedEventArgs>;
    }
    /**
     *
     * An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.
     *
     * [Api set: ExcelApi 1.1]
     */
    class Worksheet extends OfficeExtension.ClientObject {
        private m_charts;
        private m_id;
        private m_name;
        private m_pivotTables;
        private m_position;
        private m_protection;
        private m_tables;
        private m_visibility;
        /**
         *
         * Returns collection of charts that are part of the worksheet. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        charts: Excel.ChartCollection;
        /**
         *
         * Collection of PivotTables that are part of the worksheet. Read-only.
         *
         * [Api set: ExcelApi 1.3]
         */
        pivotTables: Excel.PivotTableCollection;
        /**
         *
         * Returns sheet protection object for a worksheet.
         *
         * [Api set: ExcelApi 1.2]
         */
        protection: Excel.WorksheetProtection;
        /**
         *
         * Collection of tables that are part of the worksheet. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        tables: Excel.TableCollection;
        /**
         *
         * Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        id: string;
        /**
         *
         * The display name of the worksheet.
         *
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /**
         *
         * The zero-based position of the worksheet within the workbook.
         *
         * [Api set: ExcelApi 1.1]
         */
        position: number;
        /**
         *
         * The Visibility of the worksheet.
         *
         * [Api set: ExcelApi 1.1 for reading visibility; 1.2 for setting it.]
         */
        visibility: string;
        /**
         *
         * Activate the worksheet in the Excel UI.
         *
         * [Api set: ExcelApi 1.1]
         */
        activate(): void;
        /**
         *
         * Deletes the worksheet from the workbook.
         *
         * [Api set: ExcelApi 1.1]
         */
        delete(): void;
        /**
         *
         * Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid.
         *
         * @param row The row number of the cell to be retrieved. Zero-indexed.
         * @param column the column number of the cell to be retrieved. Zero-indexed.
         *
         * [Api set: ExcelApi 1.1]
         */
        getCell(row: number, column: number): Excel.Range;
        /**
         *
         * Gets the range object specified by the address or name.
         *
         * @param address The address or the name of the range. If not specified, the entire worksheet range is returned.
         *
         * [Api set: ExcelApi 1.1]
         */
        getRange(address?: string): Excel.Range;
        /**
         *
         * The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the worksheet is blank, this function will return the top left cell.
         *
         * @param valuesOnly Considers only cells with values as used cells (ignores formatting). [Api set: ExcelApi 1.2]
         *
         * [Api set: ExcelApi 1.1]
         */
        getUsedRange(valuesOnly?: boolean): Excel.Range;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.Worksheet;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
    }
    /**
     *
     * Represents a collection of worksheet objects that are part of the workbook.
     *
     * [Api set: ExcelApi 1.1]
     */
    class WorksheetCollection extends OfficeExtension.ClientObject {
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.Worksheet>;
        /**
         *
         * Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call ".activate() on it.
         *
         * @param name The name of the worksheet to be added. If specified, name should be unqiue. If not specified, Excel determines the name of the new worksheet.
         *
         * [Api set: ExcelApi 1.1]
         */
        add(name?: string): Excel.Worksheet;
        /**
         *
         * Gets the currently active worksheet in the workbook.
         *
         * [Api set: ExcelApi 1.1]
         */
        getActiveWorksheet(): Excel.Worksheet;
        /**
         *
         * Gets a worksheet object using its Name or ID.
         *
         * @param key The Name or ID of the worksheet.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItem(key: string): Excel.Worksheet;
        /**
         *
         * Gets a worksheet object using its Name or ID. If the worksheet does not exist, the returned object's isNull property will be true.
         *
         * @param key The Name or ID of the worksheet.
         *
         * [Api set: ExcelApi 1.3]
         */
        getItemOrNull(key: string): Excel.Worksheet;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.WorksheetCollection;
    }
    /**
     *
     * Represents the protection of a sheet object.
     *
     * [Api set: ExcelApi 1.2]
     */
    class WorksheetProtection extends OfficeExtension.ClientObject {
        private m_options;
        private m_protected;
        /**
         *
         * Sheet protection options.
         *
         * [Api set: ExcelApi 1.2]
         */
        options: Excel.WorksheetProtectionOptions;
        /**
         *
         * Indicates if the worksheet is protected. Read-Only.
         *
         * [Api set: ExcelApi 1.2]
         */
        protected: boolean;
        /**
         *
         * Protects a worksheet. Fails if the worksheet has been protected.
         *
         * @param options sheet protection options.
         *
         * [Api set: ExcelApi 1.2]
         */
        protect(options?: Excel.WorksheetProtectionOptions): void;
        /**
         *
         * Unprotects a worksheet.
         *
         * [Api set: ExcelApi 1.2]
         */
        unprotect(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.WorksheetProtection;
    }
    /**
     *
     * Represents the options in sheet protection.
     *
     * [Api set: ExcelApi 1.2]
     */
    interface WorksheetProtectionOptions {
        /**
         *
         * Represents the worksheet protection option of allowing using auto filter feature.
         *
         * [Api set: ExcelApi 1.2]
         */
        allowAutoFilter?: boolean;
        /**
         *
         * Represents the worksheet protection option of allowing deleting columns.
         *
         * [Api set: ExcelApi 1.2]
         */
        allowDeleteColumns?: boolean;
        /**
         *
         * Represents the worksheet protection option of allowing deleting rows.
         *
         * [Api set: ExcelApi 1.2]
         */
        allowDeleteRows?: boolean;
        /**
         *
         * Represents the worksheet protection option of allowing formatting cells.
         *
         * [Api set: ExcelApi 1.2]
         */
        allowFormatCells?: boolean;
        /**
         *
         * Represents the worksheet protection option of allowing formatting columns.
         *
         * [Api set: ExcelApi 1.2]
         */
        allowFormatColumns?: boolean;
        /**
         *
         * Represents the worksheet protection option of allowing formatting rows.
         *
         * [Api set: ExcelApi 1.2]
         */
        allowFormatRows?: boolean;
        /**
         *
         * Represents the worksheet protection option of allowing inserting columns.
         *
         * [Api set: ExcelApi 1.2]
         */
        allowInsertColumns?: boolean;
        /**
         *
         * Represents the worksheet protection option of allowing inserting hyperlinks.
         *
         * [Api set: ExcelApi 1.2]
         */
        allowInsertHyperlinks?: boolean;
        /**
         *
         * Represents the worksheet protection option of allowing inserting rows.
         *
         * [Api set: ExcelApi 1.2]
         */
        allowInsertRows?: boolean;
        /**
         *
         * Represents the worksheet protection option of allowing using PivotTable feature.
         *
         * [Api set: ExcelApi 1.2]
         */
        allowPivotTables?: boolean;
        /**
         *
         * Represents the worksheet protection option of allowing using sort feature.
         *
         * [Api set: ExcelApi 1.2]
         */
        allowSort?: boolean;
    }
    /**
     *
     * Range represents a set of one or more contiguous cells such as a cell, a row, a column, block of cells, etc.
     *
     * [Api set: ExcelApi 1.1]
     */
    class Range extends OfficeExtension.ClientObject {
        private m_address;
        private m_addressLocal;
        private m_cellCount;
        private m_columnCount;
        private m_columnHidden;
        private m_columnIndex;
        private m_conditionalFormats;
        private m_format;
        private m_formulas;
        private m_formulasLocal;
        private m_formulasR1C1;
        private m_hidden;
        private m_numberFormat;
        private m_rowCount;
        private m_rowHidden;
        private m_rowIndex;
        private m_sort;
        private m_text;
        private m_valueTypes;
        private m_values;
        private m_worksheet;
        private m__ReferenceId;
        private _ensureInteger(num, methodName);
        private _getAdjacentRange(functionName, count, referenceRange, rowDirection, columnDirection);
        /**
         *
         * Collection of ConditionalFormats that intersect the range. Read-only.
         *
         * [Api set: ExcelApi 1.4]
         */
        conditionalFormats: Excel.ConditionalFormatCollection;
        /**
         *
         * Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        format: Excel.RangeFormat;
        /**
         *
         * Represents the range sort of the current range.
         *
         * [Api set: ExcelApi 1.2]
         */
        sort: Excel.RangeSort;
        /**
         *
         * The worksheet containing the current range. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        worksheet: Excel.Worksheet;
        /**
         *
         * Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. Sheet1!A1:B4). Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        address: string;
        /**
         *
         * Represents range reference for the specified range in the language of the user. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        addressLocal: string;
        /**
         *
         * Number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        cellCount: number;
        /**
         *
         * Represents the total number of columns in the range. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        columnCount: number;
        /**
         *
         * Represents if all columns of the current range are hidden.
         *
         * [Api set: ExcelApi 1.2]
         */
        columnHidden: boolean;
        /**
         *
         * Represents the column number of the first cell in the range. Zero-indexed. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        columnIndex: number;
        /**
         *
         * Represents the formula in A1-style notation.
         *
         * [Api set: ExcelApi 1.1]
         */
        formulas: Array<Array<any>>;
        /**
         *
         * Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.
         *
         * [Api set: ExcelApi 1.1]
         */
        formulasLocal: Array<Array<any>>;
        /**
         *
         * Represents the formula in R1C1-style notation.
         *
         * [Api set: ExcelApi 1.2]
         */
        formulasR1C1: Array<Array<any>>;
        /**
         *
         * Represents if all cells of the current range are hidden.
         *
         * [Api set: ExcelApi 1.2]
         */
        hidden: boolean;
        /**
         *
         * Represents Excel's number format code for the given cell.
         *
         * [Api set: ExcelApi 1.1]
         */
        numberFormat: Array<Array<any>>;
        /**
         *
         * Returns the total number of rows in the range. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        rowCount: number;
        /**
         *
         * Represents if all rows of the current range are hidden.
         *
         * [Api set: ExcelApi 1.2]
         */
        rowHidden: boolean;
        /**
         *
         * Returns the row number of the first cell in the range. Zero-indexed. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        rowIndex: number;
        /**
         *
         * Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        text: Array<Array<any>>;
        /**
         *
         * Represents the type of data of each cell. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        valueTypes: Array<Array<string>>;
        /**
         *
         * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
         *
         * [Api set: ExcelApi 1.1]
         */
        values: Array<Array<any>>;
        _ReferenceId: string;
        /**
         *
         * Clear range values, format, fill, border, etc.
         *
         * @param applyTo Determines the type of clear action. See Excel.ClearApplyTo for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        clear(applyTo?: string): void;
        /**
         *
         * Deletes the cells associated with the range.
         *
         * @param shift Specifies which way to shift the cells. See Excel.DeleteShiftDirection for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        delete(shift: string): void;
        /**
         *
         * Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E16".
         *
         * @param anotherRange The range object or address or range name.
         *
         * [Api set: ExcelApi 1.1]
         */
        getBoundingRect(anotherRange: Excel.Range | string): Excel.Range;
        /**
         *
         * Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.
         *
         * @param row Row number of the cell to be retrieved. Zero-indexed.
         * @param column Column number of the cell to be retrieved. Zero-indexed.
         *
         * [Api set: ExcelApi 1.1]
         */
        getCell(row: number, column: number): Excel.Range;
        /**
         *
         * Gets a column contained in the range.
         *
         * @param column Column number of the range to be retrieved. Zero-indexed.
         *
         * [Api set: ExcelApi 1.1]
         */
        getColumn(column: number): Excel.Range;
        /**
         *
         * Gets a certain number of columns to the right of the current Range object.
         *
         * @param count The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        getColumnsAfter(count?: number): Excel.Range;
        /**
         *
         * Gets a certain number of columns to the left of the current Range object.
         *
         * @param count The number of columns to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        getColumnsBefore(count?: number): Excel.Range;
        /**
         *
         * Gets an object that represents the entire column of the range.
         *
         * [Api set: ExcelApi 1.1]
         */
        getEntireColumn(): Excel.Range;
        /**
         *
         * Gets an object that represents the entire row of the range.
         *
         * [Api set: ExcelApi 1.1]
         */
        getEntireRow(): Excel.Range;
        /**
         *
         * Gets the range object that represents the rectangular intersection of the given ranges.
         *
         * @param anotherRange The range object or range address that will be used to determine the intersection of ranges.
         *
         * [Api set: ExcelApi 1.1]
         */
        getIntersection(anotherRange: Excel.Range | string): Excel.Range;
        /**
         *
         * Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.
         *
         * @param anotherRange The range object or range address that will be used to determine the intersection of ranges.
         *
         * [Api set: ExcelApi 1.3]
         */
        getIntersectionOrNull(anotherRange: Excel.Range | string): Excel.Range;
        /**
         *
         * Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".
         *
         * [Api set: ExcelApi 1.1]
         */
        getLastCell(): Excel.Range;
        /**
         *
         * Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".
         *
         * [Api set: ExcelApi 1.1]
         */
        getLastColumn(): Excel.Range;
        /**
         *
         * Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".
         *
         * [Api set: ExcelApi 1.1]
         */
        getLastRow(): Excel.Range;
        /**
         *
         * Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an exception will be thrown.
         *
         * @param rowOffset The number of rows (positive, negative, or 0) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward.
         * @param columnOffset The number of columns (positive, negative, or 0) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left.
         *
         * [Api set: ExcelApi 1.1]
         */
        getOffsetRange(rowOffset: number, columnOffset: number): Excel.Range;
        /**
         *
         * Gets a Range object similar to the current Range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.
         *
         * @param deltaRows The number of rows by which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.
         * @param deltaColumns The number of columnsby which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.
         *
         * [Api set: ExcelApi 1.2]
         */
        getResizedRange(deltaRows: number, deltaColumns: number): Excel.Range;
        /**
         *
         * Gets a row contained in the range.
         *
         * @param row Row number of the range to be retrieved. Zero-indexed.
         *
         * [Api set: ExcelApi 1.1]
         */
        getRow(row: number): Excel.Range;
        /**
         *
         * Gets a certain number of rows above the current Range object.
         *
         * @param count The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        getRowsAbove(count?: number): Excel.Range;
        /**
         *
         * Gets a certain number of rows below the current Range object.
         *
         * @param count The number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        getRowsBelow(count?: number): Excel.Range;
        /**
         *
         * Returns the used range of the given range object.
         *
         * @param valuesOnly Considers only cells with values as used cells. [Api set: ExcelApi 1.2]
         *
         * [Api set: ExcelApi 1.1]
         */
        getUsedRange(valuesOnly?: boolean): Excel.Range;
        /**
         *
         * Represents the visible rows of the current range.
         *
         * [Api set: ExcelApi 1.3]
         */
        getVisibleView(): Excel.RangeView;
        /**
         *
         * Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.
         *
         * @param shift Specifies which way to shift the cells. See Excel.InsertShiftDirection for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        insert(shift: string): Excel.Range;
        /**
         *
         * Merge the range cells into one region in the worksheet.
         *
         * @param across Set true to merge cells in each row of the specified range as separate merged cells. The default value is false.
         *
         * [Api set: ExcelApi 1.2]
         */
        merge(across?: boolean): void;
        /**
         *
         * Selects the specified range in the Excel UI.
         *
         * [Api set: ExcelApi 1.1]
         */
        select(): void;
        /**
         *
         * Unmerge the range cells into separate cells.
         *
         * [Api set: ExcelApi 1.2]
         */
        unmerge(): void;
        _KeepReference(): void;
        _ValidateArraySize(rows: number, columns: number): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.Range;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
    }
    /**
     *
     * Represents a string reference of the form SheetName!A1:B5, or a global or local named range
     *
     * [Api set: ExcelApi 1.2]
     */
    interface RangeReference {
        address: string;
    }
    /**
     *
     * RangeView represents a set of visible cells of the parent range.
     *
     * [Api set: ExcelApi 1.3]
     */
    class RangeView extends OfficeExtension.ClientObject {
        private m_cellAddresses;
        private m_columnCount;
        private m_formulas;
        private m_formulasLocal;
        private m_formulasR1C1;
        private m_index;
        private m_numberFormat;
        private m_rowCount;
        private m_rows;
        private m_text;
        private m_valueTypes;
        private m_values;
        /**
         *
         * Represents a collection of range views associated with the range. Read-only.
         *
         * [Api set: ExcelApi 1.3]
         */
        rows: Excel.RangeViewCollection;
        /**
         *
         * Represents the cell addresses of the RangeView.
         *
         * [Api set: ExcelApi 1.3]
         */
        cellAddresses: Array<Array<any>>;
        /**
         *
         * Returns the number of visible columns. Read-only.
         *
         * [Api set: ExcelApi 1.3]
         */
        columnCount: number;
        /**
         *
         * Represents the formula in A1-style notation.
         *
         * [Api set: ExcelApi 1.3]
         */
        formulas: Array<Array<any>>;
        /**
         *
         * Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.
         *
         * [Api set: ExcelApi 1.3]
         */
        formulasLocal: Array<Array<any>>;
        /**
         *
         * Represents the formula in R1C1-style notation.
         *
         * [Api set: ExcelApi 1.3]
         */
        formulasR1C1: Array<Array<any>>;
        /**
         *
         * Returns a value that represents the index of the RangeView. Read-only.
         *
         * [Api set: ExcelApi 1.3]
         */
        index: number;
        /**
         *
         * Represents Excel's number format code for the given cell.
         *
         * [Api set: ExcelApi 1.3]
         */
        numberFormat: Array<Array<any>>;
        /**
         *
         * Returns the number of visible rows. Read-only.
         *
         * [Api set: ExcelApi 1.3]
         */
        rowCount: number;
        /**
         *
         * Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.
         *
         * [Api set: ExcelApi 1.3]
         */
        text: Array<Array<any>>;
        /**
         *
         * Represents the type of data of each cell. Read-only.
         *
         * [Api set: ExcelApi 1.3]
         */
        valueTypes: Array<Array<string>>;
        /**
         *
         * Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
         *
         * [Api set: ExcelApi 1.3]
         */
        values: Array<Array<any>>;
        /**
         *
         * Gets the parent range associated with the current RangeView.
         *
         * [Api set: ExcelApi 1.3]
         */
        getRange(): Excel.Range;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.RangeView;
    }
    /**
     *
     * Represents a collection of worksheet objects that are part of the workbook.
     *
     * [Api set: ExcelApi 1.3]
     */
    class RangeViewCollection extends OfficeExtension.ClientObject {
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.RangeView>;
        /**
         *
         * Gets a RangeView Row via it's index. Zero-Indexed.
         *
         * @param index Index of the visible row.
         *
         * [Api set: ExcelApi 1.3]
         */
        getItemAt(index: number): Excel.RangeView;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.RangeViewCollection;
    }
    /**
     *
     * Represents a collection of worksheet objects that are part of the workbook.
     *
     * [Api set: ExcelApi 1.3]
     */
    class SettingCollection extends OfficeExtension.ClientObject {
        private m_settingsChanged;
        private m__items;
        /**
         *
         * Sets or adds the specified setting to the workbook.
         *
         * @param key The Key of the new setting.
         * @param value The Value for the new setting.
         *
         * [Api set: ExcelApi 1.3]
         */
        add(key: string, value: string | number | boolean | Array<any> | any): Excel.Setting;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.Setting>;
        /**
         *
         * Gets a Setting entry via the key.
         *
         * @param key Key of the setting.
         *
         * [Api set: ExcelApi 1.3]
         */
        getItem(key: string): Excel.Setting;
        /**
         *
         * Gets a Setting entry via the key. If the Setting does not exist, the returned object's isNull property will be true.
         *
         * @param key The key of the setting.
         *
         * [Api set: ExcelApi 1.3]
         */
        getItemOrNull(key: string): Excel.Setting;
        /**
         *
         * Sets or adds the specified setting to the workbook.
         *
         * @param key The Key of the new setting.
         * @param value The Value for the new setting.
         *
         *
         */
        _Add(key: string, value: string): Excel.Setting;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.SettingCollection;
        /**
         *
         * Occurs when the Settings in the document are changed.
         *
         * [Api set: ExcelApi 1.3]
         */
        onSettingsChanged: OfficeExtension.EventHandlers<Excel.SettingsChangedEventArgs>;
    }
    /**
     *
     * Setting represents a key-value pair of a setting persisted to the document.
     *
     * [Api set: ExcelApi 1.3]
     */
    class Setting extends OfficeExtension.ClientObject {
        private m_key;
        private m__Value;
        /**
         *
         * Represents the value stored for this setting.
         *
         * [Api set: ExcelApi 1.3]
         */
        value: any;
        /**
         *
         * Returns the key that represents the id of the Setting. Read-only.
         *
         * [Api set: ExcelApi 1.3]
         */
        key: string;
        /**
         *
         * Represents the value stored for this setting.
         *
         * [Api set: ExcelApi 1.3]
         */
        _Value: string;
        /**
         *
         * Deletes the setting.
         *
         * [Api set: ExcelApi 1.3]
         */
        delete(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.Setting;
    }
    /**
     *
     * A collection of all the nameditem objects that are part of the workbook.
     *
     * [Api set: ExcelApi 1.1]
     */
    class NamedItemCollection extends OfficeExtension.ClientObject {
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.NamedItem>;
        /**
         *
         * Gets a nameditem object using its name
         *
         * @param name nameditem name.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItem(name: string): Excel.NamedItem;
        /**
         *
         * Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.
         *
         * @param name nameditem name.
         *
         * [Api set: ExcelApi 1.3]
         */
        getItemOrNull(name: string): Excel.NamedItem;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.NamedItemCollection;
    }
    /**
     *
     * Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, reference to a range. This object can be used to obtain range object associated with names.
     *
     * [Api set: ExcelApi 1.1]
     */
    class NamedItem extends OfficeExtension.ClientObject {
        private m_name;
        private m_type;
        private m_value;
        private m_visible;
        private m__Id;
        /**
         *
         * The name of the object. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /**
         *
         * Indicates what type of reference is associated with the name. See Excel.NamedItemType for details. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        type: string;
        /**
         *
         * Represents the formula that the name is defined to refer to. E.g. =Sheet14!$B$2:$H$12, =4.75, etc. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        value: any;
        /**
         *
         * Specifies whether the object is visible or not.
         *
         * [Api set: ExcelApi 1.1]
         */
        visible: boolean;
        _Id: string;
        /**
         *
         * Returns the range object that is associated with the name. Throws an exception if the named item's type is not a range.
         *
         * [Api set: ExcelApi 1.1]
         */
        getRange(): Excel.Range;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.NamedItem;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
    }
    /**
     *
     * Represents an Office.js binding that is defined in the workbook.
     *
     * [Api set: ExcelApi 1.1]
     */
    class Binding extends OfficeExtension.ClientObject {
        private m_id;
        private m_type;
        private m_dataChanged;
        private m_selectionChanged;
        /**
         *
         * Represents binding identifier. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        id: string;
        /**
         *
         * Returns the type of the binding. See Excel.BindingType for details. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        type: string;
        /**
         *
         * Deletes the binding.
         *
         * [Api set: ExcelApi 1.3]
         */
        delete(): void;
        /**
         *
         * Returns the range represented by the binding. Will throw an error if binding is not of the correct type.
         *
         * [Api set: ExcelApi 1.1]
         */
        getRange(): Excel.Range;
        /**
         *
         * Returns the table represented by the binding. Will throw an error if binding is not of the correct type.
         *
         * [Api set: ExcelApi 1.1]
         */
        getTable(): Excel.Table;
        /**
         *
         * Returns the text represented by the binding. Will throw an error if binding is not of the correct type.
         *
         * [Api set: ExcelApi 1.1]
         */
        getText(): OfficeExtension.ClientResult<string>;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.Binding;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
        /**
         *
         * Occurs when data within the binding is changed.
         *
         * [Api set: ExcelApi 1.2]
         */
        onDataChanged: OfficeExtension.EventHandlers<Excel.BindingDataChangedEventArgs>;
        /**
         *
         * Occurs when the selection is changed within the binding.
         *
         * [Api set: ExcelApi 1.2]
         */
        onSelectionChanged: OfficeExtension.EventHandlers<Excel.BindingSelectionChangedEventArgs>;
    }
    /**
     *
     * Represents the collection of all the binding objects that are part of the workbook.
     *
     * [Api set: ExcelApi 1.1]
     */
    class BindingCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.Binding>;
        /**
         *
         * Returns the number of bindings in the collection. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        count: number;
        /**
         *
         * Add a new binding to a particular Range.
         *
         * @param range Range to bind the binding to. May be an Excel Range object, or a string. If string, must contain the full address, including the sheet name
         * @param bindingType Type of binding. See Excel.BindingType.
         * @param id Name of binding.
         *
         * [Api set: ExcelApi 1.3]
         */
        add(range: Excel.Range | string, bindingType: string, id: string): Excel.Binding;
        /**
         *
         * Add a new binding based on a named item in the workbook.
         *
         * @param name Name from which to create binding.
         * @param bindingType Type of binding. See Excel.BindingType.
         * @param id Name of binding.
         *
         * [Api set: ExcelApi 1.3]
         */
        addFromNamedItem(name: string, bindingType: string, id: string): Excel.Binding;
        /**
         *
         * Add a new binding based on the current selection.
         *
         * @param bindingType Type of binding. See Excel.BindingType.
         * @param id Name of binding.
         *
         * [Api set: ExcelApi 1.3]
         */
        addFromSelection(bindingType: string, id: string): Excel.Binding;
        /**
         *
         * Gets a binding object by ID.
         *
         * @param id Id of the binding object to be retrieved.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItem(id: string): Excel.Binding;
        /**
         *
         * Gets a binding object based on its position in the items array.
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItemAt(index: number): Excel.Binding;
        /**
         *
         * Gets a binding object by ID. If the binding object does not exist, the return object's isNull property will be true.
         *
         * @param id Id of the binding object to be retrieved.
         *
         * [Api set: ExcelApi 1.3]
         */
        getItemOrNull(id: string): Excel.Binding;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.BindingCollection;
    }
    /**
     *
     * Represents a collection of all the tables that are part of the workbook.
     *
     * [Api set: ExcelApi 1.1]
     */
    class TableCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.Table>;
        /**
         *
         * Returns the number of tables in the workbook. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        count: number;
        /**
         *
         * Create a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.
         *
         * @param address A Range object, or a string address or name of the range representing the data source. If the address does not contain a sheet name, the currently-active sheet is used. [Api set: ExcelApi 1.1 for string parameter; 1.3 for accepting a Range object as well]
         * @param hasHeaders Boolean value that indicates whether the data being imported has column labels. If the source does not contain headers (i.e,. when this property set to false), Excel will automatically generate header shifting the data down by one row.
         *
         * [Api set: ExcelApi 1.1]
         */
        add(address: Excel.Range | string, hasHeaders: boolean): Excel.Table;
        /**
         *
         * Gets a table by Name or ID.
         *
         * @param key Name or ID of the table to be retrieved.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItem(key: number | string): Excel.Table;
        /**
         *
         * Gets a table based on its position in the collection.
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItemAt(index: number): Excel.Table;
        /**
         *
         * Gets a table by Name or ID. If the table does not exist, the return object's isNull property will be true.
         *
         * @param key Name or ID of the table to be retrieved.
         *
         * [Api set: ExcelApi 1.3]
         */
        getItemOrNull(key: number | string): Excel.Table;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.TableCollection;
    }
    /**
     *
     * Represents an Excel table.
     *
     * [Api set: ExcelApi 1.1]
     */
    class Table extends OfficeExtension.ClientObject {
        private m_columns;
        private m_highlightFirstColumn;
        private m_highlightLastColumn;
        private m_id;
        private m_name;
        private m_rows;
        private m_showBandedColumns;
        private m_showBandedRows;
        private m_showFilterButton;
        private m_showHeaders;
        private m_showTotals;
        private m_sort;
        private m_style;
        private m_worksheet;
        /**
         *
         * Represents a collection of all the columns in the table. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        columns: Excel.TableColumnCollection;
        /**
         *
         * Represents a collection of all the rows in the table. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        rows: Excel.TableRowCollection;
        /**
         *
         * Represents the sorting for the table.
         *
         * [Api set: ExcelApi 1.2]
         */
        sort: Excel.TableSort;
        /**
         *
         * The worksheet containing the current table. Read-only.
         *
         * [Api set: ExcelApi 1.2]
         */
        worksheet: Excel.Worksheet;
        /**
         *
         * Indicates whether the first column contains special formatting.
         *
         * [Api set: ExcelApi 1.3]
         */
        highlightFirstColumn: boolean;
        /**
         *
         * Indicates whether the last column contains special formatting.
         *
         * [Api set: ExcelApi 1.3]
         */
        highlightLastColumn: boolean;
        /**
         *
         * Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        id: number;
        /**
         *
         * Name of the table.
         *
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /**
         *
         * Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.
         *
         * [Api set: ExcelApi 1.3]
         */
        showBandedColumns: boolean;
        /**
         *
         * Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.
         *
         * [Api set: ExcelApi 1.3]
         */
        showBandedRows: boolean;
        /**
         *
         * Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.
         *
         * [Api set: ExcelApi 1.3]
         */
        showFilterButton: boolean;
        /**
         *
         * Indicates whether the header row is visible or not. This value can be set to show or remove the header row.
         *
         * [Api set: ExcelApi 1.1]
         */
        showHeaders: boolean;
        /**
         *
         * Indicates whether the total row is visible or not. This value can be set to show or remove the total row.
         *
         * [Api set: ExcelApi 1.1]
         */
        showTotals: boolean;
        /**
         *
         * Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.
         *
         * [Api set: ExcelApi 1.1]
         */
        style: string;
        /**
         *
         * Clears all the filters currently applied on the table.
         *
         * [Api set: ExcelApi 1.2]
         */
        clearFilters(): void;
        /**
         *
         * Converts the table into a normal range of cells. All data is preserved.
         *
         * [Api set: ExcelApi 1.2]
         */
        convertToRange(): Excel.Range;
        /**
         *
         * Deletes the table.
         *
         * [Api set: ExcelApi 1.1]
         */
        delete(): void;
        /**
         *
         * Gets the range object associated with the data body of the table.
         *
         * [Api set: ExcelApi 1.1]
         */
        getDataBodyRange(): Excel.Range;
        /**
         *
         * Gets the range object associated with header row of the table.
         *
         * [Api set: ExcelApi 1.1]
         */
        getHeaderRowRange(): Excel.Range;
        /**
         *
         * Gets the range object associated with the entire table.
         *
         * [Api set: ExcelApi 1.1]
         */
        getRange(): Excel.Range;
        /**
         *
         * Gets the range object associated with totals row of the table.
         *
         * [Api set: ExcelApi 1.1]
         */
        getTotalRowRange(): Excel.Range;
        /**
         *
         * Reapplies all the filters currently on the table.
         *
         * [Api set: ExcelApi 1.2]
         */
        reapplyFilters(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.Table;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
    }
    /**
     *
     * Represents a collection of all the columns that are part of the table.
     *
     * [Api set: ExcelApi 1.1]
     */
    class TableColumnCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.TableColumn>;
        /**
         *
         * Returns the number of columns in the table. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        count: number;
        /**
         *
         * Adds a new column to the table.
         *
         * @param index Specifies the relative position of the new column. If null or -1, the addition happens at the end. Columns with a higher index will be shifted to the side. Zero-indexed.
         * @param values A 2-dimensional array of unformatted values of the table column.
         *
         * [Api set: ExcelApi 1.1 requires an index smaller than the total column count; 1.4 allows index to be optional (null or -1) and will append a column at the end.]
         */
        add(index?: number, values?: Array<Array<boolean | string | number>> | boolean | string | number): Excel.TableColumn;
        /**
         *
         * Gets a column object by Name or ID.
         *
         * @param key Column Name or ID.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItem(key: number | string): Excel.TableColumn;
        /**
         *
         * Gets a column based on its position in the collection.
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItemAt(index: number): Excel.TableColumn;
        /**
         *
         * Gets a column object by Name or ID. If the column does not exist, the returned object's isNull property will be true.
         *
         * @param key Column Name or ID.
         *
         * [Api set: ExcelApi 1.3]
         */
        getItemOrNull(key: number | string): Excel.TableColumn;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.TableColumnCollection;
    }
    /**
     *
     * Represents a column in a table.
     *
     * [Api set: ExcelApi 1.1]
     */
    class TableColumn extends OfficeExtension.ClientObject {
        private m_filter;
        private m_id;
        private m_index;
        private m_name;
        private m_values;
        /**
         *
         * Retrieve the filter applied to the column.
         *
         * [Api set: ExcelApi 1.2]
         */
        filter: Excel.Filter;
        /**
         *
         * Returns a unique key that identifies the column within the table. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        id: number;
        /**
         *
         * Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        index: number;
        /**
         *
         * Returns the name of the table column. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /**
         *
         * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
         *
         * [Api set: ExcelApi 1.1]
         */
        values: Array<Array<any>>;
        /**
         *
         * Deletes the column from the table.
         *
         * [Api set: ExcelApi 1.1]
         */
        delete(): void;
        /**
         *
         * Gets the range object associated with the data body of the column.
         *
         * [Api set: ExcelApi 1.1]
         */
        getDataBodyRange(): Excel.Range;
        /**
         *
         * Gets the range object associated with the header row of the column.
         *
         * [Api set: ExcelApi 1.1]
         */
        getHeaderRowRange(): Excel.Range;
        /**
         *
         * Gets the range object associated with the entire column.
         *
         * [Api set: ExcelApi 1.1]
         */
        getRange(): Excel.Range;
        /**
         *
         * Gets the range object associated with the totals row of the column.
         *
         * [Api set: ExcelApi 1.1]
         */
        getTotalRowRange(): Excel.Range;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.TableColumn;
        /** Handle identity results returned from the document
         * @private
         */
        _handleIdResult(value: any): void;
    }
    /**
     *
     * Represents a collection of all the rows that are part of the table.
     *
     * [Api set: ExcelApi 1.1]
     */
    class TableRowCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.TableRow>;
        /**
         *
         * Returns the number of rows in the table. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        count: number;
        /**
         *
         * Adds one or more rows to the table. The return object will be the top of the newly added row(s).
         *
         * @param index Specifies the relative position of the new row. If null or -1, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.
         * @param values A 2-dimensional array of unformatted values of the table row.
         *
         * [Api set: ExcelApi 1.1 for adding a single row; 1.4 allows adding of multiple rows.]
         */
        add(index?: number, values?: Array<Array<boolean | string | number>> | boolean | string | number): Excel.TableRow;
        /**
         *
         * Gets a row based on its position in the collection.
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItemAt(index: number): Excel.TableRow;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.TableRowCollection;
    }
    /**
     *
     * Represents a row in a table.
     *
     * [Api set: ExcelApi 1.1]
     */
    class TableRow extends OfficeExtension.ClientObject {
        private m_index;
        private m_values;
        /**
         *
         * Returns the index number of the row within the rows collection of the table. Zero-indexed. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        index: number;
        /**
         *
         * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
         *
         * [Api set: ExcelApi 1.1]
         */
        values: Array<Array<any>>;
        /**
         *
         * Deletes the row from the table.
         *
         * [Api set: ExcelApi 1.1]
         */
        delete(): void;
        /**
         *
         * Returns the range object associated with the entire row.
         *
         * [Api set: ExcelApi 1.1]
         */
        getRange(): Excel.Range;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.TableRow;
    }
    /**
     *
     * A format object encapsulating the range's font, fill, borders, alignment, and other properties.
     *
     * [Api set: ExcelApi 1.1]
     */
    class RangeFormat extends OfficeExtension.ClientObject {
        private m_borders;
        private m_columnWidth;
        private m_fill;
        private m_font;
        private m_horizontalAlignment;
        private m_protection;
        private m_rowHeight;
        private m_verticalAlignment;
        private m_wrapText;
        /**
         *
         * Collection of border objects that apply to the overall range selected Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        borders: Excel.RangeBorderCollection;
        /**
         *
         * Returns the fill object defined on the overall range. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        fill: Excel.RangeFill;
        /**
         *
         * Returns the font object defined on the overall range selected Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        font: Excel.RangeFont;
        /**
         *
         * Returns the format protection object for a range.
         *
         * [Api set: ExcelApi 1.2]
         */
        protection: Excel.FormatProtection;
        /**
         *
         * Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.
         *
         * [Api set: ExcelApi 1.2]
         */
        columnWidth: number;
        /**
         *
         * Represents the horizontal alignment for the specified object. See Excel.HorizontalAlignment for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        horizontalAlignment: string;
        /**
         *
         * Gets or sets the height of all rows in the range. If the row heights are not uniform null will be returned.
         *
         * [Api set: ExcelApi 1.2]
         */
        rowHeight: number;
        /**
         *
         * Represents the vertical alignment for the specified object. See Excel.VerticalAlignment for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        verticalAlignment: string;
        /**
         *
         * Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting
         *
         * [Api set: ExcelApi 1.1]
         */
        wrapText: boolean;
        /**
         *
         * Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.
         *
         * [Api set: ExcelApi 1.2]
         */
        autofitColumns(): void;
        /**
         *
         * Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.
         *
         * [Api set: ExcelApi 1.2]
         */
        autofitRows(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.RangeFormat;
    }
    /**
     *
     * Represents the format protection of a range object.
     *
     * [Api set: ExcelApi 1.2]
     */
    class FormatProtection extends OfficeExtension.ClientObject {
        private m_formulaHidden;
        private m_locked;
        /**
         *
         * Indicates if Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting.
         *
         * [Api set: ExcelApi 1.2]
         */
        formulaHidden: boolean;
        /**
         *
         * Indicates if Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting.
         *
         * [Api set: ExcelApi 1.2]
         */
        locked: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.FormatProtection;
    }
    /**
     *
     * Represents the background of a range object.
     *
     * [Api set: ExcelApi 1.1]
     */
    class RangeFill extends OfficeExtension.ClientObject {
        private m_color;
        /**
         *
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")
         *
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        /**
         *
         * Resets the range background.
         *
         * [Api set: ExcelApi 1.1]
         */
        clear(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.RangeFill;
    }
    /**
     *
     * Represents the border of an object.
     *
     * [Api set: ExcelApi 1.1]
     */
    class RangeBorder extends OfficeExtension.ClientObject {
        private m_color;
        private m_sideIndex;
        private m_style;
        private m_weight;
        /**
         *
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         *
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        /**
         *
         * Constant value that indicates the specific side of the border. See Excel.BorderIndex for details. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        sideIndex: string;
        /**
         *
         * One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        style: string;
        /**
         *
         * Specifies the weight of the border around a range. See Excel.BorderWeight for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        weight: string;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.RangeBorder;
    }
    /**
     *
     * Represents the border objects that make up range border.
     *
     * [Api set: ExcelApi 1.1]
     */
    class RangeBorderCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.RangeBorder>;
        /**
         *
         * Number of border objects in the collection. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        count: number;
        /**
         *
         * Gets a border object using its name
         *
         * @param index Index value of the border object to be retrieved. See Excel.BorderIndex for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItem(index: string): Excel.RangeBorder;
        /**
         *
         * Gets a border object using its index
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItemAt(index: number): Excel.RangeBorder;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.RangeBorderCollection;
    }
    /**
     *
     * This object represents the font attributes (font name, font size, color, etc.) for an object.
     *
     * [Api set: ExcelApi 1.1]
     */
    class RangeFont extends OfficeExtension.ClientObject {
        private m_bold;
        private m_color;
        private m_italic;
        private m_name;
        private m_size;
        private m_underline;
        /**
         *
         * Represents the bold status of font.
         *
         * [Api set: ExcelApi 1.1]
         */
        bold: boolean;
        /**
         *
         * HTML color code representation of the text color. E.g. #FF0000 represents Red.
         *
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        /**
         *
         * Represents the italic status of the font.
         *
         * [Api set: ExcelApi 1.1]
         */
        italic: boolean;
        /**
         *
         * Font name (e.g. "Calibri")
         *
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /**
         *
         * Font size.
         *
         * [Api set: ExcelApi 1.1]
         */
        size: number;
        /**
         *
         * Type of underline applied to the font. See Excel.RangeUnderlineStyle for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        underline: string;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.RangeFont;
    }
    /**
     *
     * A collection of all the chart objects on a worksheet.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.Chart>;
        /**
         *
         * Returns the number of charts in the worksheet. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        count: number;
        /**
         *
         * Creates a new chart.
         *
         * @param type Represents the type of a chart. See Excel.ChartType for details.
         * @param sourceData The Range object corresponding to the source data.
         * @param seriesBy Specifies the way columns or rows are used as data series on the chart. See Excel.ChartSeriesBy for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        add(type: string, sourceData: Excel.Range, seriesBy?: string): Excel.Chart;
        /**
         *
         * Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.
         *
         * @param name Name of the chart to be retrieved.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItem(name: string): Excel.Chart;
        /**
         *
         * Gets a chart based on its position in the collection.
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItemAt(index: number): Excel.Chart;
        /**
         *
         * Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.
            If the chart does not exist, the returned object's isNull property will be true.
         *
         * @param name Name of the chart to be retrieved.
         *
         * [Api set: ExcelApi 1.3]
         */
        getItemOrNull(name: string): Excel.Chart;
        _GetItem(key: string): Excel.Chart;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartCollection;
    }
    /**
     *
     * Represents a chart object in a workbook.
     *
     * [Api set: ExcelApi 1.1]
     */
    class Chart extends OfficeExtension.ClientObject {
        private m_axes;
        private m_dataLabels;
        private m_format;
        private m_height;
        private m_left;
        private m_legend;
        private m_name;
        private m_series;
        private m_title;
        private m_top;
        private m_width;
        private m_worksheet;
        /**
         *
         * Represents chart axes. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        axes: Excel.ChartAxes;
        /**
         *
         * Represents the datalabels on the chart. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        dataLabels: Excel.ChartDataLabels;
        /**
         *
         * Encapsulates the format properties for the chart area. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        format: Excel.ChartAreaFormat;
        /**
         *
         * Represents the legend for the chart. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        legend: Excel.ChartLegend;
        /**
         *
         * Represents either a single series or collection of series in the chart. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        series: Excel.ChartSeriesCollection;
        /**
         *
         * Represents the title of the specified chart, including the text, visibility, position and formating of the title. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        title: Excel.ChartTitle;
        /**
         *
         * The worksheet containing the current chart. Read-only.
         *
         * [Api set: ExcelApi 1.2]
         */
        worksheet: Excel.Worksheet;
        /**
         *
         * Represents the height, in points, of the chart object.
         *
         * [Api set: ExcelApi 1.1]
         */
        height: number;
        /**
         *
         * The distance, in points, from the left side of the chart to the worksheet origin.
         *
         * [Api set: ExcelApi 1.1]
         */
        left: number;
        /**
         *
         * Represents the name of a chart object.
         *
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /**
         *
         * Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
         *
         * [Api set: ExcelApi 1.1]
         */
        top: number;
        /**
         *
         * Represents the width, in points, of the chart object.
         *
         * [Api set: ExcelApi 1.1]
         */
        width: number;
        /**
         *
         * Deletes the chart object.
         *
         * [Api set: ExcelApi 1.1]
         */
        delete(): void;
        /**
         *
         * Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.
            The aspect ratio is preserved as part of the resizing.
         *
         * @param height (Optional) The desired height of the resulting image.
         * @param width (Optional) The desired width of the resulting image.
         * @param fittingMode (Optional) The method used to scale the chart to the specified to the specified dimensions (if both height and width are set)."
         *
         * [Api set: ExcelApi 1.2]
         */
        getImage(width?: number, height?: number, fittingMode?: string): OfficeExtension.ClientResult<string>;
        /**
         *
         * Resets the source data for the chart.
         *
         * @param sourceData The Range object corresponding to the source data.
         * @param seriesBy Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, Columns. See Excel.ChartSeriesBy for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        setData(sourceData: Excel.Range, seriesBy?: string): void;
        /**
         *
         * Positions the chart relative to cells on the worksheet.
         *
         * @param startCell The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's right-to-left display settings.
         * @param endCell (Optional) The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range.
         *
         * [Api set: ExcelApi 1.1]
         */
        setPosition(startCell: Excel.Range | string, endCell?: Excel.Range | string): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.Chart;
    }
    /**
     *
     * Encapsulates the format properties for the overall chart area.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartAreaFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_font;
        /**
         *
         * Represents the fill format of an object, which includes background formatting information. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        fill: Excel.ChartFill;
        /**
         *
         * Represents the font attributes (font name, font size, color, etc.) for the current object. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        font: Excel.ChartFont;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartAreaFormat;
    }
    /**
     *
     * Represents a collection of chart series.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartSeriesCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.ChartSeries>;
        /**
         *
         * Returns the number of series in the collection. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        count: number;
        /**
         *
         * Retrieves a series based on its position in the collection
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItemAt(index: number): Excel.ChartSeries;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartSeriesCollection;
    }
    /**
     *
     * Represents a series in a chart.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartSeries extends OfficeExtension.ClientObject {
        private m_format;
        private m_name;
        private m_points;
        /**
         *
         * Represents the formatting of a chart series, which includes fill and line formatting. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        format: Excel.ChartSeriesFormat;
        /**
         *
         * Represents a collection of all points in the series. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        points: Excel.ChartPointsCollection;
        /**
         *
         * Represents the name of a series in a chart.
         *
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartSeries;
    }
    /**
     *
     * encapsulates the format properties for the chart series
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartSeriesFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_line;
        /**
         *
         * Represents the fill format of a chart series, which includes background formating information. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        fill: Excel.ChartFill;
        /**
         *
         * Represents line formatting. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        line: Excel.ChartLineFormat;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartSeriesFormat;
    }
    /**
     *
     * A collection of all the chart points within a series inside a chart.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartPointsCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.ChartPoint>;
        /**
         *
         * Returns the number of chart points in the collection. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        count: number;
        /**
         *
         * Retrieve a point based on its position within the series.
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         *
         * [Api set: ExcelApi 1.1]
         */
        getItemAt(index: number): Excel.ChartPoint;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartPointsCollection;
    }
    /**
     *
     * Represents a point of a series in a chart.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartPoint extends OfficeExtension.ClientObject {
        private m_format;
        private m_value;
        /**
         *
         * Encapsulates the format properties chart point. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        format: Excel.ChartPointFormat;
        /**
         *
         * Returns the value of a chart point. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        value: any;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartPoint;
    }
    /**
     *
     * Represents formatting object for chart points.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartPointFormat extends OfficeExtension.ClientObject {
        private m_fill;
        /**
         *
         * Represents the fill format of a chart, which includes background formating information. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        fill: Excel.ChartFill;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartPointFormat;
    }
    /**
     *
     * Represents the chart axes.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartAxes extends OfficeExtension.ClientObject {
        private m_categoryAxis;
        private m_seriesAxis;
        private m_valueAxis;
        /**
         *
         * Represents the category axis in a chart. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        categoryAxis: Excel.ChartAxis;
        /**
         *
         * Represents the series axis of a 3-dimensional chart. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        seriesAxis: Excel.ChartAxis;
        /**
         *
         * Represents the value axis in an axis. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        valueAxis: Excel.ChartAxis;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartAxes;
    }
    /**
     *
     * Represents a single axis in a chart.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartAxis extends OfficeExtension.ClientObject {
        private m_format;
        private m_majorGridlines;
        private m_majorUnit;
        private m_maximum;
        private m_minimum;
        private m_minorGridlines;
        private m_minorUnit;
        private m_title;
        /**
         *
         * Represents the formatting of a chart object, which includes line and font formatting. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        format: Excel.ChartAxisFormat;
        /**
         *
         * Returns a gridlines object that represents the major gridlines for the specified axis. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        majorGridlines: Excel.ChartGridlines;
        /**
         *
         * Returns a Gridlines object that represents the minor gridlines for the specified axis. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        minorGridlines: Excel.ChartGridlines;
        /**
         *
         * Represents the axis title. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        title: Excel.ChartAxisTitle;
        /**
         *
         * Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number.
         *
         * [Api set: ExcelApi 1.1]
         */
        majorUnit: any;
        /**
         *
         * Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.
         *
         * [Api set: ExcelApi 1.1]
         */
        maximum: any;
        /**
         *
         * Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.
         *
         * [Api set: ExcelApi 1.1]
         */
        minimum: any;
        /**
         *
         * Represents the interval between two minor tick marks. "Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
         *
         * [Api set: ExcelApi 1.1]
         */
        minorUnit: any;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartAxis;
    }
    /**
     *
     * Encapsulates the format properties for the chart axis.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartAxisFormat extends OfficeExtension.ClientObject {
        private m_font;
        private m_line;
        /**
         *
         * Represents the font attributes (font name, font size, color, etc.) for a chart axis element. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        font: Excel.ChartFont;
        /**
         *
         * Represents chart line formatting. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        line: Excel.ChartLineFormat;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartAxisFormat;
    }
    /**
     *
     * Represents the title of a chart axis.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartAxisTitle extends OfficeExtension.ClientObject {
        private m_format;
        private m_text;
        private m_visible;
        /**
         *
         * Represents the formatting of chart axis title. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        format: Excel.ChartAxisTitleFormat;
        /**
         *
         * Represents the axis title.
         *
         * [Api set: ExcelApi 1.1]
         */
        text: string;
        /**
         *
         * A boolean that specifies the visibility of an axis title.
         *
         * [Api set: ExcelApi 1.1]
         */
        visible: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartAxisTitle;
    }
    /**
     *
     * Represents the chart axis title formatting.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartAxisTitleFormat extends OfficeExtension.ClientObject {
        private m_font;
        /**
         *
         * Represents the font attributes, such as font name, font size, color, etc. of chart axis title object. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        font: Excel.ChartFont;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartAxisTitleFormat;
    }
    /**
     *
     * Represents a collection of all the data labels on a chart point.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartDataLabels extends OfficeExtension.ClientObject {
        private m_format;
        private m_position;
        private m_separator;
        private m_showBubbleSize;
        private m_showCategoryName;
        private m_showLegendKey;
        private m_showPercentage;
        private m_showSeriesName;
        private m_showValue;
        /**
         *
         * Represents the format of chart data labels, which includes fill and font formatting. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        format: Excel.ChartDataLabelFormat;
        /**
         *
         * DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        position: string;
        /**
         *
         * String representing the separator used for the data labels on a chart.
         *
         * [Api set: ExcelApi 1.1]
         */
        separator: string;
        /**
         *
         * Boolean value representing if the data label bubble size is visible or not.
         *
         * [Api set: ExcelApi 1.1]
         */
        showBubbleSize: boolean;
        /**
         *
         * Boolean value representing if the data label category name is visible or not.
         *
         * [Api set: ExcelApi 1.1]
         */
        showCategoryName: boolean;
        /**
         *
         * Boolean value representing if the data label legend key is visible or not.
         *
         * [Api set: ExcelApi 1.1]
         */
        showLegendKey: boolean;
        /**
         *
         * Boolean value representing if the data label percentage is visible or not.
         *
         * [Api set: ExcelApi 1.1]
         */
        showPercentage: boolean;
        /**
         *
         * Boolean value representing if the data label series name is visible or not.
         *
         * [Api set: ExcelApi 1.1]
         */
        showSeriesName: boolean;
        /**
         *
         * Boolean value representing if the data label value is visible or not.
         *
         * [Api set: ExcelApi 1.1]
         */
        showValue: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartDataLabels;
    }
    /**
     *
     * Encapsulates the format properties for the chart data labels.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartDataLabelFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_font;
        /**
         *
         * Represents the fill format of the current chart data label. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        fill: Excel.ChartFill;
        /**
         *
         * Represents the font attributes (font name, font size, color, etc.) for a chart data label. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        font: Excel.ChartFont;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartDataLabelFormat;
    }
    /**
     *
     * Represents major or minor gridlines on a chart axis.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartGridlines extends OfficeExtension.ClientObject {
        private m_format;
        private m_visible;
        /**
         *
         * Represents the formatting of chart gridlines. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        format: Excel.ChartGridlinesFormat;
        /**
         *
         * Boolean value representing if the axis gridlines are visible or not.
         *
         * [Api set: ExcelApi 1.1]
         */
        visible: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartGridlines;
    }
    /**
     *
     * Encapsulates the format properties for chart gridlines.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartGridlinesFormat extends OfficeExtension.ClientObject {
        private m_line;
        /**
         *
         * Represents chart line formatting. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        line: Excel.ChartLineFormat;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartGridlinesFormat;
    }
    /**
     *
     * Represents the legend in a chart.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartLegend extends OfficeExtension.ClientObject {
        private m_format;
        private m_overlay;
        private m_position;
        private m_visible;
        /**
         *
         * Represents the formatting of a chart legend, which includes fill and font formatting. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        format: Excel.ChartLegendFormat;
        /**
         *
         * Boolean value for whether the chart legend should overlap with the main body of the chart.
         *
         * [Api set: ExcelApi 1.1]
         */
        overlay: boolean;
        /**
         *
         * Represents the position of the legend on the chart. See Excel.ChartLegendPosition for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        position: string;
        /**
         *
         * A boolean value the represents the visibility of a ChartLegend object.
         *
         * [Api set: ExcelApi 1.1]
         */
        visible: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartLegend;
    }
    /**
     *
     * Encapsulates the format properties of a chart legend.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartLegendFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_font;
        /**
         *
         * Represents the fill format of an object, which includes background formating information. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        fill: Excel.ChartFill;
        /**
         *
         * Represents the font attributes such as font name, font size, color, etc. of a chart legend. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        font: Excel.ChartFont;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartLegendFormat;
    }
    /**
     *
     * Represents a chart title object of a chart.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartTitle extends OfficeExtension.ClientObject {
        private m_format;
        private m_overlay;
        private m_text;
        private m_visible;
        /**
         *
         * Represents the formatting of a chart title, which includes fill and font formatting. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        format: Excel.ChartTitleFormat;
        /**
         *
         * Boolean value representing if the chart title will overlay the chart or not.
         *
         * [Api set: ExcelApi 1.1]
         */
        overlay: boolean;
        /**
         *
         * Represents the title text of a chart.
         *
         * [Api set: ExcelApi 1.1]
         */
        text: string;
        /**
         *
         * A boolean value the represents the visibility of a chart title object.
         *
         * [Api set: ExcelApi 1.1]
         */
        visible: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartTitle;
    }
    /**
     *
     * Provides access to the office art formatting for chart title.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartTitleFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_font;
        /**
         *
         * Represents the fill format of an object, which includes background formating information. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        fill: Excel.ChartFill;
        /**
         *
         * Represents the font attributes (font name, font size, color, etc.) for an object. Read-only.
         *
         * [Api set: ExcelApi 1.1]
         */
        font: Excel.ChartFont;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartTitleFormat;
    }
    /**
     *
     * Represents the fill formatting for a chart element.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartFill extends OfficeExtension.ClientObject {
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartFill;
        /**
         *
         * Clear the fill color of a chart element.
         *
         * [Api set: ExcelApi 1.1]
         */
        clear(): void;
        /**
         *
         * Sets the fill formatting of a chart element to a uniform color.
         *
         * @param color HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         *
         * [Api set: ExcelApi 1.1]
         */
        setSolidColor(color: string): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
    }
    /**
     *
     * Enapsulates the formatting options for line elements.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartLineFormat extends OfficeExtension.ClientObject {
        private m_color;
        /**
         *
         * HTML color code representing the color of lines in the chart.
         *
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        /**
         *
         * Clear the line format of a chart element.
         *
         * [Api set: ExcelApi 1.1]
         */
        clear(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartLineFormat;
    }
    /**
     *
     * This object represents the font attributes (font name, font size, color, etc.) for a chart object.
     *
     * [Api set: ExcelApi 1.1]
     */
    class ChartFont extends OfficeExtension.ClientObject {
        private m_bold;
        private m_color;
        private m_italic;
        private m_name;
        private m_size;
        private m_underline;
        /**
         *
         * Represents the bold status of font.
         *
         * [Api set: ExcelApi 1.1]
         */
        bold: boolean;
        /**
         *
         * HTML color code representation of the text color. E.g. #FF0000 represents Red.
         *
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        /**
         *
         * Represents the italic status of the font.
         *
         * [Api set: ExcelApi 1.1]
         */
        italic: boolean;
        /**
         *
         * Font name (e.g. "Calibri")
         *
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /**
         *
         * Size of the font (e.g. 11)
         *
         * [Api set: ExcelApi 1.1]
         */
        size: number;
        /**
         *
         * Type of underline applied to the font. See Excel.ChartUnderlineStyle for details.
         *
         * [Api set: ExcelApi 1.1]
         */
        underline: string;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ChartFont;
    }
    /**
     *
     * Manages sorting operations on Range objects.
     *
     * [Api set: ExcelApi 1.2]
     */
    class RangeSort extends OfficeExtension.ClientObject {
        /**
         *
         * Perform a sort operation.
         *
         * @param fields The list of conditions to sort on.
         * @param matchCase Whether to have the casing impact string ordering.
         * @param hasHeaders Whether the range has a header.
         * @param orientation Whether the operation is sorting rows or columns.
         * @param method The ordering method used for Chinese characters.
         *
         * [Api set: ExcelApi 1.2]
         */
        apply(fields: Array<Excel.SortField>, matchCase?: boolean, hasHeaders?: boolean, orientation?: string, method?: string): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
    }
    /**
     *
     * Manages sorting operations on Table objects.
     *
     * [Api set: ExcelApi 1.2]
     */
    class TableSort extends OfficeExtension.ClientObject {
        private m_fields;
        private m_matchCase;
        private m_method;
        /**
         *
         * Represents the current conditions used to last sort the table.
         *
         * [Api set: ExcelApi 1.2]
         */
        fields: Array<Excel.SortField>;
        /**
         *
         * Represents whether the casing impacted the last sort of the table.
         *
         * [Api set: ExcelApi 1.2]
         */
        matchCase: boolean;
        /**
         *
         * Represents Chinese character ordering method last used to sort the table.
         *
         * [Api set: ExcelApi 1.2]
         */
        method: string;
        /**
         *
         * Perform a sort operation.
         *
         * @param fields The list of conditions to sort on.
         * @param matchCase Whether to have the casing impact string ordering.
         * @param method The ordering method used for Chinese characters.
         *
         * [Api set: ExcelApi 1.2]
         */
        apply(fields: Array<Excel.SortField>, matchCase?: boolean, method?: string): void;
        /**
         *
         * Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.
         *
         * [Api set: ExcelApi 1.2]
         */
        clear(): void;
        /**
         *
         * Reapplies the current sorting parameters to the table.
         *
         * [Api set: ExcelApi 1.2]
         */
        reapply(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.TableSort;
    }
    /**
     *
     * Represents a condition in a sorting operation.
     *
     * [Api set: ExcelApi 1.2]
     */
    interface SortField {
        /**
         *
         * Represents whether the sorting is done in an ascending fashion.
         *
         * [Api set: ExcelApi 1.2]
         */
        ascending?: boolean;
        /**
         *
         * Represents the color that is the target of the condition if the sorting is on font or cell color.
         *
         * [Api set: ExcelApi 1.2]
         */
        color?: string;
        /**
         *
         * Represents additional sorting options for this field.
         *
         * [Api set: ExcelApi 1.2]
         */
        dataOption?: string;
        /**
         *
         * Represents the icon that is the target of the condition if the sorting is on the cell's icon.
         *
         * [Api set: ExcelApi 1.2]
         */
        icon?: Excel.Icon;
        /**
         *
         * Represents the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).
         *
         * [Api set: ExcelApi 1.2]
         */
        key: number;
        /**
         *
         * Represents the type of sorting of this condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        sortOn?: string;
    }
    /**
     *
     * Manages the filtering of a table's column.
     *
     * [Api set: ExcelApi 1.2]
     */
    class Filter extends OfficeExtension.ClientObject {
        private m_criteria;
        /**
         *
         * The currently applied filter on the given column.
         *
         * [Api set: ExcelApi 1.2]
         */
        criteria: Excel.FilterCriteria;
        /**
         *
         * Apply the given filter criteria on the given column.
         *
         * @param criteria The criteria to apply.
         *
         * [Api set: ExcelApi 1.2]
         */
        apply(criteria: Excel.FilterCriteria): void;
        /**
         *
         * Apply a "Bottom Item" filter to the column for the given number of elements.
         *
         * @param count The number of elements from the bottom to show.
         *
         * [Api set: ExcelApi 1.2]
         */
        applyBottomItemsFilter(count: number): void;
        /**
         *
         * Apply a "Bottom Percent" filter to the column for the given percentage of elements.
         *
         * @param percent The percentage of elements from the bottom to show.
         *
         * [Api set: ExcelApi 1.2]
         */
        applyBottomPercentFilter(percent: number): void;
        /**
         *
         * Apply a "Cell Color" filter to the column for the given color.
         *
         * @param color The background color of the cells to show.
         *
         * [Api set: ExcelApi 1.2]
         */
        applyCellColorFilter(color: string): void;
        /**
         *
         * Apply a "Icon" filter to the column for the given criteria strings.
         *
         * @param criteria1 The first criteria string.
         * @param criteria2 The second criteria string.
         * @param oper The operator that describes how the two criteria are joined.
         *
         * [Api set: ExcelApi 1.2]
         */
        applyCustomFilter(criteria1: string, criteria2?: string, oper?: string): void;
        /**
         *
         * Apply a "Dynamic" filter to the column.
         *
         * @param criteria The dynamic criteria to apply.
         *
         * [Api set: ExcelApi 1.2]
         */
        applyDynamicFilter(criteria: string): void;
        /**
         *
         * Apply a "Font Color" filter to the column for the given color.
         *
         * @param color The font color of the cells to show.
         *
         * [Api set: ExcelApi 1.2]
         */
        applyFontColorFilter(color: string): void;
        /**
         *
         * Apply a "Icon" filter to the column for the given icon.
         *
         * @param icon The icons of the cells to show.
         *
         * [Api set: ExcelApi 1.2]
         */
        applyIconFilter(icon: Excel.Icon): void;
        /**
         *
         * Apply a "Top Item" filter to the column for the given number of elements.
         *
         * @param count The number of elements from the top to show.
         *
         * [Api set: ExcelApi 1.2]
         */
        applyTopItemsFilter(count: number): void;
        /**
         *
         * Apply a "Top Percent" filter to the column for the given percentage of elements.
         *
         * @param percent The percentage of elements from the top to show.
         *
         * [Api set: ExcelApi 1.2]
         */
        applyTopPercentFilter(percent: number): void;
        /**
         *
         * Apply a "Values" filter to the column for the given values.
         *
         * @param values The list of values to show.
         *
         * [Api set: ExcelApi 1.2]
         */
        applyValuesFilter(values: Array<string | Excel.FilterDatetime>): void;
        /**
         *
         * Clear the filter on the given column.
         *
         * [Api set: ExcelApi 1.2]
         */
        clear(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.Filter;
    }
    /**
     *
     * Represents the filtering criteria applied to a column.
     *
     * [Api set: ExcelApi 1.2]
     */
    interface FilterCriteria {
        /**
         *
         * The HTML color string used to filter cells. Used with "cellColor" and "fontColor" filtering.
         *
         * [Api set: ExcelApi 1.2]
         */
        color?: string;
        /**
         *
         * The first criterion used to filter data. Used as an operator in the case of "custom" filtering.
             For example ">50" for number greater than 50 or "=*s" for values ending in "s".
            
             Used as a number in the case of top/bottom items/percents. E.g. "5" for the top 5 items if filterOn is set to "topItems"
         *
         * [Api set: ExcelApi 1.2]
         */
        criterion1?: string;
        /**
         *
         * The second criterion used to filter data. Only used as an operator in the case of "custom" filtering.
         *
         * [Api set: ExcelApi 1.2]
         */
        criterion2?: string;
        /**
         *
         * The dynamic criteria from the Excel.DynamicFilterCriteria set to apply on this column. Used with "dynamic" filtering.
         *
         * [Api set: ExcelApi 1.2]
         */
        dynamicCriteria?: string;
        /**
         *
         * The property used by the filter to determine whether the values should stay visible.
         *
         * [Api set: ExcelApi 1.2]
         */
        filterOn: string;
        /**
         *
         * The icon used to filter cells. Used with "icon" filtering.
         *
         * [Api set: ExcelApi 1.2]
         */
        icon?: Excel.Icon;
        /**
         *
         * The operator used to combine criterion 1 and 2 when using "custom" filtering.
         *
         * [Api set: ExcelApi 1.2]
         */
        operator?: string;
        /**
         *
         * The set of values to be used as part of "values" filtering.
         *
         * [Api set: ExcelApi 1.2]
         */
        values?: Array<string | Excel.FilterDatetime>;
    }
    /**
     *
     * Represents how to filter a date when filtering on values.
     *
     * [Api set: ExcelApi 1.2]
     */
    interface FilterDatetime {
        /**
         *
         * The date in ISO8601 format used to filter data.
         *
         * [Api set: ExcelApi 1.2]
         */
        date: string;
        /**
         *
         * How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specifity is set to "month", the filter operation will keep all rows with a date in the month of april 2009.
         *
         * [Api set: ExcelApi 1.2]
         */
        specificity: string;
    }
    /**
     *
     * Represents a cell icon.
     *
     * [Api set: ExcelApi 1.2]
     */
    interface Icon {
        /**
         *
         * Represents the index of the icon in the given set.
         *
         * [Api set: ExcelApi 1.2]
         */
        index: number;
        /**
         *
         * Represents the set that the icon is part of.
         *
         * [Api set: ExcelApi 1.2]
         */
        set: string;
    }
    /**
     * [Api set: ExcelApi 1.3]
     */
    class _V1Api extends OfficeExtension.ClientObject {
        bindingAddColumns(input: any): OfficeExtension.ClientResult<any>;
        bindingAddFromNamedItem(input: any): OfficeExtension.ClientResult<any>;
        bindingAddFromPrompt(input: any): OfficeExtension.ClientResult<any>;
        bindingAddFromSelection(input: any): OfficeExtension.ClientResult<any>;
        bindingAddRows(input: any): OfficeExtension.ClientResult<any>;
        bindingClearFormats(input: any): OfficeExtension.ClientResult<any>;
        bindingDeleteAllDataValues(input: any): OfficeExtension.ClientResult<any>;
        bindingGetAll(): OfficeExtension.ClientResult<any>;
        bindingGetById(input: any): OfficeExtension.ClientResult<any>;
        bindingGetData(input: any): OfficeExtension.ClientResult<any>;
        bindingReleaseById(input: any): OfficeExtension.ClientResult<any>;
        bindingSetData(input: any): OfficeExtension.ClientResult<any>;
        bindingSetFormats(input: any): OfficeExtension.ClientResult<any>;
        bindingSetTableOptions(input: any): OfficeExtension.ClientResult<any>;
        getSelectedData(input: any): OfficeExtension.ClientResult<any>;
        gotoById(input: any): OfficeExtension.ClientResult<any>;
        setSelectedData(input: any): OfficeExtension.ClientResult<any>;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
    }
    /**
     *
     * Represents a collection of all the PivotTables that are part of the workbook or worksheet.
     *
     * [Api set: ExcelApi 1.3]
     */
    class PivotTableCollection extends OfficeExtension.ClientObject {
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.PivotTable>;
        /**
         *
         * Gets a PivotTable by name.
         *
         * @param name Name of the PivotTable to be retrieved.
         *
         * [Api set: ExcelApi 1.3]
         */
        getItem(name: string): Excel.PivotTable;
        /**
         *
         * Gets a PivotTable by name. If the PivotTable does not exist, the return object's isNull property will be true.
         *
         * @param name Name of the PivotTable to be retrieved.
         *
         * [Api set: ExcelApi 1.3]
         */
        getItemOrNull(name: string): Excel.PivotTable;
        /**
         *
         * Refreshes all the PivotTables in the collection.
         *
         * [Api set: ExcelApi 1.3]
         */
        refreshAll(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.PivotTableCollection;
    }
    /**
     *
     * Represents an Excel PivotTable.
     *
     * [Api set: ExcelApi 1.3]
     */
    class PivotTable extends OfficeExtension.ClientObject {
        private m_name;
        private m_worksheet;
        /**
         *
         * The worksheet containing the current PivotTable. Read-only.
         *
         * [Api set: ExcelApi 1.3]
         */
        worksheet: Excel.Worksheet;
        /**
         *
         * Name of the PivotTable.
         *
         * [Api set: ExcelApi 1.3]
         */
        name: string;
        /**
         *
         * Refreshes the PivotTable.
         *
         * [Api set: ExcelApi 1.3]
         */
        refresh(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.PivotTable;
    }
    /**
     *
     * Represents a collection of all the conditional formats that are overlap the range.
     *
     * [Api set: ExcelApi 1.4]
     */
    class ConditionalFormatCollection extends OfficeExtension.ClientObject {
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.ConditionalFormat>;
        /**
         *
         * Clears all conditional formats active on the current specified range.
         *
         * [Api set: ExcelApi 1.4]
         */
        clearAll(): void;
        /**
         *
         * Returns the number of conditional formats in the workbook. Read-only.
         *
         * [Api set: ExcelApi 1.4]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         *
         * Returns a conditional format at the given index.
         *
         * @param index Index of the conditional formats to be retrieved.
         *
         * [Api set: ExcelApi 1.4]
         */
        getItemAt(index: number): Excel.ConditionalFormat;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ConditionalFormatCollection;
    }
    /**
     *
     * An object encapsulating a conditional format's range, format, rule, and other properties.
     *
     * [Api set: ExcelApi 1.4]
     */
    class ConditionalFormat extends OfficeExtension.ClientObject {
        private m_dataBarOrNull;
        private m_priority;
        private m_stopIfTrue;
        private m_type;
        /**
         *
         * Represents databars with customizable color, gradient, axis, and range format options.
            If no properties are set, a databar is created with the automatic default settings.
         *
         * [Api set: ExcelApi 1.4]
         */
        dataBarOrNull: Excel.ConditionalFormatDataBar;
        /**
         *
         * The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also
            changes other conditional formats' priorities, to allow for a contiguous priority order.
            Use a negative priority to begin from the back.
            Priorities greater than than bounds will get and set to the maximum (or minimum if negative) priority.
         *
         * [Api set: ExcelApi 1.4]
         */
        priority: number;
        /**
         *
         * If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.
            Null on databars, icon sets, and colorscales as there's no concept of StopIfTrue for these
         *
         * [Api set: ExcelApi 1.4]
         */
        stopIfTrue: boolean;
        /**
         *
         * A type of conditional format. Only one can be set at a time. Read-Only.
         *
         * [Api set: ExcelApi 1.4]
         */
        type: string;
        /**
         *
         * Deletes this conditional format.
         *
         * [Api set: ExcelApi 1.4]
         */
        delete(): void;
        /**
         *
         * Returns the range the conditonal format is applied to or null if the range is discontiguous. Read-only.
         *
         * [Api set: ExcelApi 1.4]
         */
        getRangeOrNull(): Excel.Range;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ConditionalFormat;
    }
    /**
     *
     * Represents an Excel Conditional Data Bar Type.
     *
     * [Api set: ExcelApi 1.4]
     */
    class ConditionalFormatDataBar extends OfficeExtension.ClientObject {
        private m_axisColor;
        private m_axisFormat;
        private m_barDirection;
        private m_lowerBoundRule;
        private m_negativeFormat;
        private m_positiveFormat;
        private m_showDataBarOnly;
        private m_upperBoundRule;
        /**
         *
         * The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.
         *
         * [Api set: ExcelApi 1.4]
         */
        lowerBoundRule: Excel.ConditionalFormatDataBarRule;
        /**
         *
         * Representation of all values to the left of the axis in an Excel data bar.
         *
         * [Api set: ExcelApi 1.4]
         */
        negativeFormat: Excel.ConditionalFormatDataBarNegativeFormat;
        /**
         *
         * Representation of all values to the right of the axis in an Excel data bar.
         *
         * [Api set: ExcelApi 1.4]
         */
        positiveFormat: Excel.ConditionalFormatDataBarPositiveFormat;
        /**
         *
         * The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.
         *
         * [Api set: ExcelApi 1.4]
         */
        upperBoundRule: Excel.ConditionalFormatDataBarRule;
        /**
         *
         * HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
            "" (empty string) if no axis is present or set.
         *
         * [Api set: ExcelApi 1.4]
         */
        axisColor: string;
        /**
         *
         * Representation of how the axis is determined for an Excel data bar.
         *
         * [Api set: ExcelApi 1.4]
         */
        axisFormat: string;
        /**
         *
         * Represents the direction that the data bar graphic should be based on.
         *
         * [Api set: ExcelApi 1.4]
         */
        barDirection: string;
        /**
         *
         * If true, hides the values from the cells where the data bar is applied.
         *
         * [Api set: ExcelApi 1.4]
         */
        showDataBarOnly: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ConditionalFormatDataBar;
    }
    /**
     *
     * Represents a conditional format DataBar Format for the positive side of the data bar.
     *
     * [Api set: ExcelApi 1.4]
     */
    class ConditionalFormatDataBarPositiveFormat extends OfficeExtension.ClientObject {
        private m_borderColor;
        private m_fillColor;
        private m_gradientFill;
        /**
         *
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
            "" (empty string) if no border is present or set.
         *
         * [Api set: ExcelApi 1.4]
         */
        borderColor: string;
        /**
         *
         * HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         *
         * [Api set: ExcelApi 1.4]
         */
        fillColor: string;
        /**
         *
         * Boolean representation of whether or not the DataBar has a gradient.
         *
         * [Api set: ExcelApi 1.4]
         */
        gradientFill: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ConditionalFormatDataBarPositiveFormat;
    }
    /**
     *
     * Represents a conditional format DataBar Format for the negative side of the data bar.
     *
     * [Api set: ExcelApi 1.4]
     */
    class ConditionalFormatDataBarNegativeFormat extends OfficeExtension.ClientObject {
        private m_borderColor;
        private m_fillColor;
        private m_matchPositiveBorderColor;
        private m_matchPositiveFillColor;
        /**
         *
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
            "Empty String" if no border is present or set.
         *
         * [Api set: ExcelApi 1.4]
         */
        borderColor: string;
        /**
         *
         * HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         *
         * [Api set: ExcelApi 1.4]
         */
        fillColor: string;
        /**
         *
         * Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.
         *
         * [Api set: ExcelApi 1.4]
         */
        matchPositiveBorderColor: boolean;
        /**
         *
         * Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.
         *
         * [Api set: ExcelApi 1.4]
         */
        matchPositiveFillColor: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ConditionalFormatDataBarNegativeFormat;
    }
    /**
     *
     * Represents a rule-type for a Data Bar.
     *
     * [Api set: ExcelApi 1.4]
     */
    class ConditionalFormatDataBarRule extends OfficeExtension.ClientObject {
        private m_formula;
        private m_formulaLocal;
        private m_formulaR1C1;
        private m_type;
        /**
         *
         * The formula, if required, to evaluate the databar rule on.
         *
         * [Api set: ExcelApi 1.4]
         */
        formula: string | number;
        /**
         *
         * The formula, if required, to evaluate the databar rule on in the user's language.
         *
         * [Api set: ExcelApi 1.4]
         */
        formulaLocal: string | number;
        /**
         *
         * The formula, if required, to evaluate the databar rule on in R1C1-style notation.
         *
         * [Api set: ExcelApi 1.4]
         */
        formulaR1C1: string | number;
        /**
         *
         * What the databar should be based on.
         *
         * [Api set: ExcelApi 1.4]
         */
        type: string;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): Excel.ConditionalFormatDataBarRule;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module BindingType {
        var range: string;
        var table: string;
        var text: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module BorderIndex {
        var edgeTop: string;
        var edgeBottom: string;
        var edgeLeft: string;
        var edgeRight: string;
        var insideVertical: string;
        var insideHorizontal: string;
        var diagonalDown: string;
        var diagonalUp: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module BorderLineStyle {
        var none: string;
        var continuous: string;
        var dash: string;
        var dashDot: string;
        var dashDotDot: string;
        var dot: string;
        var double: string;
        var slantDashDot: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module BorderWeight {
        var hairline: string;
        var thin: string;
        var medium: string;
        var thick: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module CalculationMode {
        var automatic: string;
        var automaticExceptTables: string;
        var manual: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module CalculationType {
        var recalculate: string;
        var full: string;
        var fullRebuild: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module ClearApplyTo {
        var all: string;
        var formats: string;
        var contents: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module ChartDataLabelPosition {
        var invalid: string;
        var none: string;
        var center: string;
        var insideEnd: string;
        var insideBase: string;
        var outsideEnd: string;
        var left: string;
        var right: string;
        var top: string;
        var bottom: string;
        var bestFit: string;
        var callout: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module ChartLegendPosition {
        var invalid: string;
        var top: string;
        var bottom: string;
        var left: string;
        var right: string;
        var corner: string;
        var custom: string;
    }
    /**
     *
     * Specifies whether the series are by rows or by columns. On Desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns; on Excel Online, "auto" will simply default to "columns".
     *
     * [Api set: ExcelApi 1.1]
     */
    module ChartSeriesBy {
        /**
         *
         * On Desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns; on Excel Online, "auto" will simply default to "columns".
         *
         */
        var auto: string;
        var columns: string;
        var rows: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module ChartType {
        var invalid: string;
        var columnClustered: string;
        var columnStacked: string;
        var columnStacked100: string;
        var _3DColumnClustered: string;
        var _3DColumnStacked: string;
        var _3DColumnStacked100: string;
        var barClustered: string;
        var barStacked: string;
        var barStacked100: string;
        var _3DBarClustered: string;
        var _3DBarStacked: string;
        var _3DBarStacked100: string;
        var lineStacked: string;
        var lineStacked100: string;
        var lineMarkers: string;
        var lineMarkersStacked: string;
        var lineMarkersStacked100: string;
        var pieOfPie: string;
        var pieExploded: string;
        var _3DPieExploded: string;
        var barOfPie: string;
        var xyscatterSmooth: string;
        var xyscatterSmoothNoMarkers: string;
        var xyscatterLines: string;
        var xyscatterLinesNoMarkers: string;
        var areaStacked: string;
        var areaStacked100: string;
        var _3DAreaStacked: string;
        var _3DAreaStacked100: string;
        var doughnutExploded: string;
        var radarMarkers: string;
        var radarFilled: string;
        var surface: string;
        var surfaceWireframe: string;
        var surfaceTopView: string;
        var surfaceTopViewWireframe: string;
        var bubble: string;
        var bubble3DEffect: string;
        var stockHLC: string;
        var stockOHLC: string;
        var stockVHLC: string;
        var stockVOHLC: string;
        var cylinderColClustered: string;
        var cylinderColStacked: string;
        var cylinderColStacked100: string;
        var cylinderBarClustered: string;
        var cylinderBarStacked: string;
        var cylinderBarStacked100: string;
        var cylinderCol: string;
        var coneColClustered: string;
        var coneColStacked: string;
        var coneColStacked100: string;
        var coneBarClustered: string;
        var coneBarStacked: string;
        var coneBarStacked100: string;
        var coneCol: string;
        var pyramidColClustered: string;
        var pyramidColStacked: string;
        var pyramidColStacked100: string;
        var pyramidBarClustered: string;
        var pyramidBarStacked: string;
        var pyramidBarStacked100: string;
        var pyramidCol: string;
        var _3DColumn: string;
        var line: string;
        var _3DLine: string;
        var _3DPie: string;
        var pie: string;
        var xyscatter: string;
        var _3DArea: string;
        var area: string;
        var doughnut: string;
        var radar: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module ChartUnderlineStyle {
        var none: string;
        var single: string;
    }
    /**
     *
     * Represents the format options for a Data Bar Axis.
     *
     * [Api set: ExcelApi 1.4]
     */
    module ConditionalFormatDataBarAxisFormat {
        var automatic: string;
        var none: string;
        var cellMidPoint: string;
    }
    /**
     *
     * Represents the Data Bar direction within a cell.
     *
     * [Api set: ExcelApi 1.4]
     */
    module ConditionalFormatDataBarDirection {
        var context: string;
        var leftToRight: string;
        var rightToLeft: string;
    }
    /**
     *
     * Represents the direction for a selection.
     *
     * [Api set: ExcelApi 1.4]
     */
    module ConditionalFormatDirection {
        var top: string;
        var bottom: string;
    }
    /**
     * [Api set: ExcelApi 1.4]
     */
    module ConditionalFormatType {
        var custom: string;
        var dataBar: string;
        var colorScale: string;
        var iconSet: string;
    }
    /**
     *
     * Represents the types of conditional format values.
     *
     * [Api set: ExcelApi 1.4]
     */
    module ConditionalFormatRuleType {
        var automatic: string;
        var lowestValue: string;
        var highestValue: string;
        var number: string;
        var percent: string;
        var formula: string;
        var percentile: string;
    }
    /**
     *
     * Represents all of the potential rule types for formats.
     *
     * [Api set: ExcelApi 1.4]
     */
    module ConditionalRangeFormatRuleType {
        var blank: string;
        var expression: string;
        var between: string;
        var notBetween: string;
        var count: string;
        var percent: string;
        var average: string;
        var unique: string;
        var error: string;
        var textContains: string;
        var dateOccurring: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module DeleteShiftDirection {
        var up: string;
        var left: string;
    }
    /**
     * [Api set: ExcelApi 1.2]
     */
    module DynamicFilterCriteria {
        var unknown: string;
        var aboveAverage: string;
        var allDatesInPeriodApril: string;
        var allDatesInPeriodAugust: string;
        var allDatesInPeriodDecember: string;
        var allDatesInPeriodFebruray: string;
        var allDatesInPeriodJanuary: string;
        var allDatesInPeriodJuly: string;
        var allDatesInPeriodJune: string;
        var allDatesInPeriodMarch: string;
        var allDatesInPeriodMay: string;
        var allDatesInPeriodNovember: string;
        var allDatesInPeriodOctober: string;
        var allDatesInPeriodQuarter1: string;
        var allDatesInPeriodQuarter2: string;
        var allDatesInPeriodQuarter3: string;
        var allDatesInPeriodQuarter4: string;
        var allDatesInPeriodSeptember: string;
        var belowAverage: string;
        var lastMonth: string;
        var lastQuarter: string;
        var lastWeek: string;
        var lastYear: string;
        var nextMonth: string;
        var nextQuarter: string;
        var nextWeek: string;
        var nextYear: string;
        var thisMonth: string;
        var thisQuarter: string;
        var thisWeek: string;
        var thisYear: string;
        var today: string;
        var tomorrow: string;
        var yearToDate: string;
        var yesterday: string;
    }
    /**
     * [Api set: ExcelApi 1.2]
     */
    module FilterDatetimeSpecificity {
        var year: string;
        var month: string;
        var day: string;
        var hour: string;
        var minute: string;
        var second: string;
    }
    /**
     * [Api set: ExcelApi 1.2]
     */
    module FilterOn {
        var bottomItems: string;
        var bottomPercent: string;
        var cellColor: string;
        var dynamic: string;
        var fontColor: string;
        var values: string;
        var topItems: string;
        var topPercent: string;
        var icon: string;
        var custom: string;
    }
    /**
     * [Api set: ExcelApi 1.2]
     */
    module FilterOperator {
        var and: string;
        var or: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module HorizontalAlignment {
        var general: string;
        var left: string;
        var center: string;
        var right: string;
        var fill: string;
        var justify: string;
        var centerAcrossSelection: string;
        var distributed: string;
    }
    /**
     * [Api set: ExcelApi 1.2]
     */
    module IconSet {
        var invalid: string;
        var threeArrows: string;
        var threeArrowsGray: string;
        var threeFlags: string;
        var threeTrafficLights1: string;
        var threeTrafficLights2: string;
        var threeSigns: string;
        var threeSymbols: string;
        var threeSymbols2: string;
        var fourArrows: string;
        var fourArrowsGray: string;
        var fourRedToBlack: string;
        var fourRating: string;
        var fourTrafficLights: string;
        var fiveArrows: string;
        var fiveArrowsGray: string;
        var fiveRating: string;
        var fiveQuarters: string;
        var threeStars: string;
        var threeTriangles: string;
        var fiveBoxes: string;
    }
    /**
     * [Api set: ExcelApi 1.2]
     */
    module ImageFittingMode {
        var fit: string;
        var fitAndCenter: string;
        var fill: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module InsertShiftDirection {
        var down: string;
        var right: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module NamedItemType {
        var string: string;
        var integer: string;
        var double: string;
        var boolean: string;
        var range: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module RangeUnderlineStyle {
        var none: string;
        var single: string;
        var double: string;
        var singleAccountant: string;
        var doubleAccountant: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module SheetVisibility {
        var visible: string;
        var hidden: string;
        var veryHidden: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module RangeValueType {
        var unknown: string;
        var empty: string;
        var string: string;
        var integer: string;
        var double: string;
        var boolean: string;
        var error: string;
    }
    /**
     * [Api set: ExcelApi 1.2]
     */
    module SortOrientation {
        var rows: string;
        var columns: string;
    }
    /**
     * [Api set: ExcelApi 1.2]
     */
    module SortOn {
        var value: string;
        var cellColor: string;
        var fontColor: string;
        var icon: string;
    }
    /**
     * [Api set: ExcelApi 1.2]
     */
    module SortDataOption {
        var normal: string;
        var textAsNumber: string;
    }
    /**
     * [Api set: ExcelApi 1.2]
     */
    module SortMethod {
        var pinYin: string;
        var strokeCount: string;
    }
    /**
     * [Api set: ExcelApi 1.1]
     */
    module VerticalAlignment {
        var top: string;
        var center: string;
        var bottom: string;
        var justify: string;
        var distributed: string;
    }
    /**
     *
     * An object containing the result of a function-evaluation operation
     *
     * [Api set: ExcelApi 1.2]
     */
    class FunctionResult<T> extends OfficeExtension.ClientObject {
        private m_error;
        private m_value;
        /**
         *
         * Error value (such as "#DIV/0") representing the error. If the error string is not set, then the function succeeded, and its result is written to the Value field. The error is always in the English locale.
         *
         * [Api set: ExcelApi 1.2]
         */
        error: string;
        /**
         *
         * The value of function evaluation. The value field will be populated only if no error has occured (i.e., the Error property is not set).
         *
         * [Api set: ExcelApi 1.2]
         */
        value: T;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): FunctionResult<T>;
    }
    /**
     *
     * An object for evaluating Excel functions.
     *
     * [Api set: ExcelApi 1.2]
     */
    class Functions extends OfficeExtension.ClientObject {
        /**
         *
         * Returns the absolute value of a number, a number without its sign.
         *
         * @param number Is the real number for which you want the absolute value.
         *
         * [Api set: ExcelApi 1.2]
         */
        abs(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the accrued interest for a security that pays periodic interest.
         *
         * @param issue Is the security's issue date, expressed as a serial date number.
         * @param firstInterest Is the security's first interest date, expressed as a serial date number.
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param rate Is the security's annual coupon rate.
         * @param par Is the security's par value.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         * @param calcMethod Is a logical value: to accrued interest from issue date = TRUE or omitted; to calculate from last coupon payment date = FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        accrInt(issue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, firstInterest: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, par: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, calcMethod?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the accrued interest for a security that pays interest at maturity.
         *
         * @param issue Is the security's issue date, expressed as a serial date number.
         * @param settlement Is the security's maturity date, expressed as a serial date number.
         * @param rate Is the security's annual coupon rate.
         * @param par Is the security's par value.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        accrIntM(issue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, par: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the arccosine of a number, in radians in the range 0 to Pi. The arccosine is the angle whose cosine is Number.
         *
         * @param number Is the cosine of the angle you want and must be from -1 to 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        acos(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse hyperbolic cosine of a number.
         *
         * @param number Is any real number equal to or greater than 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        acosh(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the arccotangent of a number, in radians in the range 0 to Pi.
         *
         * @param number Is the cotangent of the angle you want.
         *
         * [Api set: ExcelApi 1.2]
         */
        acot(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse hyperbolic cotangent of a number.
         *
         * @param number Is the hyperbolic cotangent of the angle that you want.
         *
         * [Api set: ExcelApi 1.2]
         */
        acoth(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the prorated linear depreciation of an asset for each accounting period.
         *
         * @param cost Is the cost of the asset.
         * @param datePurchased Is the date the asset is purchased.
         * @param firstPeriod Is the date of the end of the first period.
         * @param salvage Is the salvage value at the end of life of the asset.
         * @param period Is the period.
         * @param rate Is the rate of depreciation.
         * @param basis Year_basis : 0 for year of 360 days, 1 for actual, 3 for year of 365 days.
         *
         * [Api set: ExcelApi 1.2]
         */
        amorDegrc(cost: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, datePurchased: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, firstPeriod: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvage: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, period: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the prorated linear depreciation of an asset for each accounting period.
         *
         * @param cost Is the cost of the asset.
         * @param datePurchased Is the date the asset is purchased.
         * @param firstPeriod Is the date of the end of the first period.
         * @param salvage Is the salvage value at the end of life of the asset.
         * @param period Is the period.
         * @param rate Is the rate of depreciation.
         * @param basis Year_basis : 0 for year of 360 days, 1 for actual, 3 for year of 365 days.
         *
         * [Api set: ExcelApi 1.2]
         */
        amorLinc(cost: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, datePurchased: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, firstPeriod: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvage: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, period: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Checks whether all arguments are TRUE, and returns TRUE if all arguments are TRUE.
         *
         * @param values List of parameters, whose elements are 1 to 255 conditions you want to test that can be either TRUE or FALSE and can be logical values, arrays, or references.
         *
         * [Api set: ExcelApi 1.2]
         */
        and(...values: Array<boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<boolean>;
        /**
         *
         * Converts a Roman numeral to Arabic.
         *
         * @param text Is the Roman numeral you want to convert.
         *
         * [Api set: ExcelApi 1.2]
         */
        arabic(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of areas in a reference. An area is a range of contiguous cells or a single cell.
         *
         * @param reference Is a reference to a cell or range of cells and can refer to multiple areas.
         *
         * [Api set: ExcelApi 1.2]
         */
        areas(reference: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Changes full-width (double-byte) characters to half-width (single-byte) characters. Use with double-byte character sets (DBCS).
         *
         * @param text Is a text, or a reference to a cell containing a text.
         *
         * [Api set: ExcelApi 1.2]
         */
        asc(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the arcsine of a number in radians, in the range -Pi/2 to Pi/2.
         *
         * @param number Is the sine of the angle you want and must be from -1 to 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        asin(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse hyperbolic sine of a number.
         *
         * @param number Is any real number equal to or greater than 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        asinh(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the arctangent of a number in radians, in the range -Pi/2 to Pi/2.
         *
         * @param number Is the tangent of the angle you want.
         *
         * [Api set: ExcelApi 1.2]
         */
        atan(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the arctangent of the specified x- and y- coordinates, in radians between -Pi and Pi, excluding -Pi.
         *
         * @param xNum Is the x-coordinate of the point.
         * @param yNum Is the y-coordinate of the point.
         *
         * [Api set: ExcelApi 1.2]
         */
        atan2(xNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, yNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse hyperbolic tangent of a number.
         *
         * @param number Is any real number between -1 and 1 excluding -1 and 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        atanh(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the average of the absolute deviations of data points from their mean. Arguments can be numbers or names, arrays, or references that contain numbers.
         *
         * @param values List of parameters, whose elements are 1 to 255 arguments for which you want the average of the absolute deviations.
         *
         * [Api set: ExcelApi 1.2]
         */
        aveDev(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the average (arithmetic mean) of its arguments, which can be numbers or names, arrays, or references that contain numbers.
         *
         * @param values List of parameters, whose elements are 1 to 255 numeric arguments for which you want the average.
         *
         * [Api set: ExcelApi 1.2]
         */
        average(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the average (arithmetic mean) of its arguments, evaluating text and FALSE in arguments as 0; TRUE evaluates as 1. Arguments can be numbers, names, arrays, or references.
         *
         * @param values List of parameters, whose elements are 1 to 255 arguments for which you want the average.
         *
         * [Api set: ExcelApi 1.2]
         */
        averageA(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Finds average(arithmetic mean) for the cells specified by a given condition or criteria.
         *
         * @param range Is the range of cells you want evaluated.
         * @param criteria Is the condition or criteria in the form of a number, expression, or text that defines which cells will be used to find the average.
         * @param averageRange Are the actual cells to be used to find the average. If omitted, the cells in range are used.
         *
         * [Api set: ExcelApi 1.2]
         */
        averageIf(range: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, averageRange?: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Finds average(arithmetic mean) for the cells specified by a given set of conditions or criteria.
         *
         * @param averageRange Are the actual cells to be used to find the average.
         * @param values List of parameters, where the first element of each pair is the Is the range of cells you want evaluated for the particular condition , and the second element is is the condition or criteria in the form of a number, expression, or text that defines which cells will be used to find the average.
         *
         * [Api set: ExcelApi 1.2]
         */
        averageIfs(averageRange: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ...values: Array<Excel.Range | Excel.RangeReference | Excel.FunctionResult<any> | number | string | boolean>): FunctionResult<number>;
        /**
         *
         * Converts a number to text (baht).
         *
         * @param number Is a number that you want to convert.
         *
         * [Api set: ExcelApi 1.2]
         */
        bahtText(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Converts a number into a text representation with the given radix (base).
         *
         * @param number Is the number that you want to convert.
         * @param radix Is the base Radix that you want to convert the number into.
         * @param minLength Is the minimum length of the returned string.  If omitted leading zeros are not added.
         *
         * [Api set: ExcelApi 1.2]
         */
        base(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, radix: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, minLength?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the modified Bessel function In(x).
         *
         * @param x Is the value at which to evaluate the function.
         * @param n Is the order of the Bessel function.
         *
         * [Api set: ExcelApi 1.2]
         */
        besselI(x: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, n: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the Bessel function Jn(x).
         *
         * @param x Is the value at which to evaluate the function.
         * @param n Is the order of the Bessel function.
         *
         * [Api set: ExcelApi 1.2]
         */
        besselJ(x: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, n: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the modified Bessel function Kn(x).
         *
         * @param x Is the value at which to evaluate the function.
         * @param n Is the order of the function.
         *
         * [Api set: ExcelApi 1.2]
         */
        besselK(x: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, n: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the Bessel function Yn(x).
         *
         * @param x Is the value at which to evaluate the function.
         * @param n Is the order of the function.
         *
         * [Api set: ExcelApi 1.2]
         */
        besselY(x: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, n: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the beta probability distribution function.
         *
         * @param x Is the value between A and B at which to evaluate the function.
         * @param alpha Is a parameter to the distribution and must be greater than 0.
         * @param beta Is a parameter to the distribution and must be greater than 0.
         * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE.
         * @param A Is an optional lower bound to the interval of x. If omitted, A = 0.
         * @param B Is an optional upper bound to the interval of x. If omitted, B = 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        beta_Dist(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, beta: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, A?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, B?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse of the cumulative beta probability density function (BETA.DIST).
         *
         * @param probability Is a probability associated with the beta distribution.
         * @param alpha Is a parameter to the distribution and must be greater than 0.
         * @param beta Is a parameter to the distribution and must be greater than 0.
         * @param A Is an optional lower bound to the interval of x. If omitted, A = 0.
         * @param B Is an optional upper bound to the interval of x. If omitted, B = 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        beta_Inv(probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, beta: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, A?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, B?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a binary number to decimal.
         *
         * @param number Is the binary number you want to convert.
         *
         * [Api set: ExcelApi 1.2]
         */
        bin2Dec(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a binary number to hexadecimal.
         *
         * @param number Is the binary number you want to convert.
         * @param places Is the number of characters to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        bin2Hex(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a binary number to octal.
         *
         * @param number Is the binary number you want to convert.
         * @param places Is the number of characters to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        bin2Oct(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the individual term binomial distribution probability.
         *
         * @param numberS Is the number of successes in trials.
         * @param trials Is the number of independent trials.
         * @param probabilityS Is the probability of success on each trial.
         * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability mass function, use FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        binom_Dist(numberS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, trials: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, probabilityS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the probability of a trial result using a binomial distribution.
         *
         * @param trials Is the number of independent trials.
         * @param probabilityS Is the probability of success on each trial.
         * @param numberS Is the number of successes in trials.
         * @param numberS2 If provided this function returns the probability that the number of successful trials shall lie between numberS and numberS2.
         *
         * [Api set: ExcelApi 1.2]
         */
        binom_Dist_Range(trials: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, probabilityS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberS2?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.
         *
         * @param trials Is the number of Bernoulli trials.
         * @param probabilityS Is the probability of success on each trial, a number between 0 and 1 inclusive.
         * @param alpha Is the criterion value, a number between 0 and 1 inclusive.
         *
         * [Api set: ExcelApi 1.2]
         */
        binom_Inv(trials: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, probabilityS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns a bitwise 'And' of two numbers.
         *
         * @param number1 Is the decimal representation of the binary number you want to evaluate.
         * @param number2 Is the decimal representation of the binary number you want to evaluate.
         *
         * [Api set: ExcelApi 1.2]
         */
        bitand(number1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, number2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns a number shifted left by shift_amount bits.
         *
         * @param number Is the decimal representation of the binary number you want to evaluate.
         * @param shiftAmount Is the number of bits that you want to shift Number left by.
         *
         * [Api set: ExcelApi 1.2]
         */
        bitlshift(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, shiftAmount: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns a bitwise 'Or' of two numbers.
         *
         * @param number1 Is the decimal representation of the binary number you want to evaluate.
         * @param number2 Is the decimal representation of the binary number you want to evaluate.
         *
         * [Api set: ExcelApi 1.2]
         */
        bitor(number1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, number2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns a number shifted right by shift_amount bits.
         *
         * @param number Is the decimal representation of the binary number you want to evaluate.
         * @param shiftAmount Is the number of bits that you want to shift Number right by.
         *
         * [Api set: ExcelApi 1.2]
         */
        bitrshift(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, shiftAmount: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns a bitwise 'Exclusive Or' of two numbers.
         *
         * @param number1 Is the decimal representation of the binary number you want to evaluate.
         * @param number2 Is the decimal representation of the binary number you want to evaluate.
         *
         * [Api set: ExcelApi 1.2]
         */
        bitxor(number1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, number2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Rounds a number up, to the nearest integer or to the nearest multiple of significance.
         *
         * @param number Is the value you want to round.
         * @param significance Is the multiple to which you want to round.
         * @param mode When given and nonzero this function will round away from zero.
         *
         * [Api set: ExcelApi 1.2]
         */
        ceiling_Math(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mode?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Rounds a number up, to the nearest integer or to the nearest multiple of significance.
         *
         * @param number Is the value you want to round.
         * @param significance Is the multiple to which you want to round.
         *
         * [Api set: ExcelApi 1.2]
         */
        ceiling_Precise(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the character specified by the code number from the character set for your computer.
         *
         * @param number Is a number between 1 and 255 specifying which character you want.
         *
         * [Api set: ExcelApi 1.2]
         */
        char(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the left-tailed probability of the chi-squared distribution.
         *
         * @param x Is the value at which you want to evaluate the distribution, a nonnegative number.
         * @param degFreedom Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
         * @param cumulative Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        chiSq_Dist(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the right-tailed probability of the chi-squared distribution.
         *
         * @param x Is the value at which you want to evaluate the distribution, a nonnegative number.
         * @param degFreedom Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
         *
         * [Api set: ExcelApi 1.2]
         */
        chiSq_Dist_RT(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse of the left-tailed probability of the chi-squared distribution.
         *
         * @param probability Is a probability associated with the chi-squared distribution, a value between 0 and 1 inclusive.
         * @param degFreedom Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
         *
         * [Api set: ExcelApi 1.2]
         */
        chiSq_Inv(probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse of the right-tailed probability of the chi-squared distribution.
         *
         * @param probability Is a probability associated with the chi-squared distribution, a value between 0 and 1 inclusive.
         * @param degFreedom Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10.
         *
         * [Api set: ExcelApi 1.2]
         */
        chiSq_Inv_RT(probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Chooses a value or action to perform from a list of values, based on an index number.
         *
         * @param indexNum Specifies which value argument is selected. indexNum must be between 1 and 254, or a formula or a reference to a number between 1 and 254.
         * @param values List of parameters, whose elements are 1 to 254 numbers, cell references, defined names, formulas, functions, or text arguments from which CHOOSE selects.
         *
         * [Api set: ExcelApi 1.2]
         */
        choose(indexNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ...values: Array<Excel.Range | number | string | boolean | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number | string | boolean>;
        /**
         *
         * Removes all nonprintable characters from text.
         *
         * @param text Is any worksheet information from which you want to remove nonprintable characters.
         *
         * [Api set: ExcelApi 1.2]
         */
        clean(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns a numeric code for the first character in a text string, in the character set used by your computer.
         *
         * @param text Is the text for which you want the code of the first character.
         *
         * [Api set: ExcelApi 1.2]
         */
        code(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of columns in an array or reference.
         *
         * @param array Is an array or array formula, or a reference to a range of cells for which you want the number of columns.
         *
         * [Api set: ExcelApi 1.2]
         */
        columns(array: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of combinations for a given number of items.
         *
         * @param number Is the total number of items.
         * @param numberChosen Is the number of items in each combination.
         *
         * [Api set: ExcelApi 1.2]
         */
        combin(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberChosen: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of combinations with repetitions for a given number of items.
         *
         * @param number Is the total number of items.
         * @param numberChosen Is the number of items in each combination.
         *
         * [Api set: ExcelApi 1.2]
         */
        combina(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberChosen: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts real and imaginary coefficients into a complex number.
         *
         * @param realNum Is the real coefficient of the complex number.
         * @param iNum Is the imaginary coefficient of the complex number.
         * @param suffix Is the suffix for the imaginary component of the complex number.
         *
         * [Api set: ExcelApi 1.2]
         */
        complex(realNum: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, iNum: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, suffix?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Joins several text strings into one text string.
         *
         * @param values List of parameters, whose elements are 1 to 255 text strings to be joined into a single text string and can be text strings, numbers, or single-cell references.
         *
         * [Api set: ExcelApi 1.2]
         */
        concatenate(...values: Array<string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<string>;
        /**
         *
         * Returns the confidence interval for a population mean, using a normal distribution.
         *
         * @param alpha Is the significance level used to compute the confidence level, a number greater than 0 and less than 1.
         * @param standardDev Is the population standard deviation for the data range and is assumed to be known. standardDev must be greater than 0.
         * @param size Is the sample size.
         *
         * [Api set: ExcelApi 1.2]
         */
        confidence_Norm(alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, size: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the confidence interval for a population mean, using a Student's T distribution.
         *
         * @param alpha Is the significance level used to compute the confidence level, a number greater than 0 and less than 1.
         * @param standardDev Is the population standard deviation for the data range and is assumed to be known. standardDev must be greater than 0.
         * @param size Is the sample size.
         *
         * [Api set: ExcelApi 1.2]
         */
        confidence_T(alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, size: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a number from one measurement system to another.
         *
         * @param number Is the value in from_units to convert.
         * @param fromUnit Is the units for number.
         * @param toUnit Is the units for the result.
         *
         * [Api set: ExcelApi 1.2]
         */
        convert(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fromUnit: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, toUnit: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the cosine of an angle.
         *
         * @param number Is the angle in radians for which you want the cosine.
         *
         * [Api set: ExcelApi 1.2]
         */
        cos(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the hyperbolic cosine of a number.
         *
         * @param number Is any real number.
         *
         * [Api set: ExcelApi 1.2]
         */
        cosh(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the cotangent of an angle.
         *
         * @param number Is the angle in radians for which you want the cotangent.
         *
         * [Api set: ExcelApi 1.2]
         */
        cot(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the hyperbolic cotangent of a number.
         *
         * @param number Is the angle in radians for which you want the hyperbolic cotangent.
         *
         * [Api set: ExcelApi 1.2]
         */
        coth(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Counts the number of cells in a range that contain numbers.
         *
         * @param values List of parameters, whose elements are 1 to 255 arguments that can contain or refer to a variety of different types of data, but only numbers are counted.
         *
         * [Api set: ExcelApi 1.2]
         */
        count(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Counts the number of cells in a range that are not empty.
         *
         * @param values List of parameters, whose elements are 1 to 255 arguments representing the values and cells you want to count. Values can be any type of information.
         *
         * [Api set: ExcelApi 1.2]
         */
        countA(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Counts the number of empty cells in a specified range of cells.
         *
         * @param range Is the range from which you want to count the empty cells.
         *
         * [Api set: ExcelApi 1.2]
         */
        countBlank(range: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Counts the number of cells within a range that meet the given condition.
         *
         * @param range Is the range of cells from which you want to count nonblank cells.
         * @param criteria Is the condition in the form of a number, expression, or text that defines which cells will be counted.
         *
         * [Api set: ExcelApi 1.2]
         */
        countIf(range: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Counts the number of cells specified by a given set of conditions or criteria.
         *
         * @param values List of parameters, where the first element of each pair is the Is the range of cells you want evaluated for the particular condition , and the second element is is the condition in the form of a number, expression, or text that defines which cells will be counted.
         *
         * [Api set: ExcelApi 1.2]
         */
        countIfs(...values: Array<Excel.Range | Excel.RangeReference | Excel.FunctionResult<any> | number | string | boolean>): FunctionResult<number>;
        /**
         *
         * Returns the number of days from the beginning of the coupon period to the settlement date.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        coupDayBs(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of days in the coupon period that contains the settlement date.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        coupDays(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of days from the settlement date to the next coupon date.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        coupDaysNc(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the next coupon date after the settlement date.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        coupNcd(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of coupons payable between the settlement date and maturity date.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        coupNum(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the previous coupon date before the settlement date.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        coupPcd(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the cosecant of an angle.
         *
         * @param number Is the angle in radians for which you want the cosecant.
         *
         * [Api set: ExcelApi 1.2]
         */
        csc(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the hyperbolic cosecant of an angle.
         *
         * @param number Is the angle in radians for which you want the hyperbolic cosecant.
         *
         * [Api set: ExcelApi 1.2]
         */
        csch(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the cumulative interest paid between two periods.
         *
         * @param rate Is the interest rate.
         * @param nper Is the total number of payment periods.
         * @param pv Is the present value.
         * @param startPeriod Is the first period in the calculation.
         * @param endPeriod Is the last period in the calculation.
         * @param type Is the timing of the payment.
         *
         * [Api set: ExcelApi 1.2]
         */
        cumIPmt(rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startPeriod: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, endPeriod: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, type: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the cumulative principal paid on a loan between two periods.
         *
         * @param rate Is the interest rate.
         * @param nper Is the total number of payment periods.
         * @param pv Is the present value.
         * @param startPeriod Is the first period in the calculation.
         * @param endPeriod Is the last period in the calculation.
         * @param type Is the timing of the payment.
         *
         * [Api set: ExcelApi 1.2]
         */
        cumPrinc(rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startPeriod: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, endPeriod: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, type: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Averages the values in a column in a list or database that match conditions you specify.
         *
         * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
         * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
         * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        daverage(database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Counts the cells containing numbers in the field (column) of records in the database that match the conditions you specify.
         *
         * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
         * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
         * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        dcount(database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Counts nonblank cells in the field (column) of records in the database that match the conditions you specify.
         *
         * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
         * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
         * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        dcountA(database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Extracts from a database a single record that matches the conditions you specify.
         *
         * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
         * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
         * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        dget(database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number | boolean | string>;
        /**
         *
         * Returns the largest number in the field (column) of records in the database that match the conditions you specify.
         *
         * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
         * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
         * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        dmax(database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the smallest number in the field (column) of records in the database that match the conditions you specify.
         *
         * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
         * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
         * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        dmin(database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Multiplies the values in the field (column) of records in the database that match the conditions you specify.
         *
         * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
         * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
         * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        dproduct(database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Estimates the standard deviation based on a sample from selected database entries.
         *
         * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
         * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
         * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        dstDev(database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Calculates the standard deviation based on the entire population of selected database entries.
         *
         * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
         * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
         * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        dstDevP(database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Adds the numbers in the field (column) of records in the database that match the conditions you specify.
         *
         * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
         * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
         * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        dsum(database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Estimates variance based on a sample from selected database entries.
         *
         * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
         * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
         * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        dvar(database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Calculates variance based on the entire population of selected database entries.
         *
         * @param database Is the range of cells that makes up the list or database. A database is a list of related data.
         * @param field Is either the label of the column in double quotation marks or a number that represents the column's position in the list.
         * @param criteria Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition.
         *
         * [Api set: ExcelApi 1.2]
         */
        dvarP(database: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, field: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number that represents the date in Microsoft Excel date-time code.
         *
         * @param year Is a number from 1900 or 1904 (depending on the workbook's date system) to 9999.
         * @param month Is a number from 1 to 12 representing the month of the year.
         * @param day Is a number from 1 to 31 representing the day of the month.
         *
         * [Api set: ExcelApi 1.2]
         */
        date(year: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, month: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, day: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a date in the form of text to a number that represents the date in Microsoft Excel date-time code.
         *
         * @param dateText Is text that represents a date in a Microsoft Excel date format, between 1/1/1900 or 1/1/1904 (depending on the workbook's date system) and 12/31/9999.
         *
         * [Api set: ExcelApi 1.2]
         */
        datevalue(dateText: string | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the day of the month, a number from 1 to 31.
         *
         * @param serialNumber Is a number in the date-time code used by Microsoft Excel.
         *
         * [Api set: ExcelApi 1.2]
         */
        day(serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of days between the two dates.
         *
         * @param endDate startDate and endDate are the two dates between which you want to know the number of days.
         * @param startDate startDate and endDate are the two dates between which you want to know the number of days.
         *
         * [Api set: ExcelApi 1.2]
         */
        days(endDate: string | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startDate: string | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of days between two dates based on a 360-day year (twelve 30-day months).
         *
         * @param startDate startDate and endDate are the two dates between which you want to know the number of days.
         * @param endDate startDate and endDate are the two dates between which you want to know the number of days.
         * @param method Is a logical value specifying the calculation method: U.S. (NASD) = FALSE or omitted; European = TRUE.
         *
         * [Api set: ExcelApi 1.2]
         */
        days360(startDate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, endDate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, method?: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the depreciation of an asset for a specified period using the fixed-declining balance method.
         *
         * @param cost Is the initial cost of the asset.
         * @param salvage Is the salvage value at the end of the life of the asset.
         * @param life Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
         * @param period Is the period for which you want to calculate the depreciation. Period must use the same units as Life.
         * @param month Is the number of months in the first year. If month is omitted, it is assumed to be 12.
         *
         * [Api set: ExcelApi 1.2]
         */
        db(cost: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvage: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, life: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, period: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, month?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Changes half-width (single-byte) characters within a character string to full-width (double-byte) characters. Use with double-byte character sets (DBCS).
         *
         * @param text Is a text, or a reference to a cell containing a text.
         *
         * [Api set: ExcelApi 1.2]
         */
        dbcs(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the depreciation of an asset for a specified period using the double-declining balance method or some other method you specify.
         *
         * @param cost Is the initial cost of the asset.
         * @param salvage Is the salvage value at the end of the life of the asset.
         * @param life Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
         * @param period Is the period for which you want to calculate the depreciation. Period must use the same units as Life.
         * @param factor Is the rate at which the balance declines. If Factor is omitted, it is assumed to be 2 (the double-declining balance method).
         *
         * [Api set: ExcelApi 1.2]
         */
        ddb(cost: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvage: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, life: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, period: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, factor?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a decimal number to binary.
         *
         * @param number Is the decimal integer you want to convert.
         * @param places Is the number of characters to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        dec2Bin(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a decimal number to hexadecimal.
         *
         * @param number Is the decimal integer you want to convert.
         * @param places Is the number of characters to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        dec2Hex(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a decimal number to octal.
         *
         * @param number Is the decimal integer you want to convert.
         * @param places Is the number of characters to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        dec2Oct(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a text representation of a number in a given base into a decimal number.
         *
         * @param number Is the number that you want to convert.
         * @param radix Is the base Radix of the number you are converting.
         *
         * [Api set: ExcelApi 1.2]
         */
        decimal(number: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, radix: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts radians to degrees.
         *
         * @param angle Is the angle in radians that you want to convert.
         *
         * [Api set: ExcelApi 1.2]
         */
        degrees(angle: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Tests whether two numbers are equal.
         *
         * @param number1 Is the first number.
         * @param number2 Is the second number.
         *
         * [Api set: ExcelApi 1.2]
         */
        delta(number1: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, number2?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the sum of squares of deviations of data points from their sample mean.
         *
         * @param values List of parameters, whose elements are 1 to 255 arguments, or an array or array reference, on which you want DEVSQ to calculate.
         *
         * [Api set: ExcelApi 1.2]
         */
        devSq(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the discount rate for a security.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param pr Is the security's price per $100 face value.
         * @param redemption Is the security's redemption value per $100 face value.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        disc(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pr: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemption: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a number to text, using currency format.
         *
         * @param number Is a number, a reference to a cell containing a number, or a formula that evaluates to a number.
         * @param decimals Is the number of digits to the right of the decimal point. The number is rounded as necessary; if omitted, Decimals = 2.
         *
         * [Api set: ExcelApi 1.2]
         */
        dollar(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, decimals?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number.
         *
         * @param fractionalDollar Is a number expressed as a fraction.
         * @param fraction Is the integer to use in the denominator of the fraction.
         *
         * [Api set: ExcelApi 1.2]
         */
        dollarDe(fractionalDollar: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fraction: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction.
         *
         * @param decimalDollar Is a decimal number.
         * @param fraction Is the integer to use in the denominator of a fraction.
         *
         * [Api set: ExcelApi 1.2]
         */
        dollarFr(decimalDollar: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fraction: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the annual duration of a security with periodic interest payments.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param coupon Is the security's annual coupon rate.
         * @param yld Is the security's annual yield.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        duration(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, coupon: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, yld: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Rounds a number up, to the nearest integer or to the nearest multiple of significance.
         *
         * @param number Is the value you want to round.
         * @param significance Is the multiple to which you want to round.
         *
         * [Api set: ExcelApi 1.2]
         */
        ecma_Ceiling(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the serial number of the date that is the indicated number of months before or after the start date.
         *
         * @param startDate Is a serial date number that represents the start date.
         * @param months Is the number of months before or after startDate.
         *
         * [Api set: ExcelApi 1.2]
         */
        edate(startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, months: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the effective annual interest rate.
         *
         * @param nominalRate Is the nominal interest rate.
         * @param npery Is the number of compounding periods per year.
         *
         * [Api set: ExcelApi 1.2]
         */
        effect(nominalRate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, npery: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the serial number of the last day of the month before or after a specified number of months.
         *
         * @param startDate Is a serial date number that represents the start date.
         * @param months Is the number of months before or after the startDate.
         *
         * [Api set: ExcelApi 1.2]
         */
        eoMonth(startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, months: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the error function.
         *
         * @param lowerLimit Is the lower bound for integrating ERF.
         * @param upperLimit Is the upper bound for integrating ERF.
         *
         * [Api set: ExcelApi 1.2]
         */
        erf(lowerLimit: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, upperLimit?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the complementary error function.
         *
         * @param x Is the lower bound for integrating ERF.
         *
         * [Api set: ExcelApi 1.2]
         */
        erfC(x: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the complementary error function.
         *
         * @param X Is the lower bound for integrating ERFC.PRECISE.
         *
         * [Api set: ExcelApi 1.2]
         */
        erfC_Precise(X: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the error function.
         *
         * @param X Is the lower bound for integrating ERF.PRECISE.
         *
         * [Api set: ExcelApi 1.2]
         */
        erf_Precise(X: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns a number matching an error value.
         *
         * @param errorVal Is the error value for which you want the identifying number, and can be an actual error value or a reference to a cell containing an error value.
         *
         * [Api set: ExcelApi 1.2]
         */
        error_Type(errorVal: string | number | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Rounds a positive number up and negative number down to the nearest even integer.
         *
         * @param number Is the value to round.
         *
         * [Api set: ExcelApi 1.2]
         */
        even(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Checks whether two text strings are exactly the same, and returns TRUE or FALSE. EXACT is case-sensitive.
         *
         * @param text1 Is the first text string.
         * @param text2 Is the second text string.
         *
         * [Api set: ExcelApi 1.2]
         */
        exact(text1: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, text2: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<boolean>;
        /**
         *
         * Returns e raised to the power of a given number.
         *
         * @param number Is the exponent applied to the base e. The constant e equals 2.71828182845904, the base of the natural logarithm.
         *
         * [Api set: ExcelApi 1.2]
         */
        exp(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the exponential distribution.
         *
         * @param x Is the value of the function, a nonnegative number.
         * @param lambda Is the parameter value, a positive number.
         * @param cumulative Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        expon_Dist(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, lambda: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the future value of an initial principal after applying a series of compound interest rates.
         *
         * @param principal Is the present value.
         * @param schedule Is an array of interest rates to apply.
         *
         * [Api set: ExcelApi 1.2]
         */
        fvschedule(principal: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, schedule: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the (left-tailed) F probability distribution (degree of diversity) for two data sets.
         *
         * @param x Is the value at which to evaluate the function, a nonnegative number.
         * @param degFreedom1 Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
         * @param degFreedom2 Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
         * @param cumulative Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        f_Dist(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the (right-tailed) F probability distribution (degree of diversity) for two data sets.
         *
         * @param x Is the value at which to evaluate the function, a nonnegative number.
         * @param degFreedom1 Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
         * @param degFreedom2 Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
         *
         * [Api set: ExcelApi 1.2]
         */
        f_Dist_RT(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse of the (left-tailed) F probability distribution: if p = F.DIST(x,...), then F.INV(p,...) = x.
         *
         * @param probability Is a probability associated with the F cumulative distribution, a number between 0 and 1 inclusive.
         * @param degFreedom1 Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
         * @param degFreedom2 Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
         *
         * [Api set: ExcelApi 1.2]
         */
        f_Inv(probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse of the (right-tailed) F probability distribution: if p = F.DIST.RT(x,...), then F.INV.RT(p,...) = x.
         *
         * @param probability Is a probability associated with the F cumulative distribution, a number between 0 and 1 inclusive.
         * @param degFreedom1 Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
         * @param degFreedom2 Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10.
         *
         * [Api set: ExcelApi 1.2]
         */
        f_Inv_RT(probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom1: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom2: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the factorial of a number, equal to 1*2*3*...* Number.
         *
         * @param number Is the nonnegative number you want the factorial of.
         *
         * [Api set: ExcelApi 1.2]
         */
        fact(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the double factorial of a number.
         *
         * @param number Is the value for which to return the double factorial.
         *
         * [Api set: ExcelApi 1.2]
         */
        factDouble(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the logical value FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        false(): FunctionResult<boolean>;
        /**
         *
         * Returns the starting position of one text string within another text string. FIND is case-sensitive.
         *
         * @param findText Is the text you want to find. Use double quotes (empty text) to match the first character in withinText; wildcard characters not allowed.
         * @param withinText Is the text containing the text you want to find.
         * @param startNum Specifies the character at which to start the search. The first character in withinText is character number 1. If omitted, startNum = 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        find(findText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, withinText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startNum?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Finds the starting position of one text string within another text string. FINDB is case-sensitive. Use with double-byte character sets (DBCS).
         *
         * @param findText Is the text you want to find.
         * @param withinText Is the text containing the text you want to find.
         * @param startNum Specifies the character at which to start the search.
         *
         * [Api set: ExcelApi 1.2]
         */
        findB(findText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, withinText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startNum?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the Fisher transformation.
         *
         * @param x Is the value for which you want the transformation, a number between -1 and 1, excluding -1 and 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        fisher(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse of the Fisher transformation: if y = FISHER(x), then FISHERINV(y) = x.
         *
         * @param y Is the value for which you want to perform the inverse of the transformation.
         *
         * [Api set: ExcelApi 1.2]
         */
        fisherInv(y: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Rounds a number to the specified number of decimals and returns the result as text with or without commas.
         *
         * @param number Is the number you want to round and convert to text.
         * @param decimals Is the number of digits to the right of the decimal point. If omitted, Decimals = 2.
         * @param noCommas Is a logical value: do not display commas in the returned text = TRUE; do display commas in the returned text = FALSE or omitted.
         *
         * [Api set: ExcelApi 1.2]
         */
        fixed(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, decimals?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, noCommas?: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Rounds a number down, to the nearest integer or to the nearest multiple of significance.
         *
         * @param number Is the value you want to round.
         * @param significance Is the multiple to which you want to round.
         * @param mode When given and nonzero this function will round towards zero.
         *
         * [Api set: ExcelApi 1.2]
         */
        floor_Math(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mode?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Rounds a number down, to the nearest integer or to the nearest multiple of significance.
         *
         * @param number Is the numeric value you want to round.
         * @param significance Is the multiple to which you want to round.
         *
         * [Api set: ExcelApi 1.2]
         */
        floor_Precise(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the future value of an investment based on periodic, constant payments and a constant interest rate.
         *
         * @param rate Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
         * @param nper Is the total number of payment periods in the investment.
         * @param pmt Is the payment made each period; it cannot change over the life of the investment.
         * @param pv Is the present value, or the lump-sum amount that a series of future payments is worth now. If omitted, Pv = 0.
         * @param type Is a value representing the timing of payment: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
         *
         * [Api set: ExcelApi 1.2]
         */
        fv(rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pmt: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, type?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the Gamma function value.
         *
         * @param x Is the value for which you want to calculate Gamma.
         *
         * [Api set: ExcelApi 1.2]
         */
        gamma(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the natural logarithm of the gamma function.
         *
         * @param x Is the value for which you want to calculate GAMMALN, a positive number.
         *
         * [Api set: ExcelApi 1.2]
         */
        gammaLn(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the natural logarithm of the gamma function.
         *
         * @param x Is the value for which you want to calculate GAMMALN.PRECISE, a positive number.
         *
         * [Api set: ExcelApi 1.2]
         */
        gammaLn_Precise(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the gamma distribution.
         *
         * @param x Is the value at which you want to evaluate the distribution, a nonnegative number.
         * @param alpha Is a parameter to the distribution, a positive number.
         * @param beta Is a parameter to the distribution, a positive number. If beta = 1, GAMMA.DIST returns the standard gamma distribution.
         * @param cumulative Is a logical value: return the cumulative distribution function = TRUE; return the probability mass function = FALSE or omitted.
         *
         * [Api set: ExcelApi 1.2]
         */
        gamma_Dist(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, beta: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse of the gamma cumulative distribution: if p = GAMMA.DIST(x,...), then GAMMA.INV(p,...) = x.
         *
         * @param probability Is the probability associated with the gamma distribution, a number between 0 and 1, inclusive.
         * @param alpha Is a parameter to the distribution, a positive number.
         * @param beta Is a parameter to the distribution, a positive number. If beta = 1, GAMMA.INV returns the inverse of the standard gamma distribution.
         *
         * [Api set: ExcelApi 1.2]
         */
        gamma_Inv(probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, beta: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns 0.5 less than the standard normal cumulative distribution.
         *
         * @param x Is the value for which you want the distribution.
         *
         * [Api set: ExcelApi 1.2]
         */
        gauss(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the greatest common divisor.
         *
         * @param values List of parameters, whose elements are 1 to 255 values.
         *
         * [Api set: ExcelApi 1.2]
         */
        gcd(...values: Array<number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Tests whether a number is greater than a threshold value.
         *
         * @param number Is the value to test against step.
         * @param step Is the threshold value.
         *
         * [Api set: ExcelApi 1.2]
         */
        geStep(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, step?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the geometric mean of an array or range of positive numeric data.
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the mean.
         *
         * [Api set: ExcelApi 1.2]
         */
        geoMean(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Looks for a value in the top row of a table or array of values and returns the value in the same column from a row you specify.
         *
         * @param lookupValue Is the value to be found in the first row of the table and can be a value, a reference, or a text string.
         * @param tableArray Is a table of text, numbers, or logical values in which data is looked up. tableArray can be a reference to a range or a range name.
         * @param rowIndexNum Is the row number in tableArray from which the matching value should be returned. The first row of values in the table is row 1.
         * @param rangeLookup Is a logical value: to find the closest match in the top row (sorted in ascending order) = TRUE or omitted; find an exact match = FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        hlookup(lookupValue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, tableArray: Excel.Range | number | Excel.RangeReference | Excel.FunctionResult<any>, rowIndexNum: Excel.Range | number | Excel.RangeReference | Excel.FunctionResult<any>, rangeLookup?: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number | string | boolean>;
        /**
         *
         * Returns the harmonic mean of a data set of positive numbers: the reciprocal of the arithmetic mean of reciprocals.
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the harmonic mean.
         *
         * [Api set: ExcelApi 1.2]
         */
        harMean(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Converts a Hexadecimal number to binary.
         *
         * @param number Is the hexadecimal number you want to convert.
         * @param places Is the number of characters to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        hex2Bin(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a hexadecimal number to decimal.
         *
         * @param number Is the hexadecimal number you want to convert.
         *
         * [Api set: ExcelApi 1.2]
         */
        hex2Dec(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a hexadecimal number to octal.
         *
         * @param number Is the hexadecimal number you want to convert.
         * @param places Is the number of characters to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        hex2Oct(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the hour as a number from 0 (12:00 A.M.) to 23 (11:00 P.M.).
         *
         * @param serialNumber Is a number in the date-time code used by Microsoft Excel, or text in time format, such as 16:48:00 or 4:48:00 PM.
         *
         * [Api set: ExcelApi 1.2]
         */
        hour(serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the hypergeometric distribution.
         *
         * @param sampleS Is the number of successes in the sample.
         * @param numberSample Is the size of the sample.
         * @param populationS Is the number of successes in the population.
         * @param numberPop Is the population size.
         * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        hypGeom_Dist(sampleS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberSample: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, populationS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberPop: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Creates a shortcut or jump that opens a document stored on your hard drive, a network server, or on the Internet.
         *
         * @param linkLocation Is the text giving the path and file name to the document to be opened, a hard drive location, UNC address, or URL path.
         * @param friendlyName Is text or a number that is displayed in the cell. If omitted, the cell displays the linkLocation text.
         *
         * [Api set: ExcelApi 1.2]
         */
        hyperlink(linkLocation: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, friendlyName?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number | string | boolean>;
        /**
         *
         * Rounds a number up, to the nearest integer or to the nearest multiple of significance.
         *
         * @param number Is the value you want to round.
         * @param significance Is the optional multiple to which you want to round.
         *
         * [Api set: ExcelApi 1.2]
         */
        iso_Ceiling(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE.
         *
         * @param logicalTest Is any value or expression that can be evaluated to TRUE or FALSE.
         * @param valueIfTrue Is the value that is returned if logicalTest is TRUE. If omitted, TRUE is returned. You can nest up to seven IF functions.
         * @param valueIfFalse Is the value that is returned if logicalTest is FALSE. If omitted, FALSE is returned.
         *
         * [Api set: ExcelApi 1.2]
         */
        if(logicalTest: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, valueIfTrue?: Excel.Range | number | string | boolean | Excel.RangeReference | Excel.FunctionResult<any>, valueIfFalse?: Excel.Range | number | string | boolean | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number | string | boolean>;
        /**
         *
         * Returns the absolute value (modulus) of a complex number.
         *
         * @param inumber Is a complex number for which you want the absolute value.
         *
         * [Api set: ExcelApi 1.2]
         */
        imAbs(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the argument q, an angle expressed in radians.
         *
         * @param inumber Is a complex number for which you want the argument.
         *
         * [Api set: ExcelApi 1.2]
         */
        imArgument(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the complex conjugate of a complex number.
         *
         * @param inumber Is a complex number for which you want the conjugate.
         *
         * [Api set: ExcelApi 1.2]
         */
        imConjugate(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the cosine of a complex number.
         *
         * @param inumber Is a complex number for which you want the cosine.
         *
         * [Api set: ExcelApi 1.2]
         */
        imCos(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the hyperbolic cosine of a complex number.
         *
         * @param inumber Is a complex number for which you want the hyperbolic cosine.
         *
         * [Api set: ExcelApi 1.2]
         */
        imCosh(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the cotangent of a complex number.
         *
         * @param inumber Is a complex number for which you want the cotangent.
         *
         * [Api set: ExcelApi 1.2]
         */
        imCot(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the cosecant of a complex number.
         *
         * @param inumber Is a complex number for which you want the cosecant.
         *
         * [Api set: ExcelApi 1.2]
         */
        imCsc(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the hyperbolic cosecant of a complex number.
         *
         * @param inumber Is a complex number for which you want the hyperbolic cosecant.
         *
         * [Api set: ExcelApi 1.2]
         */
        imCsch(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the quotient of two complex numbers.
         *
         * @param inumber1 Is the complex numerator or dividend.
         * @param inumber2 Is the complex denominator or divisor.
         *
         * [Api set: ExcelApi 1.2]
         */
        imDiv(inumber1: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, inumber2: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the exponential of a complex number.
         *
         * @param inumber Is a complex number for which you want the exponential.
         *
         * [Api set: ExcelApi 1.2]
         */
        imExp(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the natural logarithm of a complex number.
         *
         * @param inumber Is a complex number for which you want the natural logarithm.
         *
         * [Api set: ExcelApi 1.2]
         */
        imLn(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the base-10 logarithm of a complex number.
         *
         * @param inumber Is a complex number for which you want the common logarithm.
         *
         * [Api set: ExcelApi 1.2]
         */
        imLog10(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the base-2 logarithm of a complex number.
         *
         * @param inumber Is a complex number for which you want the base-2 logarithm.
         *
         * [Api set: ExcelApi 1.2]
         */
        imLog2(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns a complex number raised to an integer power.
         *
         * @param inumber Is a complex number you want to raise to a power.
         * @param number Is the power to which you want to raise the complex number.
         *
         * [Api set: ExcelApi 1.2]
         */
        imPower(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the product of 1 to 255 complex numbers.
         *
         * @param values Inumber1, Inumber2,... are from 1 to 255 complex numbers to multiply.
         *
         * [Api set: ExcelApi 1.2]
         */
        imProduct(...values: Array<Excel.Range | number | string | boolean | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the real coefficient of a complex number.
         *
         * @param inumber Is a complex number for which you want the real coefficient.
         *
         * [Api set: ExcelApi 1.2]
         */
        imReal(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the secant of a complex number.
         *
         * @param inumber Is a complex number for which you want the secant.
         *
         * [Api set: ExcelApi 1.2]
         */
        imSec(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the hyperbolic secant of a complex number.
         *
         * @param inumber Is a complex number for which you want the hyperbolic secant.
         *
         * [Api set: ExcelApi 1.2]
         */
        imSech(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the sine of a complex number.
         *
         * @param inumber Is a complex number for which you want the sine.
         *
         * [Api set: ExcelApi 1.2]
         */
        imSin(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the hyperbolic sine of a complex number.
         *
         * @param inumber Is a complex number for which you want the hyperbolic sine.
         *
         * [Api set: ExcelApi 1.2]
         */
        imSinh(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the square root of a complex number.
         *
         * @param inumber Is a complex number for which you want the square root.
         *
         * [Api set: ExcelApi 1.2]
         */
        imSqrt(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the difference of two complex numbers.
         *
         * @param inumber1 Is the complex number from which to subtract inumber2.
         * @param inumber2 Is the complex number to subtract from inumber1.
         *
         * [Api set: ExcelApi 1.2]
         */
        imSub(inumber1: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, inumber2: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the sum of complex numbers.
         *
         * @param values List of parameters, whose elements are from 1 to 255 complex numbers to add.
         *
         * [Api set: ExcelApi 1.2]
         */
        imSum(...values: Array<Excel.Range | number | string | boolean | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the tangent of a complex number.
         *
         * @param inumber Is a complex number for which you want the tangent.
         *
         * [Api set: ExcelApi 1.2]
         */
        imTan(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the imaginary coefficient of a complex number.
         *
         * @param inumber Is a complex number for which you want the imaginary coefficient.
         *
         * [Api set: ExcelApi 1.2]
         */
        imaginary(inumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Rounds a number down to the nearest integer.
         *
         * @param number Is the real number you want to round down to an integer.
         *
         * [Api set: ExcelApi 1.2]
         */
        int(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the interest rate for a fully invested security.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param investment Is the amount invested in the security.
         * @param redemption Is the amount to be received at maturity.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        intRate(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, investment: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemption: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the interest payment for a given period for an investment, based on periodic, constant payments and a constant interest rate.
         *
         * @param rate Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
         * @param per Is the period for which you want to find the interest and must be in the range 1 to Nper.
         * @param nper Is the total number of payment periods in an investment.
         * @param pv Is the present value, or the lump-sum amount that a series of future payments is worth now.
         * @param fv Is the future value, or a cash balance you want to attain after the last payment is made. If omitted, Fv = 0.
         * @param type Is a logical value representing the timing of payment: at the end of the period = 0 or omitted, at the beginning of the period = 1.
         *
         * [Api set: ExcelApi 1.2]
         */
        ipmt(rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, per: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, type?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the internal rate of return for a series of cash flows.
         *
         * @param values Is an array or a reference to cells that contain numbers for which you want to calculate the internal rate of return.
         * @param guess Is a number that you guess is close to the result of IRR; 0.1 (10 percent) if omitted.
         *
         * [Api set: ExcelApi 1.2]
         */
        irr(values: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, guess?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Checks whether a value is an error (#VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!) excluding #N/A, and returns TRUE or FALSE.
         *
         * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
         *
         * [Api set: ExcelApi 1.2]
         */
        isErr(value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<boolean>;
        /**
         *
         * Checks whether a value is an error (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!), and returns TRUE or FALSE.
         *
         * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
         *
         * [Api set: ExcelApi 1.2]
         */
        isError(value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<boolean>;
        /**
         *
         * Returns TRUE if the number is even.
         *
         * @param number Is the value to test.
         *
         * [Api set: ExcelApi 1.2]
         */
        isEven(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Checks whether a reference is to a cell containing a formula, and returns TRUE or FALSE.
         *
         * @param reference Is a reference to the cell you want to test.  Reference can be a cell reference, a formula, or name that refers to a cell.
         *
         * [Api set: ExcelApi 1.2]
         */
        isFormula(reference: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<boolean>;
        /**
         *
         * Checks whether a value is a logical value (TRUE or FALSE), and returns TRUE or FALSE.
         *
         * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
         *
         * [Api set: ExcelApi 1.2]
         */
        isLogical(value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<boolean>;
        /**
         *
         * Checks whether a value is #N/A, and returns TRUE or FALSE.
         *
         * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
         *
         * [Api set: ExcelApi 1.2]
         */
        isNA(value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<boolean>;
        /**
         *
         * Checks whether a value is not text (blank cells are not text), and returns TRUE or FALSE.
         *
         * @param value Is the value you want tested: a cell; a formula; or a name referring to a cell, formula, or value.
         *
         * [Api set: ExcelApi 1.2]
         */
        isNonText(value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<boolean>;
        /**
         *
         * Checks whether a value is a number, and returns TRUE or FALSE.
         *
         * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
         *
         * [Api set: ExcelApi 1.2]
         */
        isNumber(value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<boolean>;
        /**
         *
         * Returns TRUE if the number is odd.
         *
         * @param number Is the value to test.
         *
         * [Api set: ExcelApi 1.2]
         */
        isOdd(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Checks whether a value is text, and returns TRUE or FALSE.
         *
         * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
         *
         * [Api set: ExcelApi 1.2]
         */
        isText(value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<boolean>;
        /**
         *
         * Returns the ISO week number in the year for a given date.
         *
         * @param date Is the date-time code used by Microsoft Excel for date and time calculation.
         *
         * [Api set: ExcelApi 1.2]
         */
        isoWeekNum(date: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the interest paid during a specific period of an investment.
         *
         * @param rate Interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
         * @param per Period for which you want to find the interest.
         * @param nper Number of payment periods in an investment.
         * @param pv Lump sum amount that a series of future payments is right now.
         *
         * [Api set: ExcelApi 1.2]
         */
        ispmt(rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, per: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Checks whether a value is a reference, and returns TRUE or FALSE.
         *
         * @param value Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value.
         *
         * [Api set: ExcelApi 1.2]
         */
        isref(value: Excel.Range | number | string | boolean | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<boolean>;
        /**
         *
         * Returns the kurtosis of a data set.
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the kurtosis.
         *
         * [Api set: ExcelApi 1.2]
         */
        kurt(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the k-th largest value in a data set. For example, the fifth largest number.
         *
         * @param array Is the array or range of data for which you want to determine the k-th largest value.
         * @param k Is the position (from the largest) in the array or cell range of the value to return.
         *
         * [Api set: ExcelApi 1.2]
         */
        large(array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, k: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the least common multiple.
         *
         * @param values List of parameters, whose elements are 1 to 255 values for which you want the least common multiple.
         *
         * [Api set: ExcelApi 1.2]
         */
        lcm(...values: Array<number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the specified number of characters from the start of a text string.
         *
         * @param text Is the text string containing the characters you want to extract.
         * @param numChars Specifies how many characters you want LEFT to extract; 1 if omitted.
         *
         * [Api set: ExcelApi 1.2]
         */
        left(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numChars?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the specified number of characters from the start of a text string. Use with double-byte character sets (DBCS).
         *
         * @param text Is the text string containing the characters you want to extract.
         * @param numBytes Specifies how many characters you want LEFT to return.
         *
         * [Api set: ExcelApi 1.2]
         */
        leftb(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numBytes?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the number of characters in a text string.
         *
         * @param text Is the text whose length you want to find. Spaces count as characters.
         *
         * [Api set: ExcelApi 1.2]
         */
        len(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of characters in a text string. Use with double-byte character sets (DBCS).
         *
         * @param text Is the text whose length you want to find.
         *
         * [Api set: ExcelApi 1.2]
         */
        lenb(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the natural logarithm of a number.
         *
         * @param number Is the positive real number for which you want the natural logarithm.
         *
         * [Api set: ExcelApi 1.2]
         */
        ln(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the logarithm of a number to the base you specify.
         *
         * @param number Is the positive real number for which you want the logarithm.
         * @param base Is the base of the logarithm; 10 if omitted.
         *
         * [Api set: ExcelApi 1.2]
         */
        log(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, base?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the base-10 logarithm of a number.
         *
         * @param number Is the positive real number for which you want the base-10 logarithm.
         *
         * [Api set: ExcelApi 1.2]
         */
        log10(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the lognormal distribution of x, where ln(x) is normally distributed with parameters Mean and Standard_dev.
         *
         * @param x Is the value at which to evaluate the function, a positive number.
         * @param mean Is the mean of ln(x).
         * @param standardDev Is the standard deviation of ln(x), a positive number.
         * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        logNorm_Dist(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mean: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse of the lognormal cumulative distribution function of x, where ln(x) is normally distributed with parameters Mean and Standard_dev.
         *
         * @param probability Is a probability associated with the lognormal distribution, a number between 0 and 1, inclusive.
         * @param mean Is the mean of ln(x).
         * @param standardDev Is the standard deviation of ln(x), a positive number.
         *
         * [Api set: ExcelApi 1.2]
         */
        logNorm_Inv(probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mean: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Looks up a value either from a one-row or one-column range or from an array. Provided for backward compatibility.
         *
         * @param lookupValue Is a value that LOOKUP searches for in lookupVector and can be a number, text, a logical value, or a name or reference to a value.
         * @param lookupVector Is a range that contains only one row or one column of text, numbers, or logical values, placed in ascending order.
         * @param resultVector Is a range that contains only one row or column, the same size as lookupVector.
         *
         * [Api set: ExcelApi 1.2]
         */
        lookup(lookupValue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, lookupVector: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, resultVector?: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number | string | boolean>;
        /**
         *
         * Converts all letters in a text string to lowercase.
         *
         * @param text Is the text you want to convert to lowercase. Characters in Text that are not letters are not changed.
         *
         * [Api set: ExcelApi 1.2]
         */
        lower(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the Macauley modified duration for a security with an assumed par value of $100.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param coupon Is the security's annual coupon rate.
         * @param yld Is the security's annual yield.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        mduration(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, coupon: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, yld: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the internal rate of return for a series of periodic cash flows, considering both cost of investment and interest on reinvestment of cash.
         *
         * @param values Is an array or a reference to cells that contain numbers that represent a series of payments (negative) and income (positive) at regular periods.
         * @param financeRate Is the interest rate you pay on the money used in the cash flows.
         * @param reinvestRate Is the interest rate you receive on the cash flows as you reinvest them.
         *
         * [Api set: ExcelApi 1.2]
         */
        mirr(values: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, financeRate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, reinvestRate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns a number rounded to the desired multiple.
         *
         * @param number Is the value to round.
         * @param multiple Is the multiple to which you want to round number.
         *
         * [Api set: ExcelApi 1.2]
         */
        mround(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, multiple: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the relative position of an item in an array that matches a specified value in a specified order.
         *
         * @param lookupValue Is the value you use to find the value you want in the array, a number, text, or logical value, or a reference to one of these.
         * @param lookupArray Is a contiguous range of cells containing possible lookup values, an array of values, or a reference to an array.
         * @param matchType Is a number 1, 0, or -1 indicating which value to return.
         *
         * [Api set: ExcelApi 1.2]
         */
        match(lookupValue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, lookupArray: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, matchType?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the largest value in a set of values. Ignores logical values and text.
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers, empty cells, logical values, or text numbers for which you want the maximum.
         *
         * [Api set: ExcelApi 1.2]
         */
        max(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the largest value in a set of values. Does not ignore logical values and text.
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers, empty cells, logical values, or text numbers for which you want the maximum.
         *
         * [Api set: ExcelApi 1.2]
         */
        maxA(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the median, or the number in the middle of the set of given numbers.
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the median.
         *
         * [Api set: ExcelApi 1.2]
         */
        median(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the characters from the middle of a text string, given a starting position and length.
         *
         * @param text Is the text string from which you want to extract the characters.
         * @param startNum Is the position of the first character you want to extract. The first character in Text is 1.
         * @param numChars Specifies how many characters to return from Text.
         *
         * [Api set: ExcelApi 1.2]
         */
        mid(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numChars: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns characters from the middle of a text string, given a starting position and length. Use with double-byte character sets (DBCS).
         *
         * @param text Is the text string containing the characters you want to extract.
         * @param startNum Is the position of the first character you want to extract in text.
         * @param numBytes Specifies how many characters to return from text.
         *
         * [Api set: ExcelApi 1.2]
         */
        midb(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numBytes: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the smallest number in a set of values. Ignores logical values and text.
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers, empty cells, logical values, or text numbers for which you want the minimum.
         *
         * [Api set: ExcelApi 1.2]
         */
        min(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the smallest value in a set of values. Does not ignore logical values and text.
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers, empty cells, logical values, or text numbers for which you want the minimum.
         *
         * [Api set: ExcelApi 1.2]
         */
        minA(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the minute, a number from 0 to 59.
         *
         * @param serialNumber Is a number in the date-time code used by Microsoft Excel or text in time format, such as 16:48:00 or 4:48:00 PM.
         *
         * [Api set: ExcelApi 1.2]
         */
        minute(serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the remainder after a number is divided by a divisor.
         *
         * @param number Is the number for which you want to find the remainder after the division is performed.
         * @param divisor Is the number by which you want to divide Number.
         *
         * [Api set: ExcelApi 1.2]
         */
        mod(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, divisor: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the month, a number from 1 (January) to 12 (December).
         *
         * @param serialNumber Is a number in the date-time code used by Microsoft Excel.
         *
         * [Api set: ExcelApi 1.2]
         */
        month(serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the multinomial of a set of numbers.
         *
         * @param values List of parameters, whose elements are 1 to 255 values for which you want the multinomial.
         *
         * [Api set: ExcelApi 1.2]
         */
        multiNomial(...values: Array<number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Converts non-number value to a number, dates to serial numbers, TRUE to 1, anything else to 0 (zero).
         *
         * @param value Is the value you want converted.
         *
         * [Api set: ExcelApi 1.2]
         */
        n(value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate.
         *
         * @param rate Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
         * @param pmt Is the payment made each period; it cannot change over the life of the investment.
         * @param pv Is the present value, or the lump-sum amount that a series of future payments is worth now.
         * @param fv Is the future value, or a cash balance you want to attain after the last payment is made. If omitted, zero is used.
         * @param type Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
         *
         * [Api set: ExcelApi 1.2]
         */
        nper(rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pmt: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, type?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the error value #N/A (value not available).
         *
         * [Api set: ExcelApi 1.2]
         */
        na(): FunctionResult<number | string>;
        /**
         *
         * Returns the negative binomial distribution, the probability that there will be Number_f failures before the Number_s-th success, with Probability_s probability of a success.
         *
         * @param numberF Is the number of failures.
         * @param numberS Is the threshold number of successes.
         * @param probabilityS Is the probability of a success; a number between 0 and 1.
         * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability mass function, use FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        negBinom_Dist(numberF: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, probabilityS: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of whole workdays between two dates.
         *
         * @param startDate Is a serial date number that represents the start date.
         * @param endDate Is a serial date number that represents the end date.
         * @param holidays Is an optional set of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays.
         *
         * [Api set: ExcelApi 1.2]
         */
        networkDays(startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, endDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, holidays?: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of whole workdays between two dates with custom weekend parameters.
         *
         * @param startDate Is a serial date number that represents the start date.
         * @param endDate Is a serial date number that represents the end date.
         * @param weekend Is a number or string specifying when weekends occur.
         * @param holidays Is an optional set of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays.
         *
         * [Api set: ExcelApi 1.2]
         */
        networkDays_Intl(startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, endDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, weekend?: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, holidays?: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the annual nominal interest rate.
         *
         * @param effectRate Is the effective interest rate.
         * @param npery Is the number of compounding periods per year.
         *
         * [Api set: ExcelApi 1.2]
         */
        nominal(effectRate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, npery: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the normal distribution for the specified mean and standard deviation.
         *
         * @param x Is the value for which you want the distribution.
         * @param mean Is the arithmetic mean of the distribution.
         * @param standardDev Is the standard deviation of the distribution, a positive number.
         * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        norm_Dist(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mean: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.
         *
         * @param probability Is a probability corresponding to the normal distribution, a number between 0 and 1 inclusive.
         * @param mean Is the arithmetic mean of the distribution.
         * @param standardDev Is the standard deviation of the distribution, a positive number.
         *
         * [Api set: ExcelApi 1.2]
         */
        norm_Inv(probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mean: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the standard normal distribution (has a mean of zero and a standard deviation of one).
         *
         * @param z Is the value for which you want the distribution.
         * @param cumulative Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        norm_S_Dist(z: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the inverse of the standard normal cumulative distribution (has a mean of zero and a standard deviation of one).
         *
         * @param probability Is a probability corresponding to the normal distribution, a number between 0 and 1 inclusive.
         *
         * [Api set: ExcelApi 1.2]
         */
        norm_S_Inv(probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Changes FALSE to TRUE, or TRUE to FALSE.
         *
         * @param logical Is a value or expression that can be evaluated to TRUE or FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        not(logical: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<boolean>;
        /**
         *
         * Returns the current date and time formatted as a date and time.
         *
         * [Api set: ExcelApi 1.2]
         */
        now(): FunctionResult<number>;
        /**
         *
         * Returns the net present value of an investment based on a discount rate and a series of future payments (negative values) and income (positive values).
         *
         * @param rate Is the rate of discount over the length of one period.
         * @param values List of parameters, whose elements are 1 to 254 payments and income, equally spaced in time and occurring at the end of each period.
         *
         * [Api set: ExcelApi 1.2]
         */
        npv(rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Converts text to number in a locale-independent manner.
         *
         * @param text Is the string representing the number you want to convert.
         * @param decimalSeparator Is the character used as the decimal separator in the string.
         * @param groupSeparator Is the character used as the group separator in the string.
         *
         * [Api set: ExcelApi 1.2]
         */
        numberValue(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, decimalSeparator?: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, groupSeparator?: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts an octal number to binary.
         *
         * @param number Is the octal number you want to convert.
         * @param places Is the number of characters to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        oct2Bin(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts an octal number to decimal.
         *
         * @param number Is the octal number you want to convert.
         *
         * [Api set: ExcelApi 1.2]
         */
        oct2Dec(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts an octal number to hexadecimal.
         *
         * @param number Is the octal number you want to convert.
         * @param places Is the number of characters to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        oct2Hex(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, places?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Rounds a positive number up and negative number down to the nearest odd integer.
         *
         * @param number Is the value to round.
         *
         * [Api set: ExcelApi 1.2]
         */
        odd(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the price per $100 face value of a security with an odd first period.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param issue Is the security's issue date, expressed as a serial date number.
         * @param firstCoupon Is the security's first coupon date, expressed as a serial date number.
         * @param rate Is the security's interest rate.
         * @param yld Is the security's annual yield.
         * @param redemption Is the security's redemption value per $100 face value.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        oddFPrice(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, issue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, firstCoupon: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, yld: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemption: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the yield of a security with an odd first period.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param issue Is the security's issue date, expressed as a serial date number.
         * @param firstCoupon Is the security's first coupon date, expressed as a serial date number.
         * @param rate Is the security's interest rate.
         * @param pr Is the security's price.
         * @param redemption Is the security's redemption value per $100 face value.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        oddFYield(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, issue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, firstCoupon: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pr: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemption: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the price per $100 face value of a security with an odd last period.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param lastInterest Is the security's last coupon date, expressed as a serial date number.
         * @param rate Is the security's interest rate.
         * @param yld Is the security's annual yield.
         * @param redemption Is the security's redemption value per $100 face value.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        oddLPrice(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, lastInterest: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, yld: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemption: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the yield of a security with an odd last period.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param lastInterest Is the security's last coupon date, expressed as a serial date number.
         * @param rate Is the security's interest rate.
         * @param pr Is the security's price.
         * @param redemption Is the security's redemption value per $100 face value.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        oddLYield(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, lastInterest: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pr: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemption: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Checks whether any of the arguments are TRUE, and returns TRUE or FALSE. Returns FALSE only if all arguments are FALSE.
         *
         * @param values List of parameters, whose elements are 1 to 255 conditions that you want to test that can be either TRUE or FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        or(...values: Array<boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<boolean>;
        /**
         *
         * Returns the number of periods required by an investment to reach a specified value.
         *
         * @param rate Is the interest rate per period.
         * @param pv Is the present value of the investment.
         * @param fv Is the desired future value of the investment.
         *
         * [Api set: ExcelApi 1.2]
         */
        pduration(rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the rank of a value in a data set as a percentage of the data set as a percentage (0..1, exclusive) of the data set.
         *
         * @param array Is the array or range of data with numeric values that defines relative standing.
         * @param x Is the value for which you want to know the rank.
         * @param significance Is an optional value that identifies the number of significant digits for the returned percentage, three digits if omitted (0.xxx%).
         *
         * [Api set: ExcelApi 1.2]
         */
        percentRank_Exc(array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the rank of a value in a data set as a percentage of the data set as a percentage (0..1, inclusive) of the data set.
         *
         * @param array Is the array or range of data with numeric values that defines relative standing.
         * @param x Is the value for which you want to know the rank.
         * @param significance Is an optional value that identifies the number of significant digits for the returned percentage, three digits if omitted (0.xxx%).
         *
         * [Api set: ExcelApi 1.2]
         */
        percentRank_Inc(array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, significance?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive.
         *
         * @param array Is the array or range of data that defines relative standing.
         * @param k Is the percentile value that is between 0 through 1, inclusive.
         *
         * [Api set: ExcelApi 1.2]
         */
        percentile_Exc(array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, k: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the k-th percentile of values in a range, where k is in the range 0..1, inclusive.
         *
         * @param array Is the array or range of data that defines relative standing.
         * @param k Is the percentile value that is between 0 through 1, inclusive.
         *
         * [Api set: ExcelApi 1.2]
         */
        percentile_Inc(array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, k: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of permutations for a given number of objects that can be selected from the total objects.
         *
         * @param number Is the total number of objects.
         * @param numberChosen Is the number of objects in each permutation.
         *
         * [Api set: ExcelApi 1.2]
         */
        permut(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberChosen: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects.
         *
         * @param number Is the total number of objects.
         * @param numberChosen Is the number of objects in each permutation.
         *
         * [Api set: ExcelApi 1.2]
         */
        permutationa(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberChosen: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the value of the density function for a standard normal distribution.
         *
         * @param x Is the number for which you want the density of the standard normal distribution.
         *
         * [Api set: ExcelApi 1.2]
         */
        phi(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the value of Pi, 3.14159265358979, accurate to 15 digits.
         *
         * [Api set: ExcelApi 1.2]
         */
        pi(): FunctionResult<number>;
        /**
         *
         * Calculates the payment for a loan based on constant payments and a constant interest rate.
         *
         * @param rate Is the interest rate per period for the loan. For example, use 6%/4 for quarterly payments at 6% APR.
         * @param nper Is the total number of payments for the loan.
         * @param pv Is the present value: the total amount that a series of future payments is worth now.
         * @param fv Is the future value, or a cash balance you want to attain after the last payment is made, 0 (zero) if omitted.
         * @param type Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
         *
         * [Api set: ExcelApi 1.2]
         */
        pmt(rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, type?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the Poisson distribution.
         *
         * @param x Is the number of events.
         * @param mean Is the expected numeric value, a positive number.
         * @param cumulative Is a logical value: for the cumulative Poisson probability, use TRUE; for the Poisson probability mass function, use FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        poisson_Dist(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mean: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the result of a number raised to a power.
         *
         * @param number Is the base number, any real number.
         * @param power Is the exponent, to which the base number is raised.
         *
         * [Api set: ExcelApi 1.2]
         */
        power(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, power: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the payment on the principal for a given investment based on periodic, constant payments and a constant interest rate.
         *
         * @param rate Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
         * @param per Specifies the period and must be in the range 1 to nper.
         * @param nper Is the total number of payment periods in an investment.
         * @param pv Is the present value: the total amount that a series of future payments is worth now.
         * @param fv Is the future value, or cash balance you want to attain after the last payment is made.
         * @param type Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
         *
         * [Api set: ExcelApi 1.2]
         */
        ppmt(rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, per: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, type?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the price per $100 face value of a security that pays periodic interest.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param rate Is the security's annual coupon rate.
         * @param yld Is the security's annual yield.
         * @param redemption Is the security's redemption value per $100 face value.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        price(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, yld: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemption: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the price per $100 face value of a discounted security.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param discount Is the security's discount rate.
         * @param redemption Is the security's redemption value per $100 face value.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        priceDisc(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, discount: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemption: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the price per $100 face value of a security that pays interest at maturity.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param issue Is the security's issue date, expressed as a serial date number.
         * @param rate Is the security's interest rate at date of issue.
         * @param yld Is the security's annual yield.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        priceMat(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, issue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, yld: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Multiplies all the numbers given as arguments.
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers, logical values, or text representations of numbers that you want to multiply.
         *
         * [Api set: ExcelApi 1.2]
         */
        product(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Converts a text string to proper case; the first letter in each word to uppercase, and all other letters to lowercase.
         *
         * @param text Is text enclosed in quotation marks, a formula that returns text, or a reference to a cell containing text to partially capitalize.
         *
         * [Api set: ExcelApi 1.2]
         */
        proper(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the present value of an investment: the total amount that a series of future payments is worth now.
         *
         * @param rate Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR.
         * @param nper Is the total number of payment periods in an investment.
         * @param pmt Is the payment made each period and cannot change over the life of the investment.
         * @param fv Is the future value, or a cash balance you want to attain after the last payment is made.
         * @param type Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
         *
         * [Api set: ExcelApi 1.2]
         */
        pv(rate: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pmt: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, type?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the quartile of a data set, based on percentile values from 0..1, exclusive.
         *
         * @param array Is the array or cell range of numeric values for which you want the quartile value.
         * @param quart Is a number: minimum value = 0; 1st quartile = 1; median value = 2; 3rd quartile = 3; maximum value = 4.
         *
         * [Api set: ExcelApi 1.2]
         */
        quartile_Exc(array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, quart: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the quartile of a data set, based on percentile values from 0..1, inclusive.
         *
         * @param array Is the array or cell range of numeric values for which you want the quartile value.
         * @param quart Is a number: minimum value = 0; 1st quartile = 1; median value = 2; 3rd quartile = 3; maximum value = 4.
         *
         * [Api set: ExcelApi 1.2]
         */
        quartile_Inc(array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, quart: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the integer portion of a division.
         *
         * @param numerator Is the dividend.
         * @param denominator Is the divisor.
         *
         * [Api set: ExcelApi 1.2]
         */
        quotient(numerator: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, denominator: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts degrees to radians.
         *
         * @param angle Is an angle in degrees that you want to convert.
         *
         * [Api set: ExcelApi 1.2]
         */
        radians(angle: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns a random number greater than or equal to 0 and less than 1, evenly distributed (changes on recalculation).
         *
         * [Api set: ExcelApi 1.2]
         */
        rand(): FunctionResult<number>;
        /**
         *
         * Returns a random number between the numbers you specify.
         *
         * @param bottom Is the smallest integer RANDBETWEEN will return.
         * @param top Is the largest integer RANDBETWEEN will return.
         *
         * [Api set: ExcelApi 1.2]
         */
        randBetween(bottom: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, top: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the rank of a number in a list of numbers: its size relative to other values in the list; if more than one value has the same rank, the average rank is returned.
         *
         * @param number Is the number for which you want to find the rank.
         * @param ref Is an array of, or a reference to, a list of numbers. Nonnumeric values are ignored.
         * @param order Is a number: rank in the list sorted descending = 0 or omitted; rank in the list sorted ascending = any nonzero value.
         *
         * [Api set: ExcelApi 1.2]
         */
        rank_Avg(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ref: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, order?: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the rank of a number in a list of numbers: its size relative to other values in the list; if more than one value has the same rank, the top rank of that set of values is returned.
         *
         * @param number Is the number for which you want to find the rank.
         * @param ref Is an array of, or a reference to, a list of numbers. Nonnumeric values are ignored.
         * @param order Is a number: rank in the list sorted descending = 0 or omitted; rank in the list sorted ascending = any nonzero value.
         *
         * [Api set: ExcelApi 1.2]
         */
        rank_Eq(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ref: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, order?: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the interest rate per period of a loan or an investment. For example, use 6%/4 for quarterly payments at 6% APR.
         *
         * @param nper Is the total number of payment periods for the loan or investment.
         * @param pmt Is the payment made each period and cannot change over the life of the loan or investment.
         * @param pv Is the present value: the total amount that a series of future payments is worth now.
         * @param fv Is the future value, or a cash balance you want to attain after the last payment is made. If omitted, uses Fv = 0.
         * @param type Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted.
         * @param guess Is your guess for what the rate will be; if omitted, Guess = 0.1 (10 percent).
         *
         * [Api set: ExcelApi 1.2]
         */
        rate(nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pmt: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, type?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, guess?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the amount received at maturity for a fully invested security.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param investment Is the amount invested in the security.
         * @param discount Is the security's discount rate.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        received(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, investment: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, discount: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Replaces part of a text string with a different text string.
         *
         * @param oldText Is text in which you want to replace some characters.
         * @param startNum Is the position of the character in oldText that you want to replace with newText.
         * @param numChars Is the number of characters in oldText that you want to replace.
         * @param newText Is the text that will replace characters in oldText.
         *
         * [Api set: ExcelApi 1.2]
         */
        replace(oldText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numChars: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, newText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Replaces part of a text string with a different text string. Use with double-byte character sets (DBCS).
         *
         * @param oldText Is text in which you want to replace some characters.
         * @param startNum Is the position of the character in oldText that you want to replace with newText.
         * @param numBytes Is the number of characters in oldText that you want to replace with newText.
         * @param newText Is the text that will replace characters in oldText.
         *
         * [Api set: ExcelApi 1.2]
         */
        replaceB(oldText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numBytes: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, newText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Repeats text a given number of times. Use REPT to fill a cell with a number of instances of a text string.
         *
         * @param text Is the text you want to repeat.
         * @param numberTimes Is a positive number specifying the number of times to repeat text.
         *
         * [Api set: ExcelApi 1.2]
         */
        rept(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numberTimes: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the specified number of characters from the end of a text string.
         *
         * @param text Is the text string that contains the characters you want to extract.
         * @param numChars Specifies how many characters you want to extract, 1 if omitted.
         *
         * [Api set: ExcelApi 1.2]
         */
        right(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numChars?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the specified number of characters from the end of a text string. Use with double-byte character sets (DBCS).
         *
         * @param text Is the text string containing the characters you want to extract.
         * @param numBytes Specifies how many characters you want to extract.
         *
         * [Api set: ExcelApi 1.2]
         */
        rightb(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numBytes?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Converts an Arabic numeral to Roman, as text.
         *
         * @param number Is the Arabic numeral you want to convert.
         * @param form Is the number specifying the type of Roman numeral you want.
         *
         * [Api set: ExcelApi 1.2]
         */
        roman(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, form?: boolean | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Rounds a number to a specified number of digits.
         *
         * @param number Is the number you want to round.
         * @param numDigits Is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero to the nearest integer.
         *
         * [Api set: ExcelApi 1.2]
         */
        round(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numDigits: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Rounds a number down, toward zero.
         *
         * @param number Is any real number that you want rounded down.
         * @param numDigits Is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero or omitted, to the nearest integer.
         *
         * [Api set: ExcelApi 1.2]
         */
        roundDown(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numDigits: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Rounds a number up, away from zero.
         *
         * @param number Is any real number that you want rounded up.
         * @param numDigits Is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero or omitted, to the nearest integer.
         *
         * [Api set: ExcelApi 1.2]
         */
        roundUp(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numDigits: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of rows in a reference or array.
         *
         * @param array Is an array, an array formula, or a reference to a range of cells for which you want the number of rows.
         *
         * [Api set: ExcelApi 1.2]
         */
        rows(array: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns an equivalent interest rate for the growth of an investment.
         *
         * @param nper Is the number of periods for the investment.
         * @param pv Is the present value of the investment.
         * @param fv Is the future value of the investment.
         *
         * [Api set: ExcelApi 1.2]
         */
        rri(nper: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, fv: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the secant of an angle.
         *
         * @param number Is the angle in radians for which you want the secant.
         *
         * [Api set: ExcelApi 1.2]
         */
        sec(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the hyperbolic secant of an angle.
         *
         * @param number Is the angle in radians for which you want the hyperbolic secant.
         *
         * [Api set: ExcelApi 1.2]
         */
        sech(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the second, a number from 0 to 59.
         *
         * @param serialNumber Is a number in the date-time code used by Microsoft Excel or text in time format, such as 16:48:23 or 4:48:47 PM.
         *
         * [Api set: ExcelApi 1.2]
         */
        second(serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the sum of a power series based on the formula.
         *
         * @param x Is the input value to the power series.
         * @param n Is the initial power to which you want to raise x.
         * @param m Is the step by which to increase n for each term in the series.
         * @param coefficients Is a set of coefficients by which each successive power of x is multiplied.
         *
         * [Api set: ExcelApi 1.2]
         */
        seriesSum(x: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, n: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, m: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, coefficients: Excel.Range | string | number | boolean | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the sheet number of the referenced sheet.
         *
         * @param value Is the name of a sheet or a reference that you want the sheet number of.  If omitted the number of the sheet containing the function is returned.
         *
         * [Api set: ExcelApi 1.2]
         */
        sheet(value?: Excel.Range | string | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the number of sheets in a reference.
         *
         * @param reference Is a reference for which you want to know the number of sheets it contains.  If omitted the number of sheets in the workbook containing the function is returned.
         *
         * [Api set: ExcelApi 1.2]
         */
        sheets(reference?: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the sign of a number: 1 if the number is positive, zero if the number is zero, or -1 if the number is negative.
         *
         * @param number Is any real number.
         *
         * [Api set: ExcelApi 1.2]
         */
        sign(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the sine of an angle.
         *
         * @param number Is the angle in radians for which you want the sine. Degrees * PI()/180 = radians.
         *
         * [Api set: ExcelApi 1.2]
         */
        sin(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the hyperbolic sine of a number.
         *
         * @param number Is any real number.
         *
         * [Api set: ExcelApi 1.2]
         */
        sinh(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the skewness of a distribution: a characterization of the degree of asymmetry of a distribution around its mean.
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the skewness.
         *
         * [Api set: ExcelApi 1.2]
         */
        skew(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the skewness of a distribution based on a population: a characterization of the degree of asymmetry of a distribution around its mean.
         *
         * @param values List of parameters, whose elements are 1 to 254 numbers or names, arrays, or references that contain numbers for which you want the population skewness.
         *
         * [Api set: ExcelApi 1.2]
         */
        skew_p(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the straight-line depreciation of an asset for one period.
         *
         * @param cost Is the initial cost of the asset.
         * @param salvage Is the salvage value at the end of the life of the asset.
         * @param life Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
         *
         * [Api set: ExcelApi 1.2]
         */
        sln(cost: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvage: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, life: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the k-th smallest value in a data set. For example, the fifth smallest number.
         *
         * @param array Is an array or range of numerical data for which you want to determine the k-th smallest value.
         * @param k Is the position (from the smallest) in the array or range of the value to return.
         *
         * [Api set: ExcelApi 1.2]
         */
        small(array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, k: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the square root of a number.
         *
         * @param number Is the number for which you want the square root.
         *
         * [Api set: ExcelApi 1.2]
         */
        sqrt(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the square root of (number * Pi).
         *
         * @param number Is the number by which p is multiplied.
         *
         * [Api set: ExcelApi 1.2]
         */
        sqrtPi(number: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Estimates standard deviation based on a sample, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1.
         *
         * @param values List of parameters, whose elements are 1 to 255 values corresponding to a sample of a population and can be values or names or references to values.
         *
         * [Api set: ExcelApi 1.2]
         */
        stDevA(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Calculates standard deviation based on an entire population, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1.
         *
         * @param values List of parameters, whose elements are 1 to 255 values corresponding to a population and can be values, names, arrays, or references that contain values.
         *
         * [Api set: ExcelApi 1.2]
         */
        stDevPA(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Calculates standard deviation based on the entire population given as arguments (ignores logical values and text).
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers corresponding to a population and can be numbers or references that contain numbers.
         *
         * [Api set: ExcelApi 1.2]
         */
        stDev_P(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Estimates standard deviation based on a sample (ignores logical values and text in the sample).
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers corresponding to a sample of a population and can be numbers or references that contain numbers.
         *
         * [Api set: ExcelApi 1.2]
         */
        stDev_S(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns a normalized value from a distribution characterized by a mean and standard deviation.
         *
         * @param x Is the value you want to normalize.
         * @param mean Is the arithmetic mean of the distribution.
         * @param standardDev Is the standard deviation of the distribution, a positive number.
         *
         * [Api set: ExcelApi 1.2]
         */
        standardize(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, mean: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, standardDev: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Replaces existing text with new text in a text string.
         *
         * @param text Is the text or the reference to a cell containing text in which you want to substitute characters.
         * @param oldText Is the existing text you want to replace. If the case of oldText does not match the case of text, SUBSTITUTE will not replace the text.
         * @param newText Is the text you want to replace oldText with.
         * @param instanceNum Specifies which occurrence of oldText you want to replace. If omitted, every instance of oldText is replaced.
         *
         * [Api set: ExcelApi 1.2]
         */
        substitute(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, oldText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, newText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, instanceNum?: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns a subtotal in a list or database.
         *
         * @param functionNum Is the number 1 to 11 that specifies the summary function for the subtotal.
         * @param values List of parameters, whose elements are 1 to 254 ranges or references for which you want the subtotal.
         *
         * [Api set: ExcelApi 1.2]
         */
        subtotal(functionNum: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ...values: Array<Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Adds all the numbers in a range of cells.
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers to sum. Logical values and text are ignored in cells, included if typed as arguments.
         *
         * [Api set: ExcelApi 1.2]
         */
        sum(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Adds the cells specified by a given condition or criteria.
         *
         * @param range Is the range of cells you want evaluated.
         * @param criteria Is the condition or criteria in the form of a number, expression, or text that defines which cells will be added.
         * @param sumRange Are the actual cells to sum. If omitted, the cells in range are used.
         *
         * [Api set: ExcelApi 1.2]
         */
        sumIf(range: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, criteria: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, sumRange?: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Adds the cells specified by a given set of conditions or criteria.
         *
         * @param sumRange Are the actual cells to sum.
         * @param values List of parameters, where the first element of each pair is the Is the range of cells you want evaluated for the particular condition , and the second element is is the condition or criteria in the form of a number, expression, or text that defines which cells will be added.
         *
         * [Api set: ExcelApi 1.2]
         */
        sumIfs(sumRange: Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, ...values: Array<Excel.Range | Excel.RangeReference | Excel.FunctionResult<any> | number | string | boolean>): FunctionResult<number>;
        /**
         *
         * Returns the sum of the squares of the arguments. The arguments can be numbers, arrays, names, or references to cells that contain numbers.
         *
         * @param values List of parameters, whose elements are 1 to 255 numbers, arrays, names, or references to arrays for which you want the sum of the squares.
         *
         * [Api set: ExcelApi 1.2]
         */
        sumSq(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the sum-of-years' digits depreciation of an asset for a specified period.
         *
         * @param cost Is the initial cost of the asset.
         * @param salvage Is the salvage value at the end of the life of the asset.
         * @param life Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
         * @param per Is the period and must use the same units as Life.
         *
         * [Api set: ExcelApi 1.2]
         */
        syd(cost: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvage: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, life: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, per: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Checks whether a value is text, and returns the text if it is, or returns double quotes (empty text) if it is not.
         *
         * @param value Is the value to test.
         *
         * [Api set: ExcelApi 1.2]
         */
        t(value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the bond-equivalent yield for a treasury bill.
         *
         * @param settlement Is the Treasury bill's settlement date, expressed as a serial date number.
         * @param maturity Is the Treasury bill's maturity date, expressed as a serial date number.
         * @param discount Is the Treasury bill's discount rate.
         *
         * [Api set: ExcelApi 1.2]
         */
        tbillEq(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, discount: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the price per $100 face value for a treasury bill.
         *
         * @param settlement Is the Treasury bill's settlement date, expressed as a serial date number.
         * @param maturity Is the Treasury bill's maturity date, expressed as a serial date number.
         * @param discount Is the Treasury bill's discount rate.
         *
         * [Api set: ExcelApi 1.2]
         */
        tbillPrice(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, discount: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the yield for a treasury bill.
         *
         * @param settlement Is the Treasury bill's settlement date, expressed as a serial date number.
         * @param maturity Is the Treasury bill's maturity date, expressed as a serial date number.
         * @param pr Is the Treasury Bill's price per $100 face value.
         *
         * [Api set: ExcelApi 1.2]
         */
        tbillYield(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pr: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the left-tailed Student's t-distribution.
         *
         * @param x Is the numeric value at which to evaluate the distribution.
         * @param degFreedom Is an integer indicating the number of degrees of freedom that characterize the distribution.
         * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        t_Dist(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the two-tailed Student's t-distribution.
         *
         * @param x Is the numeric value at which to evaluate the distribution.
         * @param degFreedom Is an integer indicating the number of degrees of freedom that characterize the distribution.
         *
         * [Api set: ExcelApi 1.2]
         */
        t_Dist_2T(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the right-tailed Student's t-distribution.
         *
         * @param x Is the numeric value at which to evaluate the distribution.
         * @param degFreedom Is an integer indicating the number of degrees of freedom that characterize the distribution.
         *
         * [Api set: ExcelApi 1.2]
         */
        t_Dist_RT(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the left-tailed inverse of the Student's t-distribution.
         *
         * @param probability Is the probability associated with the two-tailed Student's t-distribution, a number between 0 and 1 inclusive.
         * @param degFreedom Is a positive integer indicating the number of degrees of freedom to characterize the distribution.
         *
         * [Api set: ExcelApi 1.2]
         */
        t_Inv(probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the two-tailed inverse of the Student's t-distribution.
         *
         * @param probability Is the probability associated with the two-tailed Student's t-distribution, a number between 0 and 1 inclusive.
         * @param degFreedom Is a positive integer indicating the number of degrees of freedom to characterize the distribution.
         *
         * [Api set: ExcelApi 1.2]
         */
        t_Inv_2T(probability: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, degFreedom: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the tangent of an angle.
         *
         * @param number Is the angle in radians for which you want the tangent. Degrees * PI()/180 = radians.
         *
         * [Api set: ExcelApi 1.2]
         */
        tan(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the hyperbolic tangent of a number.
         *
         * @param number Is any real number.
         *
         * [Api set: ExcelApi 1.2]
         */
        tanh(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a value to text in a specific number format.
         *
         * @param value Is a number, a formula that evaluates to a numeric value, or a reference to a cell containing a numeric value.
         * @param formatText Is a number format in text form from the Category box on the Number tab in the Format Cells dialog box (not General).
         *
         * [Api set: ExcelApi 1.2]
         */
        text(value: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, formatText: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Converts hours, minutes, and seconds given as numbers to an Excel serial number, formatted with a time format.
         *
         * @param hour Is a number from 0 to 23 representing the hour.
         * @param minute Is a number from 0 to 59 representing the minute.
         * @param second Is a number from 0 to 59 representing the second.
         *
         * [Api set: ExcelApi 1.2]
         */
        time(hour: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, minute: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, second: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a text time to an Excel serial number for a time, a number from 0 (12:00:00 AM) to 0.999988426 (11:59:59 PM). Format the number with a time format after entering the formula.
         *
         * @param timeText Is a text string that gives a time in any one of the Microsoft Excel time formats (date information in the string is ignored).
         *
         * [Api set: ExcelApi 1.2]
         */
        timevalue(timeText: string | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the current date formatted as a date.
         *
         * [Api set: ExcelApi 1.2]
         */
        today(): FunctionResult<number>;
        /**
         *
         * Removes all spaces from a text string except for single spaces between words.
         *
         * @param text Is the text from which you want spaces removed.
         *
         * [Api set: ExcelApi 1.2]
         */
        trim(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the mean of the interior portion of a set of data values.
         *
         * @param array Is the range or array of values to trim and average.
         * @param percent Is the fractional number of data points to exclude from the top and bottom of the data set.
         *
         * [Api set: ExcelApi 1.2]
         */
        trimMean(array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, percent: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the logical value TRUE.
         *
         * [Api set: ExcelApi 1.2]
         */
        true(): FunctionResult<boolean>;
        /**
         *
         * Truncates a number to an integer by removing the decimal, or fractional, part of the number.
         *
         * @param number Is the number you want to truncate.
         * @param numDigits Is a number specifying the precision of the truncation, 0 (zero) if omitted.
         *
         * [Api set: ExcelApi 1.2]
         */
        trunc(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, numDigits?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns an integer representing the data type of a value: number = 1; text = 2; logical value = 4; error value = 16; array = 64.
         *
         * @param value Can be any value.
         *
         * [Api set: ExcelApi 1.2]
         */
        type(value: boolean | string | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a number to text, using currency format.
         *
         * @param number Is a number, a reference to a cell containing a number, or a formula that evaluates to a number.
         * @param decimals Is the number of digits to the right of the decimal point.
         *
         * [Api set: ExcelApi 1.2]
         */
        usdollar(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, decimals?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the Unicode character referenced by the given numeric value.
         *
         * @param number Is the Unicode number representing a character.
         *
         * [Api set: ExcelApi 1.2]
         */
        unichar(number: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Returns the number (code point) corresponding to the first character of the text.
         *
         * @param text Is the character that you want the Unicode value of.
         *
         * [Api set: ExcelApi 1.2]
         */
        unicode(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Converts a text string to all uppercase letters.
         *
         * @param text Is the text you want converted to uppercase, a reference or a text string.
         *
         * [Api set: ExcelApi 1.2]
         */
        upper(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<string>;
        /**
         *
         * Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify. By default, the table must be sorted in an ascending order.
         *
         * @param lookupValue Is the value to be found in the first column of the table, and can be a value, a reference, or a text string.
         * @param tableArray Is a table of text, numbers, or logical values, in which data is retrieved. tableArray can be a reference to a range or a range name.
         * @param colIndexNum Is the column number in tableArray from which the matching value should be returned. The first column of values in the table is column 1.
         * @param rangeLookup Is a logical value: to find the closest match in the first column (sorted in ascending order) = TRUE or omitted; find an exact match = FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        vlookup(lookupValue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, tableArray: Excel.Range | number | Excel.RangeReference | Excel.FunctionResult<any>, colIndexNum: Excel.Range | number | Excel.RangeReference | Excel.FunctionResult<any>, rangeLookup?: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number | string | boolean>;
        /**
         *
         * Converts a text string that represents a number to a number.
         *
         * @param text Is the text enclosed in quotation marks or a reference to a cell containing the text you want to convert.
         *
         * [Api set: ExcelApi 1.2]
         */
        value(text: string | boolean | number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Estimates variance based on a sample, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1.
         *
         * @param values List of parameters, whose elements are 1 to 255 value arguments corresponding to a sample of a population.
         *
         * [Api set: ExcelApi 1.2]
         */
        varA(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Calculates variance based on the entire population, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1.
         *
         * @param values List of parameters, whose elements are 1 to 255 value arguments corresponding to a population.
         *
         * [Api set: ExcelApi 1.2]
         */
        varPA(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Calculates variance based on the entire population (ignores logical values and text in the population).
         *
         * @param values List of parameters, whose elements are 1 to 255 numeric arguments corresponding to a population.
         *
         * [Api set: ExcelApi 1.2]
         */
        var_P(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Estimates variance based on a sample (ignores logical values and text in the sample).
         *
         * @param values List of parameters, whose elements are 1 to 255 numeric arguments corresponding to a sample of a population.
         *
         * [Api set: ExcelApi 1.2]
         */
        var_S(...values: Array<number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<number>;
        /**
         *
         * Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method or some other method you specify.
         *
         * @param cost Is the initial cost of the asset.
         * @param salvage Is the salvage value at the end of the life of the asset.
         * @param life Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).
         * @param startPeriod Is the starting period for which you want to calculate the depreciation, in the same units as Life.
         * @param endPeriod Is the ending period for which you want to calculate the depreciation, in the same units as Life.
         * @param factor Is the rate at which the balance declines, 2 (double-declining balance) if omitted.
         * @param noSwitch Switch to straight-line depreciation when depreciation is greater than the declining balance = FALSE or omitted; do not switch = TRUE.
         *
         * [Api set: ExcelApi 1.2]
         */
        vdb(cost: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, salvage: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, life: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, startPeriod: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, endPeriod: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, factor?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, noSwitch?: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the week number in the year.
         *
         * @param serialNumber Is the date-time code used by Microsoft Excel for date and time calculation.
         * @param returnType Is a number (1 or 2) that determines the type of the return value.
         *
         * [Api set: ExcelApi 1.2]
         */
        weekNum(serialNumber: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, returnType?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns a number from 1 to 7 identifying the day of the week of a date.
         *
         * @param serialNumber Is a number that represents a date.
         * @param returnType Is a number: for Sunday=1 through Saturday=7, use 1; for Monday=1 through Sunday=7, use 2; for Monday=0 through Sunday=6, use 3.
         *
         * [Api set: ExcelApi 1.2]
         */
        weekday(serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, returnType?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the Weibull distribution.
         *
         * @param x Is the value at which to evaluate the function, a nonnegative number.
         * @param alpha Is a parameter to the distribution, a positive number.
         * @param beta Is a parameter to the distribution, a positive number.
         * @param cumulative Is a logical value: for the cumulative distribution function, use TRUE; for the probability mass function, use FALSE.
         *
         * [Api set: ExcelApi 1.2]
         */
        weibull_Dist(x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, alpha: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, beta: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, cumulative: boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the serial number of the date before or after a specified number of workdays.
         *
         * @param startDate Is a serial date number that represents the start date.
         * @param days Is the number of nonweekend and non-holiday days before or after startDate.
         * @param holidays Is an optional array of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays.
         *
         * [Api set: ExcelApi 1.2]
         */
        workDay(startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, days: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, holidays?: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the serial number of the date before or after a specified number of workdays with custom weekend parameters.
         *
         * @param startDate Is a serial date number that represents the start date.
         * @param days Is the number of nonweekend and non-holiday days before or after startDate.
         * @param weekend Is a number or string specifying when weekends occur.
         * @param holidays Is an optional array of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays.
         *
         * [Api set: ExcelApi 1.2]
         */
        workDay_Intl(startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, days: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, weekend?: number | string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, holidays?: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the internal rate of return for a schedule of cash flows.
         *
         * @param values Is a series of cash flows that correspond to a schedule of payments in dates.
         * @param dates Is a schedule of payment dates that corresponds to the cash flow payments.
         * @param guess Is a number that you guess is close to the result of XIRR.
         *
         * [Api set: ExcelApi 1.2]
         */
        xirr(values: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>, dates: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>, guess?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the net present value for a schedule of cash flows.
         *
         * @param rate Is the discount rate to apply to the cash flows.
         * @param values Is a series of cash flows that correspond to a schedule of payments in dates.
         * @param dates Is a schedule of payment dates that corresponds to the cash flow payments.
         *
         * [Api set: ExcelApi 1.2]
         */
        xnpv(rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, values: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>, dates: number | string | Excel.Range | boolean | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns a logical 'Exclusive Or' of all arguments.
         *
         * @param values List of parameters, whose elements are 1 to 254 conditions you want to test that can be either TRUE or FALSE and can be logical values, arrays, or references.
         *
         * [Api set: ExcelApi 1.2]
         */
        xor(...values: Array<boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>>): FunctionResult<boolean>;
        /**
         *
         * Returns the year of a date, an integer in the range 1900 - 9999.
         *
         * @param serialNumber Is a number in the date-time code used by Microsoft Excel.
         *
         * [Api set: ExcelApi 1.2]
         */
        year(serialNumber: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the year fraction representing the number of whole days between start_date and end_date.
         *
         * @param startDate Is a serial date number that represents the start date.
         * @param endDate Is a serial date number that represents the end date.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        yearFrac(startDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, endDate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the yield on a security that pays periodic interest.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param rate Is the security's annual coupon rate.
         * @param pr Is the security's price per $100 face value.
         * @param redemption Is the security's redemption value per $100 face value.
         * @param frequency Is the number of coupon payments per year.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        yield(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pr: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemption: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, frequency: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the annual yield for a discounted security. For example, a treasury bill.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param pr Is the security's price per $100 face value.
         * @param redemption Is the security's redemption value per $100 face value.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        yieldDisc(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pr: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, redemption: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the annual yield of a security that pays interest at maturity.
         *
         * @param settlement Is the security's settlement date, expressed as a serial date number.
         * @param maturity Is the security's maturity date, expressed as a serial date number.
         * @param issue Is the security's issue date, expressed as a serial date number.
         * @param rate Is the security's interest rate at date of issue.
         * @param pr Is the security's price per $100 face value.
         * @param basis Is the type of day count basis to use.
         *
         * [Api set: ExcelApi 1.2]
         */
        yieldMat(settlement: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, maturity: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, issue: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, rate: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, pr: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, basis?: number | string | boolean | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /**
         *
         * Returns the one-tailed P-value of a z-test.
         *
         * @param array Is the array or range of data against which to test X.
         * @param x Is the value to test.
         * @param sigma Is the population (known) standard deviation. If omitted, the sample standard deviation is used.
         *
         * [Api set: ExcelApi 1.2]
         */
        z_Test(array: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, x: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>, sigma?: number | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
    }
    module ErrorCodes {
        var accessDenied: string;
        var apiNotFound: string;
        var generalException: string;
        var insertDeleteConflict: string;
        var invalidArgument: string;
        var invalidBinding: string;
        var invalidOperation: string;
        var invalidReference: string;
        var invalidSelection: string;
        var itemAlreadyExists: string;
        var itemNotFound: string;
        var notImplemented: string;
        var unsupportedOperation: string;
    }
}
declare module Excel {
    /**
     * The RequestContext object facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the request context is required to get access to the Excel object model from the add-in.
     */
    class RequestContext extends OfficeExtension.ClientRequestContext {
        private m_workbook;
        constructor(url?: string);
        workbook: Workbook;
    }
    /**
     * Executes a batch script that performs actions on the Excel object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
     */
    function run<T>(batch: (context: Excel.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the Excel object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param object - A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
     */
    function run<T>(object: OfficeExtension.ClientObject, batch: (context: Excel.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the Excel object model, using the RequestContext of previously-created API objects.
     * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
     */
    function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: Excel.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
}
declare module Excel {
    var _RedirectV1APIs: boolean;
    var _V1APIMap: {
        "GetDataAsync": {
            call: (ctx: any, callArgs: any) => any;
            postprocess: (response: any, callArgs: any) => any;
        };
        "GetSelectedDataAsync": {
            call: (ctx: any, callArgs: any) => any;
            postprocess: (response: any, callArgs: any) => any;
        };
        "GoToByIdAsync": {
            call: (ctx: any, callArgs: any) => any;
        };
        "AddColumnsAsync": {
            call: (ctx: any, callArgs: any) => any;
        };
        "AddFromSelectionAsync": {
            call: (ctx: any, callArgs: any) => any;
            postprocess: (response: any) => any;
        };
        "AddFromNamedItemAsync": {
            call: (ctx: any, callArgs: any) => any;
            postprocess: (response: any) => any;
        };
        "AddFromPromptAsync": {
            call: (ctx: any, callArgs: any) => any;
            postprocess: (response: any) => any;
        };
        "AddRowsAsync": {
            call: (ctx: any, callArgs: any) => any;
        };
        "GetByIdAsync": {
            call: (ctx: any, callArgs: any) => any;
            postprocess: (response: any) => any;
        };
        "ReleaseByIdAsync": {
            call: (ctx: any, callArgs: any) => any;
        };
        "GetAllAsync": {
            call: (ctx: any) => any;
            postprocess: (response: any) => any;
        };
        "DeleteAllDataValuesAsync": {
            call: (ctx: any, callArgs: any) => any;
        };
        "SetSelectedDataAsync": {
            preprocess: (callArgs: any) => any;
            call: (ctx: any, callArgs: any) => any;
        };
        "SetDataAsync": {
            preprocess: (callArgs: any) => any;
            call: (ctx: any, callArgs: any) => any;
        };
        "SetFormatsAsync": {
            preprocess: (callArgs: any) => any;
            call: (ctx: any, callArgs: any) => any;
        };
        "SetTableOptionsAsync": {
            call: (ctx: any, callArgs: any) => any;
        };
        "ClearFormatsAsync": {
            call: (ctx: any, callArgs: any) => any;
        };
    };
}
