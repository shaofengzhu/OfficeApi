# Use Office.js API from standalone web application

## Introduction
Office.js API allows us to access Office application, for example Excel, from 
an Office Add-In. For example, to update range value in the workbook, we could
use code

```js
function excelHelloWorld() {
	var ctx = new Excel.RequestContext();
	var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:B2");
	range.values = [["Hello", "World"], ["12345", "=A2 + 100"]];
	ctx.load(range);
	ctx.sync()
		.then(function () {
			console.log("Success");
			console.log("Range.values=" + JSON.stringify(range.values));
		})
		.catch(function (err) {
			console.error(JSON.stringify(err));
		});
}

```

For Excel, the same set of API is also exposed to Microsoft Graph and developer 
could use those API through Microsoft Graph with REST syntax. For example, the 
following request will return all of the worksheets in the workbook.

```
GET https://graph.microsoft.com/v1.0/me/drive/items/01SOTJQBSMQAAV74NVARCZHKIBXR3B43XL/workbook/worksheets
```

In the above URL, the part 
`https://graph.microsoft.com/v1.0/me/drive/items/01SOTJQBSMQAAV74NVARCZHKIBXR3B43XL` 
is the file item URL. The part `workbook` indicates to access the Excel 
workbook related resources. The part `worksheets` is one of the resource in 
the Excel workbook.

To update range value in the workbook, the request is

```
PATCH https://graph.microsoft.com/v1.0/me/drive/items/01SOTJQBSMQAAV74NVARCZHKIBXR3B43XL/workbook/worksheets('sheet1')/range(address='A1:B2')

{
    values: [["Hello", "World"], ["12345", "=A2 + 100"]]
}
```

Microsoft Graph supports implicit OAuth flow and a standalone web application
could use CORS to send REST request to access resource in an Excel workbook 
stored in Microsoft OneDrive.

In this article, we want to show you how to use the Office.js API from your 
standalone web application to access resource in an Excel workbook.

## Hello, World from Office Add-in
Within an Excel Add-in, the following code could be used to update a range
value

```js
function excelHelloWorld() {
	var ctx = new Excel.RequestContext();
	var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:B2");
	range.values = [["Hello", "World"], ["12345", "=A2 + 100"]];
	ctx.load(range);
	ctx.sync()
		.then(function () {
			var elem = document.createElement("div");
			elem.innerText = JSON.stringify(range.values);
			document.body.appendChild(elem);
		})
		.catch(function (err) {
			console.error(JSON.stringify(err));
		});
}
```

The above code will get the range `A1:B2` in Worksheet `Sheet1` and then update
the value in range `A1:B2` 
1. Set `A1` with `Hello`
2. Set `B1` with `World`
3. Set `A2` with `12345`
4. Set `B2` with `A2` value plus `100`


## Hello, World from standalone web application
We will run the same code as Office Add-in in the standalone web application.

After we use Azure management portal to register an OAuth implicit flow app, 
we could write a very simple HTML page to demonstrate using OAuth implicit flow to 
request OAuth access token, create a workbook session and then trigger the same code as the
Office Add-in

Let's start with a very simple HTML page that only reference `Office.Runtime.js` and `Excel.js`

```html
<!DOCTYPE html>
<html>
<head>
	<meta charset="UTF-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=Edge" />
	<meta http-equiv="Expires" content="0" />
	<title>OAuth Implicit Flow App</title>
	<script src="../Office.Runtime.js" type="text/javascript"></script>
	<script src="../Excel.js" type="text/javascript"></script>
</head>
<body
</body>
</html>
```

### Step 1
The first step is to initialize oAuth flow. Once the OAuth flow finished, the body onload
event handler will extract the oAuth token from URL.

```js
function initOAuthImplicitFlow() {
	var currentPageUrl = document.URL;
	var index = currentPageUrl.indexOf("#");
	if (index > 0) {
		currentPageUrl = currentPageUrl.substr(0, index);
	}

	var appId = "6cd9319f-2e1d-4664-bd34-857c0fe8b8eb";
	var url = "https://login.windows.net/common/oauth2/authorize?response_type=token&client_id="
		+ appId
		+ "&resource=https://graph.microsoft.com&redirect_uri="
		+ encodeURIComponent(currentPageUrl);
	window.location.assign(url);
}

function onBodyLoaded() {
	var hash = window.location.hash.substr(1); // remove #
	if (hash) {
		var pairs = hash.split('&');
		var keyValues = {};
		// If there are parameters in URL, extract key/value pairs. 
		for (var i = 0; i < pairs.length; ++i) {
			var p = pairs[i].split('=', 2);
			if (p.length == 1)
				keyValues[p[0]] = "";
			else
				keyValues[p[0]] = decodeURIComponent(p[1].replace(/\+/g, " "));
		}
		accessToken = keyValues["access_token"];
		if (accessToken) {
			document.getElementById("TxtAccessToken").value = accessToken;
		}
	}
}

```

### Step 2
The second step is to initialize Excel workbook session id.

```js
function sendHttpRequest(request){
	return new OfficeExtension.Promise(function(resolve, reject) {
		var xhr = new XMLHttpRequest();
		xhr.open(request.method, request.url);
		xhr.onload = function () {
			var resp = {
				statusCode: xhr.status,
				body: xhr.responseText
			};

			resolve(resp);
		};

		xhr.onerror = function () {
			reject("Error " + xhr.statusText);
		};

		if (request.headers) {
			for (var key in request.headers) {
				xhr.setRequestHeader(key, request.headers[key]);
			}
		}

		xhr.send(request.body);
	});
}

function getWorkbookUrl() {
	var agsUrlPrefix = "https://graph.microsoft.com/v1.0";
	var agsFileName = document.getElementById("TxtExcelFileName").value;
	var workbookUrl = agsUrlPrefix + "/me/drive/root:/" + agsFileName + ":/workbook";
	return workbookUrl;
}

function initWorkbookSession() {

	var accessToken = document.getElementById("TxtAccessToken").value;
	var requestHeaders = { Authorization: "Bearer " + accessToken };
	// create a session
	sendHttpRequest(
		{
			url: getWorkbookUrl() + "/createSession",
			method: "POST",
			body: JSON.stringify({ persistChanges: true }),
			headers: requestHeaders
		})
		.then(function (resp) {
			if (resp.statusCode !== 201) {
				throw "Invalid response:" + JSON.stringify(resp);
			}

			var session = JSON.parse(resp.body);
			document.getElementById("TxtWorkbookSessionId").value = session.id;

			OfficeExtension.ClientRequestContext.defaultRequestUrlAndHeaders =
				{
					url: getWorkbookUrl(),
					headers: {
						"Authorization": "Bearer " + document.getElementById("TxtAccessToken").value,
						"Workbook-Session-Id": session.id
					}
				};
		})
		.catch(function (ex) {
			alert(JSON.stringify(ex));
		});
}

```

In the above Javascript code, we created a utility method to send XMLHttpRequest 
using Promise pattern. We then send

```
POST https://graph.microsoft.com/v1.0/me/drive/root:/AgaveTest.xlsx:/workbook/createSession

{"persistChanges":true} 
```

to create a workbook session. Once we have all of the workbook information, we could set the
`OfficeExtension.ClientRequestContext.defaultRequestUrlAndHeaders`. By this way, when
we just call `new Excel.RequestContext()`, the code will pick up the default request url and
headers. Alternatively, we could use the following code to initialize the request context object.

```js
var ctx = new Excel.RequestContext(workbookUrl);
ctx.headers["Authorization"] = "Bearer " + accessToken;
ctx.headers["Workbook-Session-Id"] = sessionId;
```

### Step 3
After the above two steps, we could trigger the same code as Office Add-in

```js
function excelHelloWorld() {
	var ctx = new Excel.RequestContext();
	var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:B2");
	range.values = [["Hello", "World"], ["12345", "=A2 + 100"]];
	ctx.load(range);
	ctx.sync()
		.then(function () {
			var elem = document.createElement("div");
			elem.innerText = JSON.stringify(range.values);
			document.body.appendChild(elem);
		})
		.catch(function (err) {
			console.error(JSON.stringify(err));
		});
}
```

### Put them together
Let's put them together and the the whole page's HTML is

```html
<!DOCTYPE html>
<html>
<head>
	<meta charset="UTF-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=Edge" />
	<meta http-equiv="Expires" content="0" />
	<title>OAuth Implicit Flow App</title>
	<script src="../Office.Runtime.js" type="text/javascript"></script>
	<script src="../Excel.js" type="text/javascript"></script>
	<script type="text/javascript">
		function initOAuthImplicitFlow() {
			var currentPageUrl = document.URL;
			var index = currentPageUrl.indexOf("#");
			if (index > 0) {
				currentPageUrl = currentPageUrl.substr(0, index);
			}

			var appId = "6cd9319f-2e1d-4664-bd34-857c0fe8b8eb";
			var url = "https://login.windows.net/common/oauth2/authorize?response_type=token&client_id="
				+ appId
				+ "&resource=https://graph.microsoft.com&redirect_uri="
				+ encodeURIComponent(currentPageUrl);
			window.location.assign(url);
		}

		function onBodyLoaded() {
			var hash = window.location.hash.substr(1); // remove #
			if (hash) {
				var pairs = hash.split('&');
				var keyValues = {};
				// If there are parameters in URL, extract key/value pairs. 
				for (var i = 0; i < pairs.length; ++i) {
					var p = pairs[i].split('=', 2);
					if (p.length == 1)
						keyValues[p[0]] = "";
					else
						keyValues[p[0]] = decodeURIComponent(p[1].replace(/\+/g, " "));
				}
				accessToken = keyValues["access_token"];
				if (accessToken) {
					document.getElementById("TxtAccessToken").value = accessToken;
				}
			}
		}

		function sendHttpRequest(request){
			return new OfficeExtension.Promise(function(resolve, reject) {
				var xhr = new XMLHttpRequest();
				xhr.open(request.method, request.url);
				xhr.onload = function () {
					var resp = {
						statusCode: xhr.status,
						body: xhr.responseText
					};

					resolve(resp);
				};

				xhr.onerror = function () {
					reject("Error " + xhr.statusText);
				};

				if (request.headers) {
					for (var key in request.headers) {
						xhr.setRequestHeader(key, request.headers[key]);
					}
				}

				xhr.send(request.body);
			});
		}

		function getWorkbookUrl() {
			var agsUrlPrefix = "https://graph.microsoft.com/v1.0";
			var agsFileName = document.getElementById("TxtExcelFileName").value;
			var workbookUrl = agsUrlPrefix + "/me/drive/root:/" + agsFileName + ":/workbook";
			return workbookUrl;
		}

		function initWorkbookSession() {

			var accessToken = document.getElementById("TxtAccessToken").value;
			var requestHeaders = { Authorization: "Bearer " + accessToken };
			// create a session
			sendHttpRequest(
				{
					url: getWorkbookUrl() + "/createSession",
					method: "POST",
					body: JSON.stringify({ persistChanges: true }),
					headers: requestHeaders
				})
				.then(function (resp) {
					if (resp.statusCode !== 201) {
						throw "Invalid response:" + JSON.stringify(resp);
					}

					var session = JSON.parse(resp.body);
					document.getElementById("TxtWorkbookSessionId").value = session.id;

					OfficeExtension.ClientRequestContext.defaultRequestUrlAndHeaders =
						{
							url: getWorkbookUrl(),
							headers: {
								"Authorization": "Bearer " + document.getElementById("TxtAccessToken").value,
								"Workbook-Session-Id": session.id
							}
						};
				})
				.catch(function (ex) {
					alert(JSON.stringify(ex));
				});
		}

		function excelHelloWorld() {
			var ctx = new Excel.RequestContext();
			var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:B2");
			range.values = [["Hello", "World"], ["12345", "=A2 + 100"]];
			ctx.load(range);
			ctx.sync()
				.then(function () {
					var elem = document.createElement("div");
					elem.innerText = JSON.stringify(range.values);
					document.body.appendChild(elem);
				})
				.catch(function (err) {
					console.error(JSON.stringify(err));
				});
		}
	</script>
</head>
<body onload="onBodyLoaded()">
	<div>
		Step 1: <button onclick="initOAuthImplicitFlow()">Init OAuth flow and request access token</button>
	</div>
	<div>
		Access Token: <input type="text" id="TxtAccessToken" size="128" />
	</div>
	<div>
		Step 2: <button onclick="initWorkbookSession()">Init Workbook Session</button> for Excel file <input type="text" id="TxtExcelFileName" value="AgaveTest.xlsx" />.
	</div>
	<div>
		Workbook Session Id: <input type="text" id="TxtWorkbookSessionId" size="128" />
	</div>
	<div>
		Step 3: <button onclick="excelHelloWorld()">Hello, World</button>
	</div>
</body>
</html>

```

## Summary
As we could see from the above example, the only additional logic needed by the standalone 
web appliction is
1. Initialize OAuth flow to request OAuth access token
2. Create workbook session

Other than those, the standalone web application could share the same code as Office Add-in 
to access Excel APIs.

