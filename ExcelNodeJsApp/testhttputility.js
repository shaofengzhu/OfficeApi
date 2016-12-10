var OfficeExtension = require('@microsoft/office-js/office.runtime');
var Tests = {};
Tests.test1 = function(){
    return OfficeExtension.HttpUtility.sendRequest({
        url: "http://localhost:7010/FakeResponse.ashx"
    }).then(function(resp){
        console.log(JSON.stringify(resp));        
    });
};

Tests.test2 = function(){
    return OfficeExtension.HttpUtility.sendRequest({
        url: "http://localhost:7010/FakeResponse.ashx?code=204"
    }).then(function(resp){
        console.log(JSON.stringify(resp));        
    });
};

Tests.test3 = function(){
    return OfficeExtension.HttpUtility.sendRequest({
        method: "POST",
        url: "http://localhost:7010/FakeResponse.ashx?code=200",
        headers: {"content-type": "application/json"},
        body: JSON.stringify({street: "One Microsoft Way", city: "Redmond"})
    }).then(function(resp){
        console.log(JSON.stringify(resp));
    });
};

Tests.test404 = function(){
    return OfficeExtension.HttpUtility.sendRequest({
        method: "POST",
        url: "http://localhost:7010/FakeResponse.ashx?code=404",
        headers: {"content-type": "application/json"},
        body: JSON.stringify({street: "One Microsoft Way", city: "Redmond"})
    }).then(function(resp){
        console.log(JSON.stringify(resp));
    });
};

Tests.test500 = function(){
    return OfficeExtension.HttpUtility.sendRequest({
        method: "POST",
        url: "http://localhost:7010/FakeResponse.ashx?code=500",
        headers: {"content-type": "application/json"},
        body: JSON.stringify({street: "One Microsoft Way", city: "Redmond"})
    }).then(function(resp){
        console.log(JSON.stringify(resp));
    });
};

Tests.testBadAddress = function(){
    return OfficeExtension.HttpUtility.sendRequest({
        method: "POST",
        url: "http://www.non-existing.cont.com:7010/FakeResponse.ashx?code=200",
        headers: {"content-type": "application/json"},
        body: JSON.stringify({street: "One Microsoft Way", city: "Redmond"})
    }).then(function(resp){
        console.log(JSON.stringify(resp));
    });
};

function invokeTest(key, func){
    return function(){
        console.log("---" + key + "---");
        return func();
    }
}

var p = OfficeExtension.Utility._createPromiseFromResult(null);
for (var key in Tests){
    p = p.then(invokeTest(key, Tests[key]));
}
p = p.then(function(){
    console.log("---Done---");
}).catch(function(ex){
    console.log("---Error---");
    console.log(JSON.stringify(ex));
})
