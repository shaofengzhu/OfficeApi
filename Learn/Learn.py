import json
import enum
class ActionType(enum.IntEnum):
    Instantiate = 1
    Query = 2
    Method = 3
    SetProperty = 4
    Trace = 5

class ActionType2(enum.Enum):
    Instantiate = 1
    Query = 2
    Method = 3
    SetProperty = 4
    Trace = 5

class RequestInfo:
    Actions = None
    Objects = None

class ActionInfo:
    Id = 0
    Type = ActionType.Method
    Arguments = None
    ObjectPaths = None

req = RequestInfo()
req.Actions = []
a1 = ActionInfo()
a1.Type = ActionType.Instantiate
a1.Id = 20
req.Actions.append(a1)
a2 = ActionInfo()
a2.Type = ActionType.Method
a2.Id = 21
a2.Arguments = ["abc", 345]
req.Actions.append(a2)

s = json.dumps(req, default = lambda o: o.__dict__)
print(s)

j = json.loads(s)
if "Another" in j:
    print(j["Another"])
else:
    print("key not exists")
for (k, v) in j.items():
    print(k, v)


