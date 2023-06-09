##  How to manipulation JSON in VB scrips

### Functions:
``` vbscript
'create script control for javascript
Private Function GetScriptControlForJs()
    Dim ScriptControl
    Set ScriptControl = CreateObject("MSScriptControl.ScriptControl")
    ScriptControl.Language = "JScript"
    ScriptControl.AddCode ("function getByIndex(jsonObj, index) { return jsonObj[index]; } ")
    ScriptControl.AddCode ("function convertToMap(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } ")
    Set GetScriptControlForJs = ScriptControl
End Function

'convert JSON string to JSON object
Public Function ConvertToJSONObject(jsonStr)
    Set engine = GetScriptControlForJs()
    Set ConvertToJSONObject = engine.eval("(" & jsonStr & ")")
End Function

'gets by index
Function GetByIndex(jsonVariant, Index)
    Set engine = GetScriptControlForJs()
    Set GetByIndex = engine.Run("getByIndex", jsonVariant, Index)
End Function

'gets by property
Public Function GetByProperty(jsonVariant, propertyName)
    Set engine = GetScriptControlForJs()
    GetByProperty = engine.Run("getByIndex", jsonVariant, propertyName)
End Function

'convert to Map
Public Function ConvertToMap(jsonVariant)
    Dim KeysObject
    Set engine = GetScriptControlForJs()
    Set ConvertToMap = engine.Run("convertToMap", jsonVariant)
End Function

```
### Examples:

##### data source:

```json
{
  "SquadName" : "Super hero squad",
  "HomeTown" : "Metro City",
  "Formed" : 2016,
  "SecretBase" : "Super tower",
  "Active" : true,
  "Members" : [
    {
      "Name" : "Molecule Man",
      "Age" : 29,
      "SecretIdentity" : "Dan Jukes",
      "Powers" : [
        "Radiation resistance",
        "Turning tiny",
        "Radiation blast"
      ]
    },
    {
      "Name" : "Madame Uppercut",
      "Age" : 39,
      "SecretIdentity" : "Jane Wilson",
      "Powers" : [
        "Million tonne punch",
        "Damage resistance",
        "Superhuman reflexes"
      ]
    },
    {
      "Name" : "Eternal Flame",
      "Age" : 1000000,
      "SecretIdentity" : "Unknown",
      "Powers" : [
        "Immortality",
        "Heat Immunity",
        "Inferno",
        "Teleportation",
        "Interdimensional travel"
      ]
    }
  ]
}
```

##### member access:

``` vbscript
    'convert to JSON object
    Set JsonObject = ConvertToJSONObject(JsonData)
    'squadName
    Debug.Print "SquadName:" & JsonObject.SquadName
    'homeTown
    Debug.Print "HomeTown:" & JsonObject.HomeTown
    'formed
    Debug.Print "Formed:" & JsonObject.Formed
    'secretBase
    Debug.Print "SecretBase:" & JsonObject.SecretBase
    'active
    Debug.Print "Active:" & JsonObject.Active

    Debug.Print "......Done......"
    
```

##### output:

```
SquadName:Super hero squad
HomeTown:Metro City
Formed:2016
SecretBase:Super tower
Active:True
......Done......
```

##### get property:
```vbscript
    'convert to JSON object
    Set JsonObject = ConvertToJSONObject(JsonData)
    'Members
    Set Members = JsonObject.Members
    'get length of array
    lenOfArray = GetByProperty(Members, "length")    
    Debug.Print "Length of array:" & lenOfArray
```
##### output:
```
Length of array:3
```




##### loop JSON array:

```vbscript
    'convert to JSON object
    Set JsonObject = ConvertToJSONObject(JsonData)
    'Members
    Set Members = JsonObject.Members
    'get length of array
    lenOfArray = GetByProperty(Members, "length")
    'loop
    For i = 0 To CInt(lenOfArray) - 1
       Set mem = GetByIndex(Members, i)
       Debug.Print "Name:" & mem.Name
       Debug.Print "Age:" & mem.Age
       Debug.Print "SecretIdentity:" & mem.SecretIdentity
       Set Powers = mem.Powers
       lenOfPowers = GetByProperty(Powers, "length")
       Debug.Print "################Powers#################"
       For j = 0 To CInt(lenOfPowers) - 1
          Power = GetByProperty(Powers, j)
          Debug.Print Power
       Next
    Next
   Debug.Print "......Done......"
```
##### output:

```
Name:Molecule Man
Age:29
SecretIdentity:Dan Jukes
################Powers#################
Radiation resistance
Turning tiny
Radiation blast
Name:Madame Uppercut
Age:39
SecretIdentity:Jane Wilson
################Powers#################
Million tonne punch
Damage resistance
Superhuman reflexes
Name:Eternal Flame
Age:1000000
SecretIdentity:Unknown
################Powers#################
Immortality
Heat Immunity
Inferno
Teleportation
Interdimensional travel
......Done......
```

