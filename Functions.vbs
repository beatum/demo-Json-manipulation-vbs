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