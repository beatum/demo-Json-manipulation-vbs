'create response
Private Function CreateResponse(result, message, object)
    Dim myResponse
    Set myResponse = CreateObject("Scripting.Dictionary")
    myResponse.Add "result", result
    myResponse.Add "message", message
    myResponse.Add "object", object
    Set CreateResponse = myResponse
End Function

'get response
Public Function GetResponse(object)
    Set response = CreateResponse(True, "OK", object)
    Set GetResponse = response
End Function

'get error response
Public Function GetErrorResponse(message)
    Set response = CreateResponse(False, message, Empty)
    Set GetErrorResponse = response
End Function

'create script control for javascript
Private Function GetScriptControlForJs()
    On Error Resume Next
    Dim ScriptControl
    Set ScriptControl = CreateObject("MSScriptControl.ScriptControl")
    ScriptControl.Language = "JScript"
    ScriptControl.AddCode ("function getByIndex(jsonObj, index) { return jsonObj[index]; } ")
    ScriptControl.AddCode ("function convertToMap(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } ")
    Dim response
    If Err.Number <> 0 Then
        Set response = GetErrorResponse("Can't create MSScriptControl object")
        Err.Clear
    Else
        Set response = GetResponse(ScriptControl)
    End If
    Set GetScriptControlForJs = response
End Function

'convert JSON string to JSON object
Public Function ConvertToJSONObject(jsonStr)
    On Error Resume Next
    Dim response
    Set response = GetScriptControlForJs()
    If response.Item("result") = False Then
        Set ConvertToJSONObject = response
        Exit Function
    End If
    Set engine = response.Item("object")
    Set responseOfEval = engine.eval("(" & jsonStr & ")")
    If Err.Number <> 0 Then
        Set response = GetErrorResponse("Incorrect JSON string:" & Err.Description)
        Err.Clear
    Else
        Set response = GetResponse(responseOfEval)
    End If
    Set ConvertToJSONObject = response
End Function

'gets by index
Function GetByIndex(jsonVariant, Index)
    On Error Resume Next
    Dim response
    Set response = GetScriptControlForJs()
    If response.Item("result") = False Then
        Set ConvertToJSONObject = response
        Exit Function
    End If
    Set engine = response.Item("object")
    Set property = engine.Run("getByIndex", jsonVariant, Index)
    If Err.Number <> 0 Then
        Set response = GetErrorResponse("Get value by index error:" & Err.Description)
        Err.Clear
    Else
        Set response = GetResponse(property)
    End If
    Set GetByIndex = response
End Function

'gets by property
Public Function GetByProperty(jsonVariant, propertyName)
    On Error Resume Next
    Dim response
    Set response = GetScriptControlForJs()
    If response.Item("result") = False Then
        Set GetByProperty = response
        Exit Function
    End If
    Set engine = response.Item("object")
    property = engine.Run("getByIndex", jsonVariant, propertyName)
    If Err.Number <> 0 Then
        Set response = GetErrorResponse("Get value by property error:" & Err.Description)
        Err.Clear
    Else
        Set response = GetResponse(property)
    End If
    Set GetByProperty = response
End Function

'convert to Map
Public Function ConvertToMap(jsonVariant)
    Dim KeysObject
    Set engine = GetScriptControlForJs()
    Set ConvertToMap = engine.Run("convertToMap", jsonVariant)
End Function
