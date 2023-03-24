'data to change if does not match with directory
Const FilePath = "C:\Users\ve00ym279\OneDrive - YAMAHA MOTOR CO., LTD\Desktop\Scripts\vbs\POC\test_one.json"
Const KeyToRead = "datatochange"
' Const RegexValue = "^[A-z]"
Const oldStringToReplace = "Hello"
Const newStringToBeReplaced = "Hey"

'Reading the JSON file
Dim FSO, file
Set FSO = CreateObject("Scripting.FileSystemObject")
Set file = FSO.OpenTextFile(FilePath,1)
Dim json
json = file.ReadAll
wscript.echo json
file.Close

'Get Specified Value of particular key
Dim data, value
WScript.Echo json & 1
Set data = getJsonValue(FilePath,KeyToRead)
WScript.Echo json & 2
value = data(KeyToRead)

Const EnvPath = "HKEY_CURRENT_USER\Environment\Test"
Dim wshShell
Set wshShell = CreateObject(WScript.Shell)
wshShell.RegWrite EnvPath & "\" & KeyToRead, value, "REG_SZ"

Function setRegistryKey(regPath,KeyToRead,valueToWrite)
    Dim wshShell
    Set wshShell = CreateObject(WScript.Shell)
    wshShell.RegWrite regPath & "\" & KeyToRead, value, "REG_SZ"
End Function

Function getRegistryKey(regPath)
    Set shellObject = CreateObject("WScript.Shell")
    getRegistryKey = shellObject.RegRead(regPath)
End Function
' json = Replace(json, oldStringToReplace, newStringToBeReplaced)

' Set file = FSO.OpenTextFile(FilePath, 2)
' file.Write json
' file.Close
' 'Replacing the value based on key and regex
' Dim regex, matches
' Set regex = New RegExp
' regex.Pattern = "(?:" & KeyToReplace & "\s*:\s*)""(" & RegexValue & ".*)"""
' regex.Global = True
' Set matches = regex.Execute(json)
' Dim match
' ' wscript.echo matches
' For Each match in matches
'     wscript.echo match & "HELLO"
' Next
' wscript.echo "HELLO"



Function getJsonValue(jsonPath, key)
    ' Load the JSON file
    Set jsonFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(jsonPath)
    json = jsonFile.ReadAll()
    jsonFile.Close()
    
    ' Parse the JSON data
    Set jsonDict = CreateObject("Scripting.Dictionary")
    Set jsonStack = CreateObject("Scripting.Dictionary")
    jsonStack.Add "", jsonDict
    
    i = 1
    While i <= Len(json)
        c = Mid(json, i, 1)
        If c = "{" Then
            Set newDict = CreateObject("Scripting.Dictionary")
            jsonStack.Item("").Add "", newDict
            jsonStack.Add "", newDict
        ElseIf c = "}" Then
            jsonStack.Remove ""
        ElseIf c = "[" Then
            Set newList = CreateObject("Scripting.Dictionary")
            jsonStack.Item("").Add "", newList
            jsonStack.Add "", newList
        ElseIf c = "]" Then
            jsonStack.Remove ""
        ElseIf c = "," Then
            ' Do nothing
        ElseIf c = ":" Then
            keyName = jsonStack.Item("").Item("currentKey")
            jsonStack.Item("").Remove "currentKey"
            jsonStack.Item("").Add keyName, ""
        ElseIf c = """" Then
            j = i + 1
            While Mid(json, j, 1) <> """"
                If Mid(json, j, 1) = "\" And Mid(json, j+1, 1) = """" Then
                    j = j + 1
                End If
                j = j + 1
            Wend
            keyName = Mid(json, i+1, j-i-1)
            jsonStack.Item("").Add "currentKey", keyName
            i = j
        ElseIf c = "t" Then
            jsonStack.Item("").Add jsonStack.Item("").Item("currentKey"), True
            jsonStack.Item("").Remove "currentKey"
            i = i + 3
        ElseIf c = "f" Then
            jsonStack.Item("").Add jsonStack.Item("").Item("currentKey"), False
            jsonStack.Item("").Remove "currentKey"
            i = i + 4
        ElseIf c = "n" Then
            jsonStack.Item("").Add jsonStack.Item("").Item("currentKey"), Null
            jsonStack.Item("").Remove "currentKey"
            i = i + 3
        Else
            j = i
            While j <= Len(json) And Mid(json, j, 1) <> "," And Mid(json, j, 1) <> "}" And Mid(json, j, 1) <> "]" And Mid(json, j, 1) <> ":"
                j = j + 1
            Wend
            numStr = Mid(json, i, j-i)
            If InStr(numStr, ".") > 0 Then
                jsonStack.Item("").Add jsonStack.Item("").Item("currentKey"), CDbl(numStr)
            Else
                jsonStack.Item("").Add jsonStack.Item("").Item("currentKey"), CInt(numStr)
            End If
            jsonStack.Item("").Remove "currentKey"
            i = j - 1
        End If
        i = i + 1
    Wend
    
    ' Get the value for the specified key
    getJsonValue = jsonDict(key)
End Function








'If key preexits then delete rather than update then add as new and remove the _old key category after the installation is finished then remove the keys from registry
'regdelete.bat for above last part
'key env RELAY_[KEYNAME]
'use this code as support file
'In userInterface before getMiddlewareInfo (check getMiddlewareInfosequence)
'Only on version upgrade scenario
'silent execution

