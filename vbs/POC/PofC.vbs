'To Set Key Value in Registry
Function setRegistryKey(regPath,KeyToRead,valueToWrite)
    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    wshShell.RegWrite regPath & "\" & KeyToRead, valueToWrite, "REG_SZ"
End Function


'To Get Value For a Key from Registry
Function getRegistryKey(regPath, key)
    Set shellObject = CreateObject("WScript.Shell")
    On Error Resume Next
    getRegistryKeyValue = shellObject.RegRead(regPath & "\" & key)
    If Err.Number <> 0 Then
        WScript.Echo "Error"
        getRegistryKeyValue = ""
        Err.Clear
    End If
    getRegistryKey = getRegistryKeyValue
End Function

'To Match and Update Registry
Function checkMatchAndUpdate(regPath, key, matchValue, updateValue)
    If matchValue = updateValue Then
        WScript.Echo "Match"
    ElseIf matchValue = "" Then
        WScript.Echo "Empty"
        setRegistryKey registryPath, key, updateValue
    Else
        WScript.Echo "Did Not Match"
        setRegistryKey registryPath, key & "_old", matchValue
        setRegistryKey registryPath, key, updateValue
    End If
End Function

Dim existingKeyValue
Dim registryPath
Dim key
key = "TestChange"

key = "datatochange"
registryPath = "HKEY_CURRENT_USER\Environment\Test"
existingKeyValue = getRegistryKey(registryPath, key)
' I have to check diff bw set and not using anything
Function getFileContent(filePath)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set inputFile = FSO.OpenTextFile(filePath,1)
    fileContent = inputFile.ReadAll()
    inputFile.Close()
    WScript.Echo "Line 6 -> " & fileContent
    getFileContent = fileContent
End Function

Function removeQuotesFromJSONString(fileContent)
    fileContent = Replace(fileContent, """", " ")
    WScript.Echo  fileContent
    removeQuotesFromJSONString = fileContent
End Function

Function extractJSONData(fileContent, key)
    fileContent = removeQuotesFromJSONString(fileContent)
    WScript.Echo "Line21" & fileContent
    Set regex = New RegExp
    regex.Pattern = key & " : \s*([^,]*)"
    WScript.Echo regex.Pattern
    Set matches = regex.Execute(fileContent)
    WScript.Echo matches.Count
    For Each match in matches
        extractedData = match.SubMatches(0)
        extractedData = Replace(extractedData, "}", "")
        WScript.Echo extractedData
        extractJSONData = extractedData
    Next
End Function

filePath = "C:\Users\ve00ym279\OneDrive - YAMAHA MOTOR CO., LTD\Desktop\Scripts\vbs\POC\test_one.json"
fileContent = getFileContent(filePath)
WScript.Echo "Line17 -> " & fileContent

data = extractJSONData(fileContent, key)
WScript.Echo "Finally -> " & data
updateValue = "Updated Hello"
' If existingKeyValue <> "" Then
checkMatchAndUpdate registryPath, key, existingKeyValue, updateValue
' Else
    ' setRegistryKey registryPath, key, updateValue
' End If

' WScript.Echo existingKeyValue

' If existingKeyValue = "Hello" Then
'     WScript.Echo "Match"
' Else
'     WScript.Echo "Did Not Match"
'     setRegistryKey registryPath,"TestChange","Heloo"
' End If