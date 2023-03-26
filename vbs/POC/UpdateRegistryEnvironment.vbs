' Function to Read Data From Registry
Function getRegistryKey(regPath, key)
    Set shellObject = CreateObject("WScript.Shell")
    On Error Resume Next
    getRegistryKeyValue = shellObject.RegRead(regPath & "\" & key)
    If Err.Number <> 0 Then
        ' WScript.Echo "Error"
        getRegistryKeyValue = ""
        Err.Clear
    End If
    getRegistryKey = getRegistryKeyValue
End Function

' Function to Read All Data From Registry and Delete All that matches the given Prefix
Function deleteAllRegistryKeysWithPrefix(regParent, regPathChild , prefixMatch)
    Dim strComputer 
    strComputer = "."
    Dim arrKeys
    Set registryObject = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    Const parentVal = &H80000001'checkForValue(regParent)
    WScript.Echo parentVal
    WScript.Echo regPathChild
    registryObject.EnumKey parentVal, "Environment", arrKeys
    WScript.Echo "arrKeys -> " & arrKeys
    For Each subKey in arrKeys
        If Left(subKey, Len(prefixMatch)) = prefixMatch Then
            WScript.Echo "Deleting -> " & subKey
            registryObject.DeleteKey parentVal, regPathChild & "\" & subKey
        End If
    Next
End Function

' Function to Write Data To Registry
Function setRegistryKey(regPath,KeyToRead,valueToWrite)
    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    wshShell.RegWrite regPath & "\" & KeyToRead, valueToWrite, "REG_SZ"
End Function

' Function to Delete Keys in Regisrty
Function deleteRegistryKey(regPath,keyToDelete)
    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    ' registryPath = Split(regPath,"\")
    ' pathVal = checkForValue(registryPath(0))
    wshShell.RegDelete regPath & "\" & keyToDelete
End Function

Function deleteRegistryKeyDirect(regParent,regPath,KeyToDelete)
    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    pathVal = checkForValue(regParent)
    wshShell.RegDelete pathVal, regPath & "\" & KeyToDelete
End Function

Function checkForValue(registryParent)
    If registryParent = "HKEY_LOCAL_MACHINE" Then
        WScript.Echo registryParent
        Const checkForValueConst1 = &H80000002
        checkForValue = checkForValueConst1
    ElseIf registryParent = "HKEY_CURRENT_USER" Then
        WScript.Echo registryParent
        Const checkForValueConst2 = &H80000001
        checkForValue = checkForValueConst2
    End If  
End Function

' Function to Match and Update Registry If Exists then store previous value as _OLD suffix
Function checkMatchAndUpdate(regPath, key, matchValue, updateValue)
    If matchValue = updateValue Then
        ' WScript.Echo "Match"
    ElseIf matchValue = "" Then
        ' WScript.Echo "Empty"
        setRegistryKey registryPath, key, updateValue
    Else
        ' WScript.Echo "Did Not Match"
        setRegistryKey registryPath, key & "_OLD", matchValue
        setRegistryKey registryPath, key, updateValue
    End If
End Function

' Function to get file content
Function getFileContent(filePath)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set inputFile = FSO.OpenTextFile(filePath,1)
    fileContent = inputFile.ReadAll()
    inputFile.Close()
    ' WScript.Echo "Line 6 -> " & fileContent
    getFileContent = fileContent
End Function

' ' Function to extract JSON value for a given Key
' Function extractJSONData(fileContent, key)
'     fileContent = removeQuotesFromJSONString(fileContent)
'     ' WScript.Echo "Line21" & fileContent
'     Set regex = New RegExp
'     regex.Pattern = key & " : \s*([^,]*)"
'     ' WScript.Echo regex.Pattern
'     Set matches = regex.Execute(fileContent)
'     ' WScript.Echo matches.Count
'     For Each match in matches
'         extractedData = match.SubMatches(0)
'         extractedData = Replace(extractedData, "}", "")
'         ' WScript.Echo extractedData
'         extractJSONData = extractedData
'     Next
' End Function

' Function to remove Quotes from JSON String
Function removeQuotesFromJSONString(fileContent)
    fileContent = Replace(fileContent, """", " ")
    ' WScript.Echo  fileContent
    removeQuotesFromJSONString = fileContent
End Function

' Function getAllKeyValsFromJSONStringOld(fileContent)
'     fileContent = removeQuotesFromJSONString(fileContent)
'     fileContent = Replace(fileContent,"{","")
'     fileContent = Replace(fileContent,"}","")
'     fileContent = Replace(fileContent," ","")
'     WScript.Echo fileContent
'     Dim arrayOfData
'     arrayOfData = Split(fileContent,",")
'     arrayOfDataLen = UBound(arrayOfData) + 1
'     Dim dataDictonary
'     Set dataDictonary = CreateObject("Scripting.Dictionary")
'     For i=0 To arrayOfDataLen - 1
'         Dim currentKey
'         Dim currentKeyVal
'         currentKey = Split(arrayOfData(i), ":")(0)
'         currentKeyVal = Split(arrayOfData(i), ":")(1)
'         WScript.Echo currentKey & " -> " & currentKeyVal
'         currentKey = Replace(currentKey,vbLf,"")
'         dataDictonary.Add "RELAY_PROPERTIES" & UCase(currentKey),currentKeyVal
'     Next 'i
'     Set getAllKeyValsFromJSONString = dataDictonary
' End Function

Function getAllKeyValsFromJSONString(fileContent,prefix)
    fileContent = removeQuotesFromJSONString(fileContent)
    fileContent = Replace(fileContent,"{","")
    fileContent = Replace(fileContent,"}","")
    fileContent = Replace(fileContent," ","")
    WScript.Echo fileContent
    Set regex = New RegExp
    regex.Pattern = "([^:,]+):(\[.*?\]|[^,]+)"
    regex.Global = True
    set matches = regex.Execute(fileContent)
    Dim dataDictonary
    Set dataDictonary = CreateObject("Scripting.Dictionary")
    WScript.Echo matches.Count
    For Each match in matches
        Dim arrayOfData
        arrayOfData = Split(match,":")
        currentKey = arrayOfData(0)
        currentKeyVal = arrayOfData(1)
        WScript.Echo currentKey & "|->" & currentKeyVal
        dataDictonary.Add prefix & UCase(currentKey), currentKeyVal
    Next
    Set getAllKeyValsFromJSONString = dataDictonary
End Function

Dim existingKeyValue
Dim registryPath
Dim registryParent, registryPathChild

registryPath = "HKEY_CURRENT_USER\Environment\Test"
registryParent = "HKEY_CURRENT_USER"
registryPathChild = "Environment\Test"
' filePath = "C:\Users\Soumya Mitra\Documents\Work&Learn\Git\VBScript\vbs\POC\test_one.json"
filePathInRegistry = "HKEY_CURRENT_USER\Environment\Test"
filePath = getRegistryKey(filePathInRegistry,"PathToJson")
WScript.Echo filePath

MsgBox(filePath)

Dim keyValDict
Dim fileContent
fileContent = getFileContent(filePath)
' keyValDict = CreateObject("Scripting.Dictionary")
' getAllKeyValsFromJSONString fileContent
Dim prefixMatch
prefixMatch = "RELAY_PROPS_"
Set keyValDict = getAllKeyValsFromJSONString(fileContent,prefixMatch)
' Remove Keys if already exists


' deleteAllRegistryKeysWithPrefix registryParent,registryPathChild,prefixMatch

' If keyValDict.Count > 0 Then
'     For Each key in keyValDict.Keys
'         On Error Resume Next
'         WScript.Echo key & " will be removed"
'         ' deleteRegistryKey registryParent, registryPathChild, key
'         deleteRegistryKey registryPath, key
'         If Err.Number <> 0 Then
'         End If
'     Next
' End If


' Add New Keys
For Each key in keyValDict.Keys
    WScript.Echo key & "->" & keyValDict.Item(key)
    setRegistryKey registryPath,key,keyValDict.Item(key)
Next