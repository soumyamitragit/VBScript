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


key = "datatochange"
data = extractJSONData(fileContent, key)
WScript.Echo "Finally -> " & data
