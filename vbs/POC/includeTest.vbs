Dim FSO, file, scriptText
Set FSO = CreateObject("Scripting.FileSystemObject")

externalScriptFile = FSO.GetAbsolutePathName("UpdateRegistryEnvironment.vbs")

Set file = FSO.OpenTextFile(externalScriptFile)
scriptText = file.ReadAll()
file.Close

ExecuteGlobal scriptText

a = getRegistryKey("HKEY_CURRENT_USER\Environment\Test","PathToJson")
WScript.Echo a