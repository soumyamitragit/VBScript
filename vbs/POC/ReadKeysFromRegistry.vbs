Function getRegistryKeys(regPath)
    Const HKEY_LOCAL_MACHINE = &H80000001
    Dim strComputer 
    strComputer = "."
    Set registryObject = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")
    ' MsgBox wmi
    Dim keyPath
    keyPath = "Environment\test"

    Dim regProvider
    Set regProvider = registryObject.Get("StdRegProv")

    Dim subKeys
    regProvider.EnumValues HKEY_LOCAL_MACHINE, keyPath, subKeys

    For i = 0 To UBound(subKeys)
        Dim valueData
        ' regProvider.GetStringValue HKEY_LOCAL_MACHINE, keyPath, subKeys(i), valueData
        ' WScript.Echo subKeys(i) & "=" & valueData
        WScript.Echo subKeys(i)
    Next
End Function

getRegistryKeys ""