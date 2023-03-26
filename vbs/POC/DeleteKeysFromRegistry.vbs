Function getRegistryKeys(registryParent, keyPath)
    value = checkForValue(registryParent)
    Dim strComputer 
    strComputer = "."
    Set registryObject = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")
    ' MsgBox wmi

    Dim regProvider
    Set regProvider = registryObject.Get("StdRegProv")

    Dim subKeys
    regProvider.EnumValues value, keyPath, subKeys

    getRegistryKeys = subKeys
End Function

Function deleteRegistryKey(regParent,regPath,keyToDelete)
    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    ' registryPath = Split(regPath,"\")
    ' pathVal = checkForValue(registryPath(0))
    wshShell.RegDelete regParent & "\" & regPath & "\" & keyToDelete
End Function

Function deleteAllRegistryKeysWithPrefix(regParent,regPath,keyPrefix)
    subKeys = getRegistryKeys(regParent,regPath)
    subKeysToDelete = prefixSearch(subKeys,keyPrefix)
    For i = 0 To UBound(subKeysToDelete)
        Dim valueData
        WScript.Echo subKeysToDelete(i) & " IS BEING DELETED "
        deleteRegistryKey regParent,regPath,subKeysToDelete(i)
    Next
End Function

Function prefixSearch(subKeys,keyPrefix)
    Dim subKeysToDelete()
    ' Empting out the array
    ReDim subKeysToDelete(-1)
    If Not IsNull(subKeys) Then
        For i = 0 To UBound(subKeys)
            If Left(subKeys(i), Len(keyPrefix)) = keyPrefix Then
                Redim Preserve subKeysToDelete(UBound(subKeysToDelete)+1)
                subKeysToDelete(UBound(subKeysToDelete)) = subKeys(i)
            End If
        Next
    End If
    prefixSearch = subKeysToDelete
End Function

Function checkForValue(registryParent)
    If registryParent = "HKEY_LOCAL_MACHINE" Then
        ' WScript.Echo registryParent
        Const checkForValueConst1 = &H80000002
        checkForValue = checkForValueConst1
    ElseIf registryParent = "HKEY_CURRENT_USER" Then
        ' WScript.Echo registryParent
        Const checkForValueConst2 = &H80000001
        checkForValue = checkForValueConst2
    End If  
End Function
' Increase Prefix Count By 1 when you add a new prefix
prefixCount = 2
Dim prefixArray()
ReDim prefixArray(prefixCount-1)
' Add your Prefix as prefixArray(Increment of last integer)
prefixArray(0) = "RELAY_PROPS_"
prefixArray(1) = "RELAY_PROPERTIES_"
For i = 0 To UBound(prefixArray)
    deleteAllRegistryKeysWithPrefix "HKEY_CURRENT_USER","Environment\test",prefixArray(i)
Next