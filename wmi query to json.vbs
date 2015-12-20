'========================================
'
' Script para captura de dados do WMI
'
' @link https://github.com/e-ruiz/wmi-inventory
' @license MIT License (MIT) https://github.com/e-ruiz/wmi-inventory/blob/master/LICENSE
' @see http://stackoverflow.com/questions/7668183/does-vbscript-support-introspection-for-objects
' 
'========================================

inventoryItems = Array("Win32_Bios"_
                      ,"Win32_ComputerSystem"_
                      ,"Win32_OperatingSystem"_
                      ,"Win32_Processor"_
                      ,"Win32_DiskDrive"_
                      ,"Win32_LogicalDisk"_
                      ,"Win32_NetworkAdapter"_
                      ,"Win32_BaseBoard"_
                      )


For Each inventoryItem In inventoryItems
    wscript.echo "==["&inventoryItem&"]==============="
    'wscript.echo ToJson(getWMICollection(inventoryItem))
    
    For Each invItem In getWMICollection(inventoryItem)
        For Each oProp In invItem.Properties_
            wscript.echo oProp.Name & ": " & ToString(oProp.Value)
        Next
    Next
    
    wscript.echo ""
    wscript.echo ""
    
Next


'
' WMI Connection
'
'
Function getWMICollection(strCollection)
	
    Set stdOut = WScript.StdOut
    stdOut.writeline ""
	
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
	
    If Err.Number <> 0 Then
        stdOut.writeline Now()&" - "&  "WMI Connection error #"&Err.Number&" on "&CompName&": "&Err.Description
		Err.Clear
    Else
        stdOut.writeline Now()&" - "&  "WMI connection to PC was successfull."
        stdOut.writeline "Getting " & strCollection & " data"
        Dim colItems
        Set colItems = objWMIService.ExecQuery(_
             "SELECT * FROM " & strCollection)

        Set getWMICollection = colItems
	End If
End Function


'
' Percorre os dados do WMI e tenta gerar JSON
'
Function ToJson(WMIColection)
    Dim objItem
    Dim json
    For Each objItem In WMIColection
        Dim oProp        
        For Each oProp In objItem.Properties_
            json = json & oProp.Name & ":'" & ToString( oProp.Value ) & "',"
        Next
        'WScript.Echo
    Next
    
    ' adiciona chaves no inicio e fim,
    ' tambem remove a ultima virgula que sobra depois do loop
    json = "{" & Left(json, Len(json)-1) & "}"
    
    ToJson = json
End Function

'
' funcao que tenta converter para string
' contem bugs, dependendo do tipo de dado de entrada
'
Function ToString( vX )
    ToString = ""
    On Error Resume Next
    ToString = CStr(vX)
    On Error GoTo 0
End Function
