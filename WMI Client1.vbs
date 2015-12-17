'
' Script para captura de dados do WMI
'
' @version 17/12/2015
' @link http://stackoverflow.com/questions/7668183/does-vbscript-support-introspection-for-objects
' 
'------------------------------------

Option Explicit

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

'
' busca os dados no WMI
'
Dim objWMIService
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Dim colItems
Set colItems = objWMIService.ExecQuery( _
     "SELECT * FROM Win32_SystemEnclosure" _
   , "WQL" _
   , wbemFlagReturnImmediately + wbemFlagForwardOnly _
)

'
' percorre os dados e tenta gerar JSON
'
Dim objItem
For Each objItem In colItems
    Dim oProp
    For Each oProp In objItem.Properties_
        WScript.Echo oProp.Name + ":'" + ToString( oProp.Value ) + "',"
    Next
    WScript.Echo
Next

'
' funcao que tenta tratar o tipo de dado
'
Function ToString( vX )
  ToString = ""
 On Error Resume Next
  ToString = CStr(vX)
 On Error GoTo 0
End Function
