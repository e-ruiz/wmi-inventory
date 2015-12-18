'
' Script para captura de dados do WMI
'
' @version 17/12/2015
' @link http://stackoverflow.com/questions/7668183/does-vbscript-support-introspection-for-objects
' 
'------------------------------------


'
' Consulta WMI
' 
' @return Collection
'
Function WMIColection(collection)
    
    wbemFlagReturnImmediately = &h10
    wbemFlagForwardOnly = &h20
    strSelect = "SELECT * FROM " & collection
    
    Dim objWMIService
    Dim colItems
    
    '
    ' busca os dados no WMI
    '
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery( _
         strSelect _
       , "WQL" _
       , wbemFlagReturnImmediately + wbemFlagForwardOnly _
    )
    
    Set WMIColection = colItems
    
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
        WScript.Echo
    Next
    json = "{" & Left(json, Len(json)-1) & "}"
    
    ToJson = json
End Function

'
' funcao que tenta tratar o tipo de dado
'
Function ToString( vX )
  ToString = ""
 On Error Resume Next
  ToString = CStr(vX)
 On Error GoTo 0
End Function
