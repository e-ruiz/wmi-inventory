'
' Exemplo de Requisição HTTP 
'
'
Set http = CreateObject("Microsoft.XMLHTTP")

url = "http://" & InputBox("http://")

wscript.echo "POST: " + url
postData = "apikey=" + "96464-4687-6351-7963519"
http.open "POST", url, False
http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
http.send postData

If http.status = 200 And http.readyState = 4 Then
    wscript.echo
    wscript.echo "Status: "  & http.readyState
    wscript.echo 
    wscript.echo "HTTP Status: " & http.status
    wscript.echo
    wscript.echo "Headers: " & http.getAllResponseHeaders()
    
    wscript.echo "Body: "
    wscript.echo http.responseText
Else
    wscript.echo "Error!"
    wscript.echo "HTTP Status: " & http.status
    wscript.echo "Status: "      & http.readyState
    wscript.echo "Headers: "     & http.getAllResponseHeaders()
    wscript.echo "Body: "        & http.responseText
End If

wscript.echo "Done!"
