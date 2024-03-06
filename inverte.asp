<%
Dim minhaString, resultado
minhaString = "Hello, World!"

Function InverteString(str)
    Dim i, resultado
    resultado = ""

    For i = Len(str) To 1 Step -1
        resultado = resultado & Mid(str, i, 1)
    Next

    InverteString = resultado
End Function

resultado = InverteString(minhaString)
Response.Write "String original: " & minhaString & "<br>"
Response.Write "String invertida: " & resultado
%>
