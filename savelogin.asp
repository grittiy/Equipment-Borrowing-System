<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
</head>
<body>
<%

    email = request.Form("email")
    password = request.Form("password")

if email="" Then
response.redirect("login.asp")
ElseIf password="" Then
response.redirect("login.asp")
End If



    sql="SELECT * FROM member  WHERE email ='"&email&"' AND password ='"&password&"'  ;"

    Set conn =Server.CreateObject("ADODB.Connection")
    conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql,conn,1,3

If Not  rs.eof Then 

 email=rs("email")
    Password=rs("password")
    status=rs("status")

    
   If status="user"Then
   idmember=CInt(rs("idmember"))
'response.redirect("menuborrow2565.asp?id="&id&"")
response.redirect("user.asp?idmember="&idmember&"")
 End If 
 else
 response.redirect("login.asp")
 End if
   'End If
      
 ' Else 
' response.redirect("index.asp")
    
    %>


	</body>
</html>