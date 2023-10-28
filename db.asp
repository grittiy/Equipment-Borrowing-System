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



    sql="SELECT * FROM office WHERE email ='"&email&"' AND password ='"&password&"' AND idoffice ='"&idoffice&"' ;"
	sql2="SELECT * FROM member WHERE email ='"&email&"' AND password ='"&password&"' ;"

    Set conn =Server.CreateObject("ADODB.Connection")
    conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql,conn,1,3

	Set rs2 = Server.CreateObject("ADODB.Recordset")
    rs2.Open sql2,conn,1,3

 'Username=rs("username")
    'Password=rs("password")
    'status=rs("status")

    
  If rs.eof Then
	If rs2.eof Then
		response.redirect("menuborrow2565.asp")
	Else 

		response.redirect("loginadmin.asp")
	End If

  Else 
	response.redirect("login.asp")
  End If 
   'End If
      
 ' Else 
' response.redirect("index.asp")
    
    %>


	</body>
</html>