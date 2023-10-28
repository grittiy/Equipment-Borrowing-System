<%
idcategory2 =request.Form("idcategory2")
category2 = request.Form("category2")
brand = request.Form("brand")


sql2="delete FROM category2  WHERE idcategory2 ='"&idcategory2&"' order by  idcategory2;"
sql="SELECT * FROM category2;"

Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs2 = Server.CreateObject("ADODB.Recordset")

rs2.Open sql2,conn,1,3
rs.Open sql,conn,1,3


		rs.addnew
		rs("idcategory2")=idcategory2
		rs("category2") =category2
		rs("brand") =brand

		rs.update
response.redirect "inputcategory2.asp"
%>