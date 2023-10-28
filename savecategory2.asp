<%
category2 = request.Form("category2")
brand = request.Form("brand")

if category2="" Then
response.redirect("inputcategory2.asp")
ElseIf brand="" Then
response.redirect("inputcategory2.asp")
End If

sql="SELECT * FROM category2  WHERE category2 ='"&category2&"' and brand ='"&brand&"';"


Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

If rs.eof And category2<>"" Then
		rs.addnew
		rs("category2") =category2
		rs("brand") =brand
		rs.update

	response.redirect "inputcategory2.asp"
elseif category2<>"" Then
	response.redirect "inputcategory2.asp"
Else
%>
<%End if%>

