<%
genre = request.Form("genre")

if genre="" Then
response.redirect("inputgenre.asp")

End If

sql="SELECT * FROM genre  WHERE genre ='"&genre&"';"




Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

If rs.eof And genre<>"" Then
		rs.addnew
		rs("genre") =genre
		rs.update

	response.redirect "inputgenre.asp"
elseif genre<>"" Then
	response.redirect "inputgenre.asp"
Else
%>
<%End if%>

