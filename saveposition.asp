
<%
positionname = request.Form("positionname")
position = request.Form("position")

if positionname="" Then
response.redirect("intputposition.asp")
ElseIf positionname="" Then
response.redirect("intputposition.asp")
End If

sql="SELECT * FROM position  WHERE positionname ='"&positionname&"' and position ='"&position&"';"


Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

If rs.eof And positionname<>"" Then
		rs.addnew
		rs("positionname") =positionname
		rs("position") =position
		rs.update

	response.redirect "intputposition.asp"
elseif positionname<>"" Then
	response.redirect "intputposition.asp"

%>
<%End if%>

