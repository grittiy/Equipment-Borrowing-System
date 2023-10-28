<%
idposition = request.Form("idposition")
positionname = request.Form("positionname")
position = request.Form("position")

sql="SELECT * FROM  position order by idposition"

sql1="delete FROM position  WHERE idposition ='"&idposition&"' order by idposition;"

Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql1,conn,1,3

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

		rs.addnew
		rs("idposition") =idposition
		rs("positionname") =positionname
		rs("position") =position
		rs.update
response.redirect "intputposition.asp"
%>