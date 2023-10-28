<%
idgenre =request.Form("idgenre")
genre = request.Form("genre")


sql2="delete FROM genre  WHERE idgenre ='"&idgenre&"' order by  idgenre;"
sql="SELECT * FROM genre;"

Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs2 = Server.CreateObject("ADODB.Recordset")

rs2.Open sql2,conn,1,3
rs.Open sql,conn,1,3


		rs.addnew
		rs("idgenre")=idgenre
		rs("genre") =genre

		rs.update
response.redirect "inputgenre.asp"
%>