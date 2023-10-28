<%
dele = request.form("dele")

if dele="" Then
response.redirect("searchborrow.asp")
Else

sql="delete  from borrow WHERE idborrow in ("& dele &")"

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")

'rs.Open "delete  from province where idprovince in (" & dele & ")" ,conn,1,3
rs.Open sql,conn,1,3

response.redirect("searchborrow.asp")
End if
%>
