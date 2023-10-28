<%
'sql="delete  from facultystd WHERE idfac='"&request("idfac")&"' ;" 

Set conn = Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")

rs.Open "delete from borrow  WHERE idborrow='"+request("id")+"' ;" ,conn,1,3


'rs.Open sql,conn,1,3
response.redirect("searchborrow.asp")
%>
