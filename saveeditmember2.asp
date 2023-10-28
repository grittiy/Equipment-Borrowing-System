<%idmember=request("idmember")
sql="SELECT * FROM member  WHERE idmember ='"&idmember&"' order by idmember;"

Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

%>
<%
idmember=request.Form("idmember")

pname=request.Form("pname")
fname=request.Form("fname")
lname=request.Form("lname")
sex=request.Form("sex")
age=request.Form("age")

idposition=request.Form("idposition")

person=(request.Form("person"))
agency=request.Form("agency")
address=request.Form("address")
phone=request.Form("phone")
fax=(request.Form("fax"))
email=(request.Form("email"))
password=(request.Form("password"))

bdate=(request.Form("bdate"))


dayy=CDbl(request.Form("dayy"))
monthh=(request.Form("monthh"))
yearr=CDbl(request.Form("yearr"))
status="user"

sql="SELECT * FROM member order by idmember;"

sql1="delete  FROM member  WHERE idmember ='"&idmember&"' order by idmember;"




Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql1,conn,1,3

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

 	rs.addnew
		rs("idmember") =idmember
		rs("pname") =pname
		rs("fname") =fname
		rs("lname") =lname
		rs("sex") =sex
		rs("age") =age
		rs("status") =status


		rs("idposition") =idposition
		rs("password") =password
		rs("person") =person
		rs("agency") =agency
		rs("address") =address
		rs("phone") =phone
		rs("fax") =fax
		rs("email") =email
		rs("bdate") =bdate
		rs("datesave")=Now()
		rs.update
		response.redirect ("showallmember2.asp?idmember="&idmember&"")
%>