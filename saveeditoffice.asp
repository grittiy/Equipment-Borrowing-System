<%
idoffice=request.Form("idoffice")

pname=request.Form("pname")
fname=request.Form("fname")
lname=request.Form("lname")
sex=request.Form("sex")
age=request.Form("age")

idgenre=request.Form("idgenre")

address=request.Form("address")
phone=request.Form("phone")
email=(request.Form("email"))
sdate=(request.Form("sdate"))

password=request.Form("password")

dayy=CDbl(request.Form("dayy"))
monthh=(request.Form("monthh"))
yearr=CDbl(request.Form("yearr"))
status="admin"

sql="SELECT * FROM office order by idoffice;"

sql1="delete  FROM office  WHERE idoffice ='"&idoffice&"' order by idoffice;"



Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql1,conn,1,3

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

 	rs.addnew
		rs("idoffice") =idoffice
		rs("pname") =pname
		rs("fname") =fname
		rs("lname") =lname
		rs("sex") =sex
		rs("age") =age
		rs("status") =status
				rs("password") =password


		rs("idgenre") =idgenre

		rs("address") =address
		rs("phone") =phone
		rs("email") =email
		rs("sdate") =sdate
		rs("datesave")=Now()
		rs.update
		response.redirect "searchofficer.asp"
%>