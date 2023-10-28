<%
idtool=request.Form("idtool")
toolname=request.Form("toolname")
model=request.Form("model")

idcategory2=request.Form("idcategory2")

size=request.Form("size")
weight=request.Form("weight")
color=(request.Form("color"))

details=request.Form("details")
unitprice=request.Form("unitprice")
quantity=request.Form("quantity")

dayy=CDbl(request.Form("dayy"))
monthh=(request.Form("monthh"))
yearr=CDbl(request.Form("yearr"))
idate=request.Form("idate")

sql="SELECT * FROM tool order by idtool;"

sql1="delete  FROM tool  WHERE idtool ='"&idtool&"' order by idtool;"



Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql1,conn,1,3

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

 	rs.addnew
		rs("idtool") =idtool
		rs("toolname") =toolname
		rs("model") =model

		rs("idcategory2") =idcategory2

		rs("size") =size
		rs("weight") =weight
		rs("color") =color
		rs("details") =details
		rs("unitprice") =unitprice
		rs("quantity") =quantity

		rs("idate") =idate
		rs("datesave")=Now()
		rs.update
		response.redirect "searchtool.asp"
%>