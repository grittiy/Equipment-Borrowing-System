<%
idborrow=request.Form("idborrow")
idmember=CDbl(request.Form("idmember"))
idofficer=CDbl(request.Form("idofficer"))

idtool=request.Form("idtool")

quantity=request.Form("quantity")
amount=request.Form("amount")

bdate=(request.Form("bdate"))
edate=(request.Form("edate"))


dayy=CDbl(request.Form("dayy"))
monthh=(request.Form("monthh"))
yearr=CDbl(request.Form("yearr"))

dayy2=(request.Form("dayy2"))
monthh2=(request.Form("monthh2"))
yearr2=(request.Form("yearr2"))


sql="SELECT * FROM borrow order by idborrow;"

sql1="delete  FROM borrow  WHERE idborrow ='"&idborrow&"' order by idborrow;"



Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql1,conn,1,3

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

 	rs.addnew
	
rs("idborrow") =idborrow
		rs("idmember") =idmember
		rs("idofficer") =idofficer
		rs("idtool") =idtool
		rs("quantity") =quantity
		rs("amount") =amount
		rs("bdate") =bdate
		rs("edate") =edate

		rs("datesave")=Now()
		rs.update
		response.redirect "searchborrow.asp"
%>