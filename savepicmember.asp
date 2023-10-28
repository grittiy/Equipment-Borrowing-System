<!--#include file=getupload.asp-->

<%


idmember = uploaddata.Item("idmember").Item("value")
contenttype = uploaddata.Item("pict").Item("contenttype")
picture = TextToBinary(uploaddata.Item("pict").Item("value"))



sql="SELECT * FROM member  WHERE idmember="&idmember&";"
sql1="DELETE  FROM member  WHERE idmember="&idmember&";"
sql2="SELECT   *  FROM  member"
Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

		
         idmember =rs("idmember")
		 pname=rs("pname")
		 fname=rs("fname")
		 lname=rs("lname")
		 sex = rs("sex")
		 age=rs("age")
		 person=rs("person")
		 agency=rs("agency")
		 address=rs("address")
		 idposition=rs("idposition")
		 phone=rs("phone")
		 bdate=rs("bdate")
		 fax=rs("fax")
		 email=rs("email")
		  status=rs("status")

	   

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql1,conn,1,3

Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.Open sql2,conn,1,3

	rs2.addnew
	
		
rs2("idmember") =idmember
		rs2("pname") =pname
		rs2("fname") =fname
		rs2("lname") =lname
		rs2("sex") =sex
		rs2("datesave") =Now()
		rs2("age") =age
		rs2("idposition") =idposition
		rs2("person") =person
		rs2("agency") =agency
		rs2("address") =address
		rs2("phone") =phone
			rs2("fax") =fax
			rs2("email") =email
			rs2("bdate") =bdate
			rs2("status") =status

	If LenB(picture)<>0 Then
	 rs2("contenttype") = contenttype
	 rs2("picture").AppendChunk=picture&chrB(0)
	 end If
	rs2.update

	response.redirect"searchmember.asp"

%>