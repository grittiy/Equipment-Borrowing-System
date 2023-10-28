<!--#include file=getupload.asp-->

<%


idoffice = uploaddata.Item("idoffice").Item("value")
contenttype = uploaddata.Item("pict").Item("contenttype")
picture = TextToBinary(uploaddata.Item("pict").Item("value"))



sql="SELECT * FROM office  WHERE idoffice="&idoffice&";"
sql1="DELETE  FROM office  WHERE idoffice="&idoffice&";"
sql2="SELECT   *  FROM  office"
Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

		
         idoffice =rs("idoffice")
		 pname=rs("pname")
		 fname=rs("fname")
		 lname=rs("lname")
		 sex = rs("sex")
		 age=rs("age")
		 address=rs("address")
		 idgenre=rs("idgenre")
		 phone=rs("phone")
		 sdate=rs("sdate")
		 email=rs("email")
		  status=rs("status")

	   

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql1,conn,1,3

Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.Open sql2,conn,1,3

	rs2.addnew
	
		
rs2("idoffice") =idoffice
		rs2("pname") =pname
		rs2("fname") =fname
		rs2("lname") =lname
		rs2("sex") =sex
		rs2("datesave") =Now()
		rs2("age") =age
		rs2("idgenre") =idgenre
		rs2("address") =address
		rs2("phone") =phone
			rs2("email") =email
			rs2("sdate") =sdate
			rs2("status") =status

	If LenB(picture)<>0 Then
	 rs2("contenttype") = contenttype
	 rs2("picture").AppendChunk=picture&chrB(0)
	 end If
	rs2.update

	response.redirect"searchofficer.asp"

%>