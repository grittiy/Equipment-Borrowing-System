<!--#include file=getupload.asp-->

<%


idtool = uploaddata.Item("idtool").Item("value")
contenttype = uploaddata.Item("pict").Item("contenttype")
picture = TextToBinary(uploaddata.Item("pict").Item("value"))



sql="SELECT * FROM tool  WHERE idtool="&idtool&";"
sql1="DELETE  FROM tool  WHERE idtool="&idtool&";"
sql2="SELECT   *  FROM  tool"
Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

		
         idtool =rs("idtool")
		 toolname=rs("toolname")
		 model=rs("model")
		 idcategory2=rs("idcategory2")
		 size = rs("size")
		 weight=rs("weight")
		 color=rs("color")
		 details=rs("details")
		 unitprice=rs("unitprice")
		 quantity=rs("quantity")
		 idate=rs("idate")
	   

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql1,conn,1,3

Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.Open sql2,conn,1,3

	rs2.addnew
	
		
rs2("idtool") =idtool
		rs2("toolname") =toolname
		rs2("model") =model
		rs2("idcategory2") =idcategory2
		rs2("size") =size
		rs2("datesave") =Now()
		rs2("weight") =weight
		rs2("color") =color
		rs2("details") =details
		rs2("unitprice") =unitprice
			rs2("quantity") =quantity
			rs2("idate") =idate
	If LenB(picture)<>0 Then
	 rs2("contenttype") = contenttype
	 rs2("picture").AppendChunk=picture&chrB(0)
	 end If
	rs2.update

	response.redirect"searchtool.asp"

%>