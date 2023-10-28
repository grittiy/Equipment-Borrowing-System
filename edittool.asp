<html>

<head>
<title>แก้ไขข้อมูลเครื่องมือ</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<p align="center">&nbsp;<font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:28pt;"><b>แก้ไขข้อมูลเครื่องมือ</b></span></font></p>
	<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT * FROM  tool WHERE idtool='"+request("id")+"';" ,conn,1,3
%>
<form method="post" action="showedittool.asp">
    <table align="center" width="582" cellpadding="0" cellspacing="0">
        <tr>
            <td width="174" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รหัสเครื่องมือ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="350" height="41">
                <p> <font face="TH SarabunPSK" color="#6600CC"><span style="font-size:16pt;"><input  type="hidden" name="idtool" maxlength="13" size="15" style="font-family:SOV_Thanamas; font-size:20; color:blue; background-color:silver; border-style:outset;" value='<%=rs("idtool")%>'
		><%=rs("idtool")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="174" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ชื่อเครื่องมือ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="350" height="41">
                <p> <font face="TH Baijam"><input type="text" name="toolname" maxlength="100" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="20" value='<%=rs("toolname")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="174" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ชื่อรุ่น</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="350" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="model" maxlength="100" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="20" value='<%=rs("model")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="174" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>หมวดหมู่เครื่องมือ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="350" height="41">
<p><font face="TH Baijam"><span style="font-size:20pt;"><select name="idcategory2" size="1" type="text" style="font-family:'TH Mali Grade 6'; font-size:20; color:maroon; background-color:rgb(204,153,255); border-style:outset;">
		<%
			sql1="SELECT * FROM category2 order by idcategory2;"

			Set conn =Server.CreateObject("ADODB.Connection")
			conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

			Set rs1 = Server.CreateObject("ADODB.Recordset")
			rs1.Open sql1,conn,1,3
				
			Do While Not rs1.eof
			
		%>

            <option value="<%=rs1("idcategory2")%>" <%if CInt(rs1("idcategory2"))=rs("idcategory2")  then%>selected<%End if%>><%=rs1("category2")%>&nbsp;[<%=rs1("brand")%>]</CENTER>
			</option>

			<%
			rs1.movenext
			Loop
			%>


            </select></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="174" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ขนาด</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="350" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="size" maxlength="30" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="20" value='<%=rs("size")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="174" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>น้ำหนัก</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="350" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="weight" maxlength="7" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="10" value='<%=rs("weight")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="174" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>สี</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="350" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="color" maxlength="30" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="20" value='<%=rs("color")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="174" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รายละเอียด</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="350" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="details" maxlength="225" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="20" value='<%=rs("details")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="174" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ราคาต่อหน่วย</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="350" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="unitprice" maxlength="10" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="10" value='<%=rs("unitprice")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="174" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>&nbsp;จำนวนในคลัง</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="350" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="quantity" maxlength="2" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="5" value='<%=rs("quantity")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="174" height="43">
                <p align="right"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>วันที่เข้าคลัง</b></span></font></p>
            </td>
            <td width="58" height="43">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="350" height="43">
                <p><font face="TH Baijam" color="blue"><span style="font-size:16pt;"><select name="dayy" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:maroon; background-color:rgb(204,153,255);">
                <option value="1">
		<%
		x=1 
		Do While x<=31
		%>

                    </option>
<option value="<%=x%>" <%if x=Day(rs("idate"))then%>selected<%end if%>><%=x%></option>
		<%
		x=x+1
		loop
		%></option>
                </select> <select name="monthh" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:maroon; background-color:rgb(204,153,255);">
                 <option value="01" <%If Cint(month(rs("idate")))=1 then%>selected<%End if%>>มกราคม</option>
    <option value="02" <%If Cint(month(rs("idate")))=2 then%>selected<%End if%>>กุมภาพันธ์</option>
    <option value="03" <%If Cint(month(rs("idate")))=3 then%>selected<%End if%>>มีนาคม</option>
    <option value="04" <%If Cint(month(rs("idate")))=4 then%>selected<%End if%>>เมษายน</option>
    <option value="05" <%If Cint(month(rs("idate")))=5 then%>selected<%End if%>>พฤษภาคม</option>
    <option value="06" <%If Cint(month(rs("idate")))=6 then%>selected<%End if%>>มิถุนายน</option>
    <option value="07" <%If Cint(month(rs("idate")))=7 then%>selected<%End if%>>กรกฏาคม</option>
    <option value="08" <%If Cint(month(rs("idate")))=8 then%>selected<%End if%>>สิงหาคม</option>
    <option value="09" <%If Cint(month(rs("idate")))=9 then%>selected<%End if%>>กันยายน</option>
    <option value="10" <%If Cint(month(rs("idate")))=10 then%>selected<%End if%>>ตุลาคม</option>
    <option value="11" <%If Cint(month(rs("idate")))=11 then%>selected<%End if%>>พฤศจิกายน</option>
    <option value="12" <%If Cint(month(rs("idate")))=12 then%>selected<%End if%>>ธันวาคม</option>
                </select> &nbsp;<select name="yearr" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:maroon; background-color:rgb(204,153,255);">
                <option value="2018"><%
		y=1000 
		Do While y<=9999
		%>

                    </option>
<option value="<%=y%>" <%if y=CInt(Year(rs("idate")))  then%>selected<%End if%>><%=y%></option>
		<%
		y=y+1
		loop
		%></option>
                </select></span></font></p>
            </td>
        </tr>
    </table>

<p align="center"><font face="TH Sarabun New"><input type="submit" name="แก้ไขข้อมูล" value="แก้ไขข้อมูล" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);"></font></p>
</FORM>
</body>

</html>
