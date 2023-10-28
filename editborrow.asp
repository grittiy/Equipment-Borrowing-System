<html>

<head>
<title>แก้ไขข้อมูลใบยืมเครื่องมือ</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<p align="center">&nbsp;<font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:28pt;"><b>แก้ไขข้อมูลใบยืมเครื่องมือ</b></span></font></p>
	<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT * FROM  borrow WHERE idborrow='"+request("id")+"';" ,conn,1,3
%>
<form method="post" action="showeditborrow.asp">
    <table align="center" width="637" cellpadding="0" cellspacing="0">
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รหัสใบยืมเครื่องมือ</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p><font face="TH SarabunPSK" color="#6600CC"><span style="font-size:16pt;"><input  type="hidden" name="idborrow" maxlength="13" size="15" style="font-family:SOV_Thanamas; font-size:20; color:blue; background-color:silver; border-style:outset;" value='<%=rs("idborrow")%>'
		><%=rs("idborrow")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>สมาชิก</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
<p><font face="TH Baijam"><span style="font-size:20pt;"><select name="idmember" size="1" type="text" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255); border-style:outset;">
		<%
			sql1="SELECT * FROM member order by idmember;"

			Set conn =Server.CreateObject("ADODB.Connection")
			conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

			Set rs1 = Server.CreateObject("ADODB.Recordset")
			rs1.Open sql1,conn,1,3
				
			Do While Not rs1.eof
			
		%>

            <option value="<%=rs1("idmember")%>"><%=rs1("pname")%><%=rs1("fname")%>&nbsp;<%=rs1("lname")%>&nbsp;(อายุ<%=rs1("age")%>)&nbsp;<%=rs1("agency")%>&nbsp;โทรสาร[<%=rs1("fax")%>]</CENTER>
			</option>

			<%
			rs1.movenext
			Loop
			%>


            </select></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>เครื่องมือ</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
<p><font face="TH Baijam"><span style="font-size:20pt;"><select name="idtool" size="1" type="text" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255); border-style:outset;">
		<%
			sql2="SELECT * FROM tool order by idtool;"

			Set conn =Server.CreateObject("ADODB.Connection")
			conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

			Set rs2 = Server.CreateObject("ADODB.Recordset")
			rs2.Open sql2,conn,1,3
				
			Do While Not rs2.eof
			
		%>

            <option value="<%=rs2("idtool")%>"><%=rs2("toolname")%>&nbsp;รุ่น<%=rs2("model")%>&nbsp;สี<%=rs2("color")%>&nbsp;ราคาต่อหน่วย&nbsp;<%=rs2("unitprice")%>&nbsp;บาท</CENTER>
			</option>

			<%
			rs2.movenext
			Loop
			%>


            </select></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>เจ้าหน้าที่</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
<p><font face="TH Baijam"><span style="font-size:20pt;"><select name="idofficer" size="1" type="text" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255); border-style:outset;">
		<%
			sql3="SELECT * FROM office order by idoffice;"

			Set conn =Server.CreateObject("ADODB.Connection")
			conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

			Set rs3 = Server.CreateObject("ADODB.Recordset")
			rs3.Open sql3,conn,1,3
				
			Do While Not rs3.eof
			
		%>

            <option value="<%=rs3("idoffice")%>"><%=rs3("pname")%><%=rs3("fname")%>&nbsp;<%=rs3("lname")%>&nbsp;(อายุ<%=rs3("age")%>)&nbsp;เบอร์โทร[<%=rs3("phone")%>]</CENTER>
			</option>

			<%
			rs3.movenext
			Loop
			%>


            </select></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>วันที่ยืม</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p>&nbsp;<font face="SOV_Thanamas" color="blue"><span style="font-size:16pt;"><select name="dayy" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255);">
                <option value="1">
		<%
		x=1 
		Do While x<=31
		%>

                    </option>
<option value="<%=x%>" <%if x=Day(rs("bdate"))then%>selected<%end if%>><%=x%></option>
		<%
		x=x+1
		loop
		%></option>
                </select> <select name="monthh" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255);">
                 <option value="01" <%If Cint(month(rs("bdate")))=1 then%>selected<%End if%>>มกราคม</option>
    <option value="02" <%If Cint(month(rs("bdate")))=2 then%>selected<%End if%>>กุมภาพันธ์</option>
    <option value="03" <%If Cint(month(rs("bdate")))=3 then%>selected<%End if%>>มีนาคม</option>
    <option value="04" <%If Cint(month(rs("bdate")))=4 then%>selected<%End if%>>เมษายน</option>
    <option value="05" <%If Cint(month(rs("bdate")))=5 then%>selected<%End if%>>พฤษภาคม</option>
    <option value="06" <%If Cint(month(rs("bdate")))=6 then%>selected<%End if%>>มิถุนายน</option>
    <option value="07" <%If Cint(month(rs("bdate")))=7 then%>selected<%End if%>>กรกฏาคม</option>
    <option value="08" <%If Cint(month(rs("bdate")))=8 then%>selected<%End if%>>สิงหาคม</option>
    <option value="09" <%If Cint(month(rs("bdate")))=9 then%>selected<%End if%>>กันยายน</option>
    <option value="10" <%If Cint(month(rs("bdate")))=10 then%>selected<%End if%>>ตุลาคม</option>
    <option value="11" <%If Cint(month(rs("bdate")))=11 then%>selected<%End if%>>พฤศจิกายน</option>
    <option value="12" <%If Cint(month(rs("bdate")))=12 then%>selected<%End if%>>ธันวาคม</option>
                </select> <select name="yearr" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255);">
                <option value="2018"><%
		y=1000 
		Do While y<=9999
		%>

                    </option>
<option value="<%=y%>" <%if y=CInt(Year(rs("bdate"))) then%>selected<%End if%>><%=y%></option>
		<%
		y=y+1
		loop
		%></option>
                </select></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>จำนวนเครื่องมือที่ยืม</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="quantity" maxlength="2" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,255);" size="5" value='<%=rs("quantity")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>วันที่คืน</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p>&nbsp;<font face="SOV_Thanamas" color="blue"><span style="font-size:16pt;"><select name="dayy2" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255);">
                <option value="1">
		<%
		x=1 
		Do While x<=31
		%>

                    </option>
<option value="<%=x%>" <%if x=Day(rs("edate"))then%>selected<%end if%>><%=x%></option>
		<%
		x=x+1
		loop
		%></option>
                </select> <select name="monthh2" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255);">
                 <option value="01" <%If Cint(month(rs("edate")))=1 then%>selected<%End if%>>มกราคม</option>
    <option value="02" <%If Cint(month(rs("edate")))=2 then%>selected<%End if%>>กุมภาพันธ์</option>
    <option value="03" <%If Cint(month(rs("edate")))=3 then%>selected<%End if%>>มีนาคม</option>
    <option value="04" <%If Cint(month(rs("edate")))=4 then%>selected<%End if%>>เมษายน</option>
    <option value="05" <%If Cint(month(rs("edate")))=5 then%>selected<%End if%>>พฤษภาคม</option>
    <option value="06" <%If Cint(month(rs("edate")))=6 then%>selected<%End if%>>มิถุนายน</option>
    <option value="07" <%If Cint(month(rs("edate")))=7 then%>selected<%End if%>>กรกฏาคม</option>
    <option value="08" <%If Cint(month(rs("edate")))=8 then%>selected<%End if%>>สิงหาคม</option>
    <option value="09" <%If Cint(month(rs("edate")))=9 then%>selected<%End if%>>กันยายน</option>
    <option value="10" <%If Cint(month(rs("edate")))=10 then%>selected<%End if%>>ตุลาคม</option>
    <option value="11" <%If Cint(month(rs("edate")))=11 then%>selected<%End if%>>พฤศจิกายน</option>
    <option value="12" <%If Cint(month(rs("edate")))=12 then%>selected<%End if%>>ธันวาคม</option>
                </select> <select name="yearr2" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255);">
                <option value="2018"><%
		y=1000 
		Do While y<=9999
		%>

                    </option>
<option value="<%=y%>" <%if y=CInt(Year(rs("edate"))) then%>selected<%End if%>><%=y%></option>
		<%
		y=y+1
		loop
		%></option>
                </select></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>จำนวนเงิน</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="amount" maxlength="10" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,255);" size="10" value='<%=rs("amount")%>'></font></p>
            </td>
        </tr>
    </table>
<p align="center"><font face="TH Baijam"><input type="submit" name="ตกลง" value="ตกลง" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);"> 
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="ยกเลิก" value="ยกเลิก" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(102,0,102); background-color:rgb(51,204,255);">&nbsp;</font>&nbsp;</p>
</FORM>
</body>

</html>
