<html>

<head>
<title>ระบบฐานข้อมูลเจ้าหน้าที่</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<p align="center"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:28pt;"><b>แก้ไขข้อมูลเจ้าหน้าที่</b></span></font></p>
	<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT * FROM  office WHERE idoffice='"+request("id")+"';" ,conn,1,3
%>
<form method="post" action="showeditofficer.asp">
    <table align="center" width="699" cellpadding="0" cellspacing="0">
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รหัสเจ้าหน้าที่</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
                <p><font face="TH SarabunPSK" color="#6600CC"><span style="font-size:16pt;"><input  type="hidden" name="idoffice" maxlength="13" size="15" style="font-family:SOV_Thanamas; font-size:20; color:blue; background-color:silver; border-style:outset;" value='<%=rs("idoffice")%>'
		><%=rs("idoffice")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ชื่อ-นามสกุล</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
                <p><font face="TH Baijam"><select name="pname" size="1" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);">
              <option value="นาย" <%If rs("pname")="นาย"then%>selected<%End if%>>นาย</option>
            <option value="นาง" <%If rs("pname")="นาง"then%>selected<%End if%>>นาง</option>
			<option value="นางสาว" <%If rs("pname")="นางสาว"then%>selected<%End if%>>นางสาว</option>
            </select> &nbsp;<input type="text" name="fname" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20" value='<%=rs("fname")%>'> 
                &nbsp;&nbsp;<input type="text" name="lname" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20" value='<%=rs("lname")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>อายุ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="age" maxlength="2" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="5" value='<%=rs("age")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>เพศ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
                <p><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><input type="radio" name="sex" value="ชาย" <%If rs("sex")="ชาย"then%>checked<%End if%>> 
            <b>ชาย &nbsp;&nbsp;&nbsp; 
            <input type="radio" name="sex" value="หญิง" <%If rs("sex")="หญิง"then%>checked<%End if%>> หญิง</b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ตำแหน่ง</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
<p><font face="TH Baijam"><span style="font-size:20pt;"><select name="idgenre" size="1" type="text" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(153,255,255); border-style:outset;">
		<%
			sql1="SELECT * FROM genre order by idgenre;"

			Set conn =Server.CreateObject("ADODB.Connection")
			conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

			Set rs1 = Server.CreateObject("ADODB.Recordset")
			rs1.Open sql1,conn,1,3
				
			Do While Not rs1.eof
			
		%>

            <option value="<%=rs1("idgenre")%>" <%if CInt(rs1("idgenre"))=rs("idgenre")  then%>selected<%End if%>><%=rs1("genre")%></CENTER>
			</option>

			<%
			rs1.movenext
			Loop
			%>


            </select></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">                            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ที่อยู่</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
                <p><font face="TH Baijam"><input type="text" name="address" maxlength="225" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20" value='<%=rs("address")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>อีเมล์</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
                <p><font face="TH Baijam"><input type="text" name="email" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20" value='<%=rs("email")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รหัสผ่าน</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
                <p><font face="TH Sarabun New"><input type="password" name="password" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20" value='<%=rs("password")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="43">
                <p align="right"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>หมายเลขโทรศัพท์</b></span></font></p>
            </td>
            <td width="58" height="43">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="43">
                <p><font face="TH Baijam"><input type="text" name="phone" maxlength="10" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20" value='<%=rs("phone")%>'>&nbsp;</font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="43">
                <p align="right"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>วันเริ่มปฏิบัติงาน</b></span></font></p>
            </td>
            <td width="58" height="43">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="43">
                <p><font face="TH Baijam" color="blue"><span style="font-size:16pt;"><select name="dayy" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(153,255,255);">
                <option value="1">
		<%
		x=1 
		Do While x<=31
		%>

                    </option>
<option value="<%=x%>" <%if x=Day(rs("sdate"))then%>selected<%end if%>><%=x%></option>
		<%
		x=x+1
		loop
		%></option>
                </select> <select name="monthh" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(153,255,255);">
                 <option value="01" <%If Cint(month(rs("sdate")))=1 then%>selected<%End if%>>มกราคม</option>
    <option value="02" <%If Cint(month(rs("sdate")))=2 then%>selected<%End if%>>กุมภาพันธ์</option>
    <option value="03" <%If Cint(month(rs("sdate")))=3 then%>selected<%End if%>>มีนาคม</option>
    <option value="04" <%If Cint(month(rs("sdate")))=4 then%>selected<%End if%>>เมษายน</option>
    <option value="05" <%If Cint(month(rs("sdate")))=5 then%>selected<%End if%>>พฤษภาคม</option>
    <option value="06" <%If Cint(month(rs("sdate")))=6 then%>selected<%End if%>>มิถุนายน</option>
    <option value="07" <%If Cint(month(rs("sdate")))=7 then%>selected<%End if%>>กรกฏาคม</option>
    <option value="08" <%If Cint(month(rs("sdate")))=8 then%>selected<%End if%>>สิงหาคม</option>
    <option value="09" <%If Cint(month(rs("sdate")))=9 then%>selected<%End if%>>กันยายน</option>
    <option value="10" <%If Cint(month(rs("sdate")))=10 then%>selected<%End if%>>ตุลาคม</option>
    <option value="11" <%If Cint(month(rs("sdate")))=11 then%>selected<%End if%>>พฤศจิกายน</option>
    <option value="12" <%If Cint(month(rs("sdate")))=12 then%>selected<%End if%>>ธันวาคม</option>
                </select> &nbsp;<select name="yearr" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(153,255,255);">
                <option value="2018"><%
		y=1000 
		Do While y<=9999
		%>

                    </option>
<option value="<%=y%>" <%if y=CInt(Year(rs("sdate")))  then%>selected<%End if%>><%=y%></option>
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
