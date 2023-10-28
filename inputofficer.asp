<html>

<head>
<title>ระบบฐานข้อมูลเจ้าหน้าที่</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<p align="left"><a href="menuborrow2565.asp"><span style="font-size:18pt;"><b><font face="TH Baijam" color="navy">หน้าหลัก</font></b></span></a><span style="font-size:18pt;"><b><font face="TH Baijam" color="navy"> 
</font><a href="inputofficer.asp"><font face="TH Baijam" color="#DA4453">เพิ่มข้อมูลเจ้าหน้าที่</font></a><font face="TH Baijam" color="navy"> 
</font><a href="searchofficer.asp"><font face="TH Baijam" color="navy">ค้นหาข้อมูลเจ้าหน้าที่</font></a></b></span><font face="TH Baijam">&nbsp;</font></p>
<p align="center"><font face="TH Baijam" color="navy"><span style="font-size:28pt;"><b>เพิ่มข้อมูลเจ้าหน้าที่</b></span></font></p>
<form method="post" action="outputofficer.asp">
    <table align="center" width="699" cellpadding="0" cellspacing="0">
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ชื่อ-นามสกุล</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
                <p><font face="TH Baijam"><select name="pname" size="1" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);">
            <option value="นาย">นาย</option>
            <option value="นาง">นาง</option>
			<option value="นางสาว">นางสาว</option>
            </select> &nbsp;<input type="text" name="fname" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20"> 
                &nbsp;&nbsp;<input type="text" name="lname" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>อายุ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="age" maxlength="2" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="5"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>เพศ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
                <p><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><input type="radio" name="sex" value="ชาย"> 
            <b>ชาย &nbsp;&nbsp;&nbsp; 
            <input type="radio" name="sex" value="หญิง"> หญิง</b></span></font></p>
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
			sql="SELECT * FROM genre order by idgenre;"

			Set conn =Server.CreateObject("ADODB.Connection")
			conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql,conn,1,3
				
			Do While Not rs.eof
			
		%>

            <option value="<%=rs("idgenre")%>"><%=rs("genre")%></CENTER>
			</option>

			<%
			rs.movenext
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
                <p><font face="TH Baijam"><input type="text" name="address" maxlength="225" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>อีเมล์</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
                <p><font face="TH Baijam"><input type="text" name="email" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รหัสผ่าน</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="460" height="41">
                <p><font face="TH Sarabun New"><input type="password" name="password" maxlength="20" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20"></font></p>
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
                <p><font face="TH Baijam"><input type="text" name="phone" maxlength="10" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20">&nbsp;</font></p>
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
<option value="<%=x%>" <%if x=CInt(Day(Now()))then%>selected<%end if%>><%=x%></option>
		<%
		x=x+1
		loop
		%></option>
                </select> <select name="monthh" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(153,255,255);">
                <option value="01" <%If Cint(Month(now))=1 then%>selected<%End if%>>มกราคม</option>
    <option value="02" <%If Cint(Month(now))=2 then%>selected<%End if%>>กุมภาพันธ์</option>
    <option value="03" <%If Cint(Month(now))=3 then%>selected<%End if%>>มีนาคม</option>
    <option value="04" <%If Cint(Month(now))=4 then%>selected<%End if%>>เมษายน</option>
    <option value="05" <%If Cint(Month(now))=5 then%>selected<%End if%>>พฤษภาคม</option>
    <option value="06" <%If Cint(Month(now))=6 then%>selected<%End if%>>มิถุนายน</option>
    <option value="07" <%If Cint(Month(now))=7 then%>selected<%End if%>>กรกฏาคม</option>
    <option value="08" <%If Cint(Month(now))=8 then%>selected<%End if%>>สิงหาคม</option>
    <option value="09" <%If Cint(Month(now))=9 then%>selected<%End if%>>กันยายน</option>
    <option value="10" <%If Cint(Month(now))=10 then%>selected<%End if%>>ตุลาคม</option>
    <option value="11" <%If Cint(Month(now))=11 then%>selected<%End if%>>พฤศจิกายน</option>
    <option value="12" <%If Cint(Month(now))=12 then%>selected<%End if%>>ธันวาคม</option>
                </select> &nbsp;<select name="yearr" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(153,255,255);">
                <option value="2018"><%
		y=1000 
		Do While y<=9999
		%>

                    </option>
<option value="<%=y%>" <%if y=CInt(Year(Now()))+543 then%>selected<%End if%>><%=y%></option>
		<%
		y=y+1
		loop
		%></option>
                </select></span></font></p>
            </td>
        </tr>
    </table>
<p align="center"><font face="TH Baijam"><input type="submit" name="ตกลง" value="ตกลง" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);"> 
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="ยกเลิก" value="ยกเลิก" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(102,0,102); background-color:rgb(51,204,255);">&nbsp;</font></p>
</FORM>
</body>

</html>
