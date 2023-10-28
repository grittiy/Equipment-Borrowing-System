<html>

<head>
<title>ข้อมูลเจ้าหน้าที่ซ้ำ</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<FORM METHOD=POST ACTION="outputofficer.asp">
<%
pname=request.Form("pname")
fname=request.Form("fname")
lname=request.Form("lname")
sex=request.Form("sex")
age=request.Form("age")
phone=request.Form("phone")
email=request.Form("email")
address=request.Form("address")
password=request.Form("password")


idgenre = CInt(request.Form("idgenre"))


dayy=CDbl(request.Form("dayy"))
monthh=CDbl(request.Form("monthh"))
yearr=CDbl(request.Form("yearr"))
sdate = request.Form("sdate")
status="admin"



sql="SELECT * FROM office  WHERE fname ='"&fname&"' AND  lname ='"&lname&"';"

Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

If rs.eof Then


		rs.addnew
		rs("pname") =pname
		rs("fname") =fname
		rs("lname") =lname
		rs("sex") =sex
		rs("age") =age
		rs("address") =address
		rs("phone") =phone
		rs("email") =email
		rs("idgenre") =idgenre
		rs("status") =status
				rs("password") =password


		rs("sdate") =sdate
		rs("datesave")=Now()
		rs.update

		response.redirect "inputofficer.asp"

Else
%>



<p align="center">&nbsp;<font face="TH Sarabun New" color="navy"><span style="font-size:28pt;"><b>ข้อมูลเจ้าหน้าที่ซ้ำ</b></span></font></p>
    <table align="center" width="708" cellpadding="0" cellspacing="0">
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>ชื่อ-นามสกุล</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="469" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><select name="pname" size="1" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);">
            <option value="นาย"  <%If pname="นาย"then%>selected<%End if%>>นาย</option>
            <option value="นาง" <%If pname="นาง"then%>selected<%End if%>>นาง</option>
			<option value="นางสาว" <%If pname="นางสาว"then%>selected<%End if%>>นางสาว</option>
            </select> &nbsp;<input type="text" name="fname" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20" value="<%=fname%>"> 
                &nbsp;&nbsp;<input type="text" name="lname" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20" value="<%=lname%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>อายุ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="469" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="age" maxlength="2" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="5" value="<%=age%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>เพศ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="469" height="41">
                <p>&nbsp;<font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><input type="radio" name="sex" value="ชาย" <%If sex="ชาย"then%>checked<%End if%>> 
            <b>ชาย &nbsp;&nbsp;&nbsp; 
            <input type="radio" name="sex" value="หญิง" <%If sex="หญิง"then%>checked<%End if%>> หญิง</b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>ตำแหน่ง</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="469" height="41">
                <p>&nbsp;<font face="Angsana New"><span style="font-size:20pt;"><select name="idgenre" size="1" type="text" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(153,255,255); border-style:outset;">
		<%
			sql="SELECT * FROM genre order by idgenre;"

			Set conn =Server.CreateObject("ADODB.Connection")
			conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql,conn,1,3
				
			Do While Not rs.eof
			
		%>

            <option value='<%=rs("idgenre")%>'  <%if CInt(rs("idgenre"))=idgenre  then%>selected<%End if%>><%=rs("genre")%></CENTER>
			</option>

			<%
			rs.movenext
			Loop
			%>


            </select></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">                            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>ที่อยู่</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="469" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="address" maxlength="225" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20" value="<%=address%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>อีเมล์</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="469" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="email" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20" value="<%=email%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>รหัสผ่าน</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="469" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="password" name="password" maxlength="20" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20" value="<%=password%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="43">
                <p align="right">&nbsp;<font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>หมายเลขโทรศัพท์</b></span></font></p>
            </td>
            <td width="58" height="43">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="469" height="43">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="phone" maxlength="10" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(153,255,255);" size="20" value="<%=phone%>"></font>&nbsp;</p>
            </td>
        </tr>
        <tr>
            <td width="181" height="43">
                <p align="right">&nbsp;<font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>วันเริ่มปฏิบัติงาน</b></span></font></p>
            </td>
            <td width="58" height="43">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="469" height="43">
                <p>&nbsp;<font face="SOV_Thanamas" color="blue"><span style="font-size:16pt;"><select name="dayy" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(153,255,255);">
                <option value="1">
		<%
		x=1 
		Do While x<=31
		%>

                    </option>
<option value="<%=x%>" <%if x=CInt(dayy)then%>selected<%end if%>><%=x%></option>
		<%
		x=x+1
		loop
		%></option>
                </select> <select name="monthh" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(153,255,255);">
                <option value="01" <%If Cint(monthh)=1 then%>selected<%End if%>>มกราคม</option>
    <option value="02" <%If Cint(monthh)=2 then%>selected<%End if%>>กุมภาพันธ์</option>
    <option value="03" <%If Cint(monthh)=3 then%>selected<%End if%>>มีนาคม</option>
    <option value="04" <%If Cint(monthh)=4 then%>selected<%End if%>>เมษายน</option>
    <option value="05" <%If Cint(monthh)=5 then%>selected<%End if%>>พฤษภาคม</option>
    <option value="06" <%If Cint(monthh)=6 then%>selected<%End if%>>มิถุนายน</option>
    <option value="07" <%If Cint(monthh)=7 then%>selected<%End if%>>กรกฏาคม</option>
    <option value="08" <%If Cint(monthh)=8 then%>selected<%End if%>>สิงหาคม</option>
    <option value="09" <%If Cint(monthh)=9 then%>selected<%End if%>>กันยายน</option>
    <option value="10" <%If Cint(monthh)=10 then%>selected<%End if%>>ตุลาคม</option>
    <option value="11" <%If Cint(monthh)=11 then%>selected<%End if%>>พฤศจิกายน</option>
    <option value="12" <%If Cint(monthh)=12 then%>selected<%End if%>>ธันวาคม</option>
                </select> <select name="yearr" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(153,255,255);">
                <option value="2018"><%
		y=1000 
		Do While y<=9999
		%>

                    </option>
<option value="<%=y%>" <%if y=CInt(yearr) then%>selected<%End if%>><%=y%></option>
		<%
		y=y+1
		loop
		%></option>
                </select></span></font>&nbsp;</p>
            </td>
        </tr>
    </table>
	
<p align="center"><font face="TH Sarabun New"><input type="submit" name="ตกลง" value="ตกลง" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);"> 
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="ยกเลิก" value="ยกเลิก" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(102,0,102); background-color:rgb(51,204,255);"></font>&nbsp;</p>
	<%End if%>

</FORM>
</body>

</html>
