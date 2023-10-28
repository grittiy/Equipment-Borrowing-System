<FORM METHOD=POST ACTION="outputmember.asp">
<%
pname=request.Form("pname")
fname=request.Form("fname")
lname=request.Form("lname")
sex=request.Form("sex")
age=request.Form("age")

idposition=request.Form("idposition")

person=(request.Form("person"))
agency=request.Form("agency")
address=request.Form("address")
phone=request.Form("phone")
fax=(request.Form("fax"))
email=(request.Form("email"))
password=request.Form("password")

dayy=CDbl(request.Form("dayy"))
monthh=(request.Form("monthh"))
yearr=CDbl(request.Form("yearr"))
bdate = request.Form("bdate")
status="user"

sqlmember ="SELECT * FROM member;"
Set rsmember = Server.CreateObject("ADODB.Recordset")
Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

rsmember.Open sqlmember,conn,1,3
If rsmember.eof Then
	rsmember.addnew
		rsmember("idmember") =idmember
		rsmember("pname") =pname
		rsmember("fname") =fname
		rsmember("lname") =lname
		rsmember("sex") =sex
		rsmember("age") =age
		rsmember("status") =status


		rsmember("idposition") =idposition

		rsmember("person") =person
		rsmember("agency") =agency
		rsmember("address") =address
		rsmember("phone") =phone
		rsmember("fax") =fax
		rsmember("email") =email
		rsmember("password") =password
		rsmember("bdate") =bdate
		rsmember("datesave")=Now()
		rsmember.update
		response.redirect "inputmember.asp"
		Else 

sql="SELECT * FROM member  WHERE fname ='"&fname&"' AND  lname ='"&lname&"';"





Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3


If rs.eof Then


		rs.addnew

		rs("pname") =pname
		rs("fname") =fname
		rs("lname") =lname
		rs("sex") =sex
		rs("age") =age
		rs("status") =status


		rs("idposition") =idposition

		rs("person") =person
		rs("agency") =agency
		rs("address") =address
		rs("phone") =phone
		rs("fax") =fax
		rs("email") =email
		rs("password") =password
		rs("bdate") =bdate
		rs("datesave")=Now()
		rs.update
		response.redirect "inputmember.asp"

Else
%>

<p align="center">&nbsp;<font face="TH Sarabun New" color="navy"><span style="font-size:28pt;"><b>ข้อมูลสมาชิกซ้ำ</b></span></font></p>
</FORM>
<form method="post" action="outputmember.asp">
    <table align="center" width="982" cellpadding="0" cellspacing="0">
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>ชื่อ-นามสกุล</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><select name="pname" size="1" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);">
             <option value="นาย"  <%If pname="นาย"then%>selected<%End if%>>นาย</option>
            <option value="นาง" <%If pname="นาง"then%>selected<%End if%>>นาง</option>
			<option value="นางสาว" <%If pname="นางสาว"then%>selected<%End if%>>นางสาว</option>
            </select> &nbsp;<input type="text" name="fname" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value="<%=fname%>"> 
                &nbsp;&nbsp;<input type="text" name="lname" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value="<%=lname%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>อายุ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="age" maxlength="2" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="5" value="<%=age%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>เพศ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><input type="radio" name="sex" value="ชาย" <%If sex="ชาย"then%>checked<%End if%>> 
            <b>ชาย &nbsp;&nbsp;&nbsp; 
            <input type="radio" name="sex" value="หญิง" <%If sex="หญิง"then%>checked<%End if%>> หญิง</b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>ตำแหน่ง</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="Angsana New"><span style="font-size:20pt;"><select name="idposition" size="1" type="text" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,0); border-style:outset;">
		<%
			sql="SELECT * FROM position order by position ,positionname;"

			Set conn =Server.CreateObject("ADODB.Connection")
			conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql,conn,1,3
				
			Do While Not rs.eof
			
		%>

            <option value="<%=rs("idposition")%>"
			<%if CInt(rs("idposition"))=idposition  then%>selected<%End if%>><%=rs("position")%>&nbsp;[<%=rs("positionname")%>]
			</option>

			<%
			rs.movenext
			Loop
			%>


            </select></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>หน่วยงาน/ผู้ประกอบการ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="agency" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value="<%=agency%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><input type="radio" name="person" value="ภาครัฐ" <%If person="ภาครัฐ"then%>checked<%End if%>><b>ภาครัฐ 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="person" value="เอกชน" <%If person="เอกชน"then%>checked<%End if%>>เอกชน</b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>ที่อยู่</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="address" maxlength="225" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value="<%=address%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>อีเมล์</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="email" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value="<%=email%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>รหัสผ่าน</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="password" maxlength="20" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value="<%=password%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>หมายเลขโทรศัพท์</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="phone" maxlength="10" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value="<%=phone%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>โทรสาร</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="fax" maxlength="9" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value="<%=fax%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>วันเกิด</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="SOV_Thanamas" color="blue"><span style="font-size:16pt;"><select name="dayy" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,51);">
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
                </select> <select name="monthh" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,51);">
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
                </select> <select name="yearr" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,51);">
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
                </select></span></font></p>
            </td>
        </tr>
    </table>

<p align="center"><font face="TH Baijam"><input type="submit" name="ตกลง" value="ตกลง" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);"> 
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="ยกเลิก" value="ยกเลิก" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(102,0,102); background-color:rgb(51,204,255);">&nbsp;</font></p>
	<%
	End If 
	End If
	%>
</FORM>
