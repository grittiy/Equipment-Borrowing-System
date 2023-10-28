<%idmember=request("idmember")
sql="SELECT * FROM member  WHERE idmember ='"&idmember&"' order by idmember;"

Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

%>

<%
idborrow=request.Form("idborrow")
idmember=CDbl(request.Form("idmember"))
idofficer=CDbl(request.Form("idofficer"))

idtool=request.Form("idtool")

quantity=request.Form("quantity")
amount=request.Form("amount")

dayy=(request.Form("dayy"))
monthh=(request.Form("monthh"))
yearr=(request.Form("yearr"))
bdate=(request.Form("bdate"))


dayy2=(request.Form("dayy2"))
monthh2=(request.Form("monthh2"))
yearr2=(request.Form("yearr2"))
edate=(request.Form("edate"))


sql="SELECT * FROM borrow  WHERE idborrow ='"&idborrow&"' order by idborrow;"


Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

If rs.eof Then


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

		response.redirect ("user.asp?idmember="&idmember&"")

Else
%>

<html>

<head>
<title>เพิ่มข้อมูลใบยืมเครื่องมือ</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<p align="center">&nbsp;<font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:28pt;"><b>ข้อมูลใบยืมเครื่องมือซ้ำ</b></span></font></p>
<form method="post" action="outputborrow2.asp?idmember=<%=idmember%>">
    <table align="center" width="637" cellpadding="0" cellspacing="0">
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รหัสใบยืมเครื่องมือ</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p> <font face="TH Baijam"><input type="text" name="idborrow" maxlength="9" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,255);" size="10" value="<%=idborrow%>"></font></p>
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
			sql7="SELECT * FROM member order by idmember;"


			Set rs7 = Server.CreateObject("ADODB.Recordset")
			rs7.Open sql7,conn,1,3
				
			Do While Not rs7.eof
			
		%>

            <option value='<%=rs7("idmember")%>' <%if CInt(rs7("idmember"))=idmember  then%>selected<%End if%>><%=rs7("pname")%><%=rs7("fname")%>&nbsp;<%=rs7("lname")%>&nbsp;(อายุ<%=rs7("age")%>)&nbsp;<%=rs7("agency")%>&nbsp;โทรสาร[<%=rs7("fax")%>]</CENTER>
			</option>

			<%
			rs7.movenext
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
			sql1="SELECT * FROM tool order by idtool,toolname,model,color,unitprice;"


			Set rs1 = Server.CreateObject("ADODB.Recordset")
			rs1.Open sql1,conn,1,3
				
			Do While Not rs1.eof
			
		%>

            <option value='<%=rs1("idtool")%>'  <%If (rs1("idtool"))=idtool  then%>selected<%End if%>><%=rs1("toolname")%>&nbsp;รุ่น<%=rs1("model")%>&nbsp;สี<%=rs1("color")%>&nbsp;ราคาต่อหน่วย&nbsp;<%=rs1("unitprice")%>&nbsp;บาท</CENTER>
			</option>

			<%
			rs1.movenext
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
			sql="SELECT * FROM office order by idoffice;"


			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql,conn,1,3
				
			Do While Not rs.eof
			
		%>

            <option value='<%=rs("idoffice")%>'  <%if CInt(rs("idoffice"))=idoffice  then%>selected<%End if%>><%=rs("pname")%><%=rs("fname")%>&nbsp;<%=rs("lname")%>&nbsp;(อายุ<%=rs("age")%>)&nbsp;เบอร์โทร[<%=rs("phone")%>]</CENTER>
			</option>

			<%
			rs.movenext
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
                <p><font face="SOV_Thanamas" color="blue"><span style="font-size:16pt;"><select name="dayy" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255);">
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
                </select> <select name="monthh" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255);">
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
                </select> <select name="yearr" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255);">
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
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>จำนวนเครื่องมือที่ยืม</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="quantity" maxlength="2" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,255);" size="5" value="<%=quantity%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>วันที่คืน</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p><font face="SOV_Thanamas" color="blue"><span style="font-size:16pt;"><select name="dayy2" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255);">
                <option value="1">
		<%
		x=1 
		Do While x<=31
		%>

                    </option>
<option value="<%=x%>" <%if x=CInt(dayy2)then%>selected<%end if%>><%=x%></option>
		<%
		x=x+1
		loop
		%></option>
                </select> <select name="monthh2" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255);">
                <option value="01" <%If Cint(monthh2)=1 then%>selected<%End if%>>มกราคม</option>
    <option value="02" <%If Cint(monthh2)=2 then%>selected<%End if%>>กุมภาพันธ์</option>
    <option value="03" <%If Cint(monthh2)=3 then%>selected<%End if%>>มีนาคม</option>
    <option value="04" <%If Cint(monthh2)=4 then%>selected<%End if%>>เมษายน</option>
    <option value="05" <%If Cint(monthh2)=5 then%>selected<%End if%>>พฤษภาคม</option>
    <option value="06" <%If Cint(monthh2)=6 then%>selected<%End if%>>มิถุนายน</option>
    <option value="07" <%If Cint(monthh2)=7 then%>selected<%End if%>>กรกฏาคม</option>
    <option value="08" <%If Cint(monthh2)=8 then%>selected<%End if%>>สิงหาคม</option>
    <option value="09" <%If Cint(monthh2)=9 then%>selected<%End if%>>กันยายน</option>
    <option value="10" <%If Cint(monthh2)=10 then%>selected<%End if%>>ตุลาคม</option>
    <option value="11" <%If Cint(monthh2)=11 then%>selected<%End if%>>พฤศจิกายน</option>
    <option value="12" <%If Cint(monthh2)=12 then%>selected<%End if%>>ธันวาคม</option>
                </select> <select name="yearr2" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,255);">
                <option value="2018"><%
		y=1000 
		Do While y<=9999
		%>

                    </option>
<option value="<%=y%>" <%if y=CInt(yearr2) then%>selected<%End if%>><%=y%></option>
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
<p align="left"><font face="TH Baijam"><input type="text" name="amount" maxlength="10" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,255);" size="10" value="<%=amount%>"></font></p>
            </td>
        </tr>
    </table>
<p align="center"><font face="TH Baijam"><input type="submit" name="ตกลง" value="ตกลง" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);"> 
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="ยกเลิก" value="ยกเลิก" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(102,0,102); background-color:rgb(51,204,255);">&nbsp;</font>&nbsp;</p>
	<%End if%>
</FORM>
</body>

</html>
