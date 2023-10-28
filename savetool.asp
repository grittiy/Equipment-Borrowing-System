<html>

<head>
<title>ข้อมูลเครื่องมือซ้ำ</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<FORM METHOD=POST ACTION="outputtool.asp">
<%
idtool=request.Form("idtool")
toolname=request.Form("toolname")
model=request.Form("model")

idcategory2=request.Form("idcategory2")

size=request.Form("size")
weight=request.Form("weight")
color=(request.Form("color"))

details=request.Form("details")
unitprice=request.Form("unitprice")
quantity=request.Form("quantity")

dayy=CDbl(request.Form("dayy"))
monthh=(request.Form("monthh"))
yearr=CDbl(request.Form("yearr"))
idate=request.Form("idate")


sql="SELECT * FROM tool  WHERE idtool ='"&idtool&"' order by idtool;"

Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

If rs.eof Then


		rs.addnew
		rs("idtool") =idtool
		rs("toolname") =toolname
		rs("model") =model

		rs("idcategory2") =idcategory2

		rs("size") =size
		rs("weight") =weight
		rs("color") =color
		rs("details") =details
		rs("unitprice") =unitprice
		rs("quantity") =quantity

		rs("idate") =idate
		rs("datesave")=Now()
		rs.update

		response.redirect "inputtool.asp"

Else
%>
<%End if%>


<p align="center">&nbsp;<font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:28pt;"><b>ข้อมูลเครื่องมือซ้ำ</b></span></font></p>
    <table align="center" width="531" cellpadding="0" cellspacing="0">
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รหัสเครื่องมือ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
                <p> <font face="TH Baijam"><input type="text" name="idtool" maxlength="11" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="20" value="<%=idtool%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ชื่อเครื่องมือ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
                <p> <font face="TH Baijam"><input type="text" name="toolname" maxlength="100" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="20" value="<%=toolname%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ชื่อรุ่น</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="model" maxlength="100" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="20" value="<%=model%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>หมวดหมู่เครื่องมือ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
<p><font face="TH Baijam"><span style="font-size:20pt;"><select name="idcategory2" size="1" type="text" style="font-family:'TH Mali Grade 6'; font-size:20; color:maroon; background-color:rgb(204,153,255); border-style:outset;">
		<%
			sql="SELECT * FROM category2 order by idcategory2;"

			Set conn =Server.CreateObject("ADODB.Connection")
			conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql,conn,1,3
				
			Do While Not rs.eof
			
		%>

            <option value="<%=rs("idcategory2")%>"><%=rs("category2")%>&nbsp;[<%=rs("brand")%>]</CENTER>
			</option>

			<%
			rs.movenext
			Loop
			%>


            </select></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ขนาด</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="size" maxlength="30" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="20" value="<%=size%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>น้ำหนัก</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="weight" maxlength="7" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="10" value="<%=weight%>"> 
                </font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>Kg</b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>สี</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="color" maxlength="30" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="20" value="<%=color%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รายละเอียด</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="details" maxlength="225" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="20" value="<%=details%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ราคาต่อหน่วย</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="unitprice" maxlength="10" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="10" value="<%=unitprice%>"> 
                </font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>บาท</b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>&nbsp;จำนวนในคลัง</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
<p align="left"><font face="TH Baijam"><input type="text" name="quantity" maxlength="2" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:maroon; background-color:rgb(204,153,255);" size="5" value="<%=quantity%>"></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="43">
                <p align="right"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>วันที่เข้าคลัง</b></span></font></p>
            </td>
            <td width="58" height="43">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="43">
                <p><font face="SOV_Thanamas" color="blue"><span style="font-size:16pt;"><select name="dayy" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:maroon; background-color:rgb(204,153,255);">
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
                </select> <select name="monthh" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:maroon; background-color:rgb(204,153,255);">
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
                </select> <select name="yearr" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:maroon; background-color:rgb(204,153,255);">
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
<p align="center"><font face="TH Baijam"><input type="submit" name="ตกลง" value="ตกลง" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);"> 
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="ยกเลิก" value="ยกเลิก" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(102,0,102); background-color:rgb(51,204,255);">&nbsp;</font></p>
</FORM>
</body>

</html>
