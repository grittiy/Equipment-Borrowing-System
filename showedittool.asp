<html>

<head>
<title>ระบบฐานข้อมูลเครื่องมือ</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<FORM METHOD=POST ACTION="saveedittool.asp">
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

%>
<INPUT TYPE="hidden" NAME="idtool" value="<%=idtool%>">
<INPUT TYPE="hidden" NAME="toolname" value="<%=toolname%>">
<INPUT TYPE="hidden" NAME="model" value="<%=model%>">

<INPUT TYPE="hidden" NAME="idcategory2" value="<%=idcategory2%>">

<INPUT TYPE="hidden" NAME="size" value="<%=size%>">
<INPUT TYPE="hidden" NAME="weight" value="<%=weight%>">
<INPUT TYPE="hidden" NAME="color" value="<%=color%>">
<INPUT TYPE="hidden" NAME="details" value="<%=details%>">
<INPUT TYPE="hidden" NAME="unitprice" value="<%=unitprice%>">
<INPUT TYPE="hidden" NAME="quantity" value="<%=quantity%>">

<INPUT TYPE="hidden" NAME="dayy"		value="<%=dayy%>">
<INPUT TYPE="hidden" NAME="monthh"		value="<%=monthh%>">
<INPUT TYPE="hidden" NAME="yearr"		value="<%=yearr%>">

<input type="Hidden" name="idate" value="<%=yearr%>/<%=monthh%>/<%=dayy%>">


<p align="center">&nbsp;<font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:28pt;"><b>ระบบฐานข้อมูลเครื่องมือ</b></span></font></p>
    <table align="center" width="531" cellpadding="0" cellspacing="0">
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รหัสเครื่องมือ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=idtool%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ชื่อเครื่องมือ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=toolname%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ชื่อรุ่น</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=model%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>หมวดหมู่เครื่องมือ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
				<%
		sql="SELECT * FROM category2  WHERE idcategory2 ='"&idcategory2&"' order by idcategory2, category2,brand;"

		Set conn =Server.CreateObject("ADODB.Connection")
		conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql,conn,1,3
		%>
            <td width="292" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=rs("category2")%>&nbsp;[</span></font><font color="#CC0000" face="TH Baijam"><span style="font-size:16pt;"><%=rs("brand")%></span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;">]</span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ขนาด</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=size%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>น้ำหนัก</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=weight%> 
                </span></font><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>Kg</b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>สี</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=color%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รายละเอียด</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=details%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ราคาต่อหน่วย</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=unitprice%> 
                </span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>&nbsp;จำนวนในคลัง</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="292" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=quantity%></span></font></p>
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
                <p><font face="TH Baijam" color="#3300CC"><span style="font-size:16pt;"><b>วันที่ 
                </b></span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><% =dayy%>&nbsp;</span></font><font face="TH Baijam" color="#3300CC"><span style="font-size:16pt;"><b>เดือน</b></span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"> 
                 
                <% select case  monthh
						                case "01" 
										          mm ="มกราคม"
										case "02" 
										          mm ="กุมภาพันธ์"
										case "03" 
										          mm ="มีนาคม"
			 							 case "04" 
		 								          mm ="เมษายน"		  		  
										 case "05" 
										          mm ="พฤษภาคม"
									     case "06" 
										          mm ="มิถุนายน"
									     case "07" 
										          mm ="กรกฎาคม"
									     case "08" 
										          mm ="สิงหาคม"
									     case "09" 
										          mm ="กันยายน"
									     case "10" 
										          mm ="ตุลาคม"
								         case "11" 
										          mm ="พฤศจิกายน"
									      case "12" 
										          mm ="ธันวาคม"
					            end select	%> <%=mm%> 
                </span></font><font face="TH Baijam" color="#3300CC"><span style="font-size:16pt;"><b>พ.ศ 
                </b><%=yearr%></span></font><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="blue"><span style="font-size:16pt;">  &nbsp;</span></font></p>
            </td>
        </tr>
    </table>

<p align="center"><font face="TH Sarabun New"><input type="submit" name="บันทึกข้อมูล" value="บันทึกข้อมูล" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);"></font><font face="TH Baijam">&nbsp;</font></p>
</FORM>
</body>

</html>
