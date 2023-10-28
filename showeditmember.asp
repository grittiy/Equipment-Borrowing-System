<html>

<head>
<title>แก้ไขข้อมูลสมาชิก</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<FORM METHOD=POST ACTION="saveeditmember.asp">
<font face="TH Baijam">&nbsp;</font></FORM>
<form method="post" action="saveeditmember.asp">
    <p>
<font face="TH Baijam"><INPUT TYPE="hidden" NAME="idmember" value="<%=idmember%>">
<INPUT TYPE="hidden" NAME="pname" value="<%=pname%>">
<INPUT TYPE="hidden" NAME="fname" value="<%=fname%>">
<INPUT TYPE="hidden" NAME="lname" value="<%=lname%>">
<INPUT TYPE="hidden" NAME="sex" value="<%=sex%>">
<INPUT TYPE="hidden" NAME="age" value="<%=age%>">

<INPUT TYPE="hidden" NAME="idposition" value="<%=idposition%>">

<INPUT TYPE="hidden" NAME="person" value="<%=person%>">
<INPUT TYPE="hidden" NAME="agency" value="<%=agency%>">
<INPUT TYPE="hidden" NAME="address" value="<%=address%>">
<INPUT TYPE="hidden" NAME="phone" value="<%=phone%>">
<INPUT TYPE="hidden" NAME="fax" value="<%=fax%>">
<INPUT TYPE="hidden" NAME="email" value="<%=email%>">
<INPUT TYPE="hidden" NAME="password" value="<%=password%>">


<INPUT TYPE="hidden" NAME="dayy"		value="<%=dayy%>">
<INPUT TYPE="hidden" NAME="monthh"		value="<%=monthh%>">
<INPUT TYPE="hidden" NAME="yearr"		value="<%=yearr%>">

<input type="Hidden" name="bdate" value="<%=yearr%>/<%=monthh%>/<%=dayy%>">


    </font></p>
</form>
<FORM METHOD=POST ACTION="saveeditmember.asp">
<p align="center"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:28pt;"><b>แก้ไขข้อมูลสมาชิก</b></span></font></p>

    <table align="center" width="982" cellpadding="0" cellspacing="0">
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รหัสสมาชิก</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41">
                <p><font face="TH Baijam">&nbsp;</font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=idmember%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ชื่อ-นามสกุล</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41">
                <p><font face="TH Baijam">&nbsp;</font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=pname%> 
                <%=fname
%> &nbsp; <%=lname%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>อายุ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41">
                <p><font face="TH Baijam">&nbsp;</font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=age%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>เพศ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41">
                <p><font face="TH Baijam">&nbsp;</font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=sex%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ตำแหน่ง</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
		<%
		sql="SELECT * FROM position  WHERE idposition ='"&idposition&"' order by idposition, position,positionname;"

		Set conn =Server.CreateObject("ADODB.Connection")
		conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql,conn,1,3
		%>
            <td width="587" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs("positionname")%>&nbsp;[</span></font><font color="#CC0000" face="TH Baijam"><span style="font-size:16pt;"><%=rs("position")%></span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;">]</span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>หน่วยงาน/ผู้ประกอบการ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41">
                <p><font face="TH Baijam">&nbsp;</font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=agency%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam">&nbsp;</font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41">
                <p><font face="TH Baijam">&nbsp;</font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=person
%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ที่อยู่</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41">
                <p><font face="TH Baijam">&nbsp;</font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=address%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>อีเมล์</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41">
                <p><font face="TH Baijam">&nbsp;</font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=email
%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รหัสผ่าน</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41">
                <p><font face="TH Baijam">&nbsp;*********</font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>หมายเลขโทรศัพท์</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41">
                <p><font face="TH Baijam">&nbsp;</font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=phone%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>โทรสาร</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41">
                <p><font face="TH Baijam">&nbsp;</font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=fax%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>วันเกิด</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41">
                <p><font face="TH Baijam" color="#3300CC"><span style="font-size:16pt;"><b>&nbsp;วันที่ 
                </b></span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=dayy%>&nbsp;</span></font><font face="TH Baijam" color="#3300CC"><span style="font-size:16pt;"><b>เดือน</b></span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"> 
                 
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
                </b><%=yearr%></span></font></p>
            </td>
        </tr>
    </table>
<p align="center"><font face="TH Baijam"><input type="submit" name="บันทึกข้อมูล" value="บันทึกข้อมูล" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);"></font></p>
</FORM>
</body>

</html>
