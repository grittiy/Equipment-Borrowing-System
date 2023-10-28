<html>

<head>
<title>แสดงข้อมูลสมาชิกทั้งหมด</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<FORM METHOD=POST ACTION="savepicmember.asp" name="frmMain" enctype="multipart/form-data">

    <font color="#003333" face="TH Baijam"><%
sql = "SELECT * FROM  member WHERE idmember='"+request("id")+"';"



Set conn = Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql,conn,1,3

Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.Open sql,conn,1,3

session("idmember")=rs("idmember")
status="user"
%>

<INPUT TYPE="hidden" NAME="idmember"  value="<%=rs("idmember")%>">

 
    </font><p align="center"><font color="#003333" face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="#003333"><span style="font-size:28pt;"><b>แสดงข้อมูลสมาชิกทั้งหมด</b></span></font></p>

    <table align="center" width="982" cellpadding="0" cellspacing="0">
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>รหัสสมาชิก</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="408" height="41">
                <p><font color="#003333" face="TH Baijam">&nbsp;</font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("idmember")%></span></font></p>
            </td>
            <td width="179" height="123" rowspan="3">
<p align="right"><font color="#003333" face="TH Baijam"><img src="showpicprofile.asp" style="border : solid #6BA7C4 2px;"></font>
					</p>
				
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>ชื่อ-นามสกุล</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="408" height="41">
                <p><font color="#003333" face="TH Baijam">&nbsp;</font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("pname")%> 
            <%=rs("fname")%> &nbsp;<%=rs("lname")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>อายุ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="408" height="41">
                <p><font color="#003333" face="TH Baijam">&nbsp;</font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("age")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>เพศ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#003333" face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><%=rs("sex")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>ตำแหน่ง</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>

		<%
		idposition= rs1("idposition")
		sql="SELECT * FROM position  WHERE idposition ='"&idposition&"' order by idposition, position,positionname;"

		Set conn =Server.CreateObject("ADODB.Connection")
		conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql,conn,1,3
		%>
            <td width="587" height="41" colspan="2">
                <p><font face="TH Baijam" color="#003333"><span style="font-size:16pt;">&nbsp;</span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=rs("positionname")%>&nbsp;[</span></font><font color="#CC0000" face="TH Baijam"><span style="font-size:16pt;"><%=rs("position")%></span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;">]</span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>หน่วยงาน/ผู้ประกอบการ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#003333" face="TH Baijam">&nbsp;</font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs2("agency")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font color="#003333" face="TH Baijam">&nbsp;</font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#003333" face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><%=rs2("person")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>ที่อยู่</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#003333" face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><%=rs2("address")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>อีเมล์</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#003333" face="TH Baijam">&nbsp;</font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs2("email")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>หมายเลขโทรศัพท์</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#003333" face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><%=rs2("phone")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>โทรสาร</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#003333" face="TH Baijam">&nbsp;</font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs2("fax")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>วันเกิด</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>&nbsp;</b></span></font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs2("bdate")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>รูปภาพ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><b>&nbsp;</b></span><span style="font-size:18pt;"><input type="file" name="pict" maxlength="50" size="20" style="font-family:'TH Mali Grade 6'; font-weight:bolder; font-size:16pt; color:rgb(0,0,153); background-color:rgb(255,204,0); border-color:maroon; border-style:none;"></span></font></p>
            </td>
        </tr>
    </table>

    <p align="center"><span style="font-size:16pt;"><font color="#003333" face="TH Baijam"><input type="submit" name="ตกลง" value="ใส่/เปลี่ยนรูปภาพ" style="font-family:'TH Mali Grade 6'; font-size:16; color:black; background-color:rgb(255,204,51);"></font></span></FORM>
    <p align="center"><font color="#003333" face="TH Baijam"><input type="submit" name="แสดงข้อมูลใหม่" value="แสดงข้อมูลใหม่" style="color:white; background-color:green;"></font><font face="TH Baijam">
	</font></form>
</body>

</html>
