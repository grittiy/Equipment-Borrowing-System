<html>

<head>
<title>ค้นหาข้อมูลสมาชิก</title>
<meta name="generator" content="Namo WebEditor v5.0">
<script language="JavaScript">
	function ClickCheckAll(vol)
	{
		var i=0;
		for(i=0;i<=document.frmMain.chk.length-1;i++)
		{
			if(vol.checked == true)
			{
				document.frmMain.chk[i].checked=true;				
			}
			else
			{
				document.frmMain.chk[i].checked=false;	
			}
		}
	}
</script>
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<p align="left"><a href="menuborrow2565.asp"><span style="font-size:18pt;"><b><font face="TH Baijam" color="navy">หน้าหลัก</font></b></span></a><span style="font-size:18pt;"><b><font face="TH Baijam" color="navy"> 
</font><a href="inputmember.asp"><font face="TH Baijam" color="navy">เพิ่มข้อมูลสมาชิก</font></a><font face="TH Baijam" color="navy"> 
</font><a href="searchmember.asp"><font face="TH Baijam" color="#DA4453">ค้นหาข้อมูลสมาชิก</font></a></b></span></p>
<form method="post" action="searchmember.asp">
    <p align="center"><font face="TH Baijam" color="#CC0000"><span style="font-size:28pt;"><b><u>ค้นหาข้อมูลสมาชิก</u></b></span></font></p>
<table align="center" width="467" cellpadding="0" cellspacing="0">
    <tr>
        <td width="175" height="47">            <p align="right"><font face="TH Baijam" color="purple"><span style="font-size:16pt;"><b>ค้นหาข้อมูล</b></span></font></p>
        </td>
        <td width="50" height="47">                                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="220" height="47">
            <p align="left"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="black"><span style="font-size:16pt;"><input  type="text" name="searchtext" maxlength="50" size="25" style="font-family:'TH Mali Grade 6'; font-size:20; color:rgb(0,0,153); background-color:rgb(255,204,0); border-style:outset;"></span></font></p>
        </td>
    </tr>
    <tr>
        <td width="175" height="53">            <p align="right"><font face="TH Baijam" color="purple"><span style="font-size:16pt;"><b>เลือกค้นหาข้อมูลที่ต้องการ</b></span></font></p>
        </td>
        <td width="50" height="53">                                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="220" height="53">
            <p align="left"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="black"><span style="font-size:16pt;"><select name="searchtype" size="1" type="text" style="font-family:'TH Mali Grade 6'; font-size:20; color:rgb(0,0,153); background-color:rgb(255,204,0); border-style:outset;">
                <option value="9">--โปรดเลือก--</option>
                <option value="1">รหัสสมาชิก</option>
                <option value="2">ชื่อ</option>
                <option value="3">นามสกุล</option>
                <option value="4">หน่วยงาน/ผู้ประกอบการ</option>
            </select></span></font></p>
        </td>
    </tr>
</table>
<p align="center"><font face="TH Baijam">&nbsp;<input type="submit" name="ค้นหาข้อมูล" value="ค้นหาข้อมูล" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:20; color:rgb(102,0,0); text-align:center; background-color:rgb(255,102,204); border-top-color:rgb(0,0,0); border-right-color:rgb(0,0,0); border-bottom-color:rgb(0,0,0);"></font></p>
</FORM>
<FORM METHOD=POST ACTION="del2member.asp" name="frmMain" OnSubmit="return onDelete();">

<table align="center" width="1158" cellpadding="0" cellspacing="0">
    <tr bgcolor="#CC00FF">
        <td width="1158" colspan="9">            <p align="left">

<font face="TH Baijam" color="#990033"><span style="font-size:16pt;"><input name="CheckAll" type="checkbox" id="CheckAll" value="Y" onClick="ClickCheckAll(this);"><b>เลือกทั้งหมด</b></span></font>
	 
  </p>
        </td>
    </tr>
	<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject ("ADODB.Recordset")
Set rs2 = Server.CreateObject ("ADODB.Recordset")



searchtext = request.Form("searchtext")
searchtype = CInt(request.Form("searchtype"))


if searchtype=1 then
	sql2="SELECT * FROM member  WHERE idmember like '%"&searchtext&"%' order by idmember;"
elseif searchtype=2 then
	sql="SELECT * FROM member  WHERE fname like '%"&searchtext&"%' order by idmember ;"
elseif searchtype=3 then
	sql="SELECT * FROM member  WHERE lname like '%"&searchtext&"%' order by idmember ;"
elseif searchtype=4 Then
	sql="SELECT * FROM member  WHERE agency ='"&searchtext&"' order by idmember;"
elseif searchtype=5 Then
	sql2="SELECT * FROM member  WHERE idposition like '%"&searchtext&"%' order by idmember ;"

rs2.Open sql2,conn,1,3

idposition=CInt(rs2("idposition"))

	sql="SELECT * FROM member  WHERE idposition ='"&idposition&"' order by idmember;"

elseif searchtype=0 Or searchtype=9 Or searchtext="" Then
	sql="SELECT * FROM member order by idmember;"

end if

rs.Open sql,conn,1,3

x=1
Do While Not rs.eof 

%>
    <tr bgcolor="#FFCCFF">
        <td width="53" height="49">				<p align="left">
<font face="TH Baijam" color="black"><span style="font-size:16pt;"><INPUT TYPE="checkbox" name="dele"  value="<%=rs("idmember")%>" id="chk"></span></font>
				</p>
        </td>
        <td width="167" height="49">				
                <p align="right"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="purple"><span style="font-size:16pt;"><b>ที่</b></span></font></p>
        </td>
        <td width="77" height="49">				
                <p align="center"><font face="TH Baijam">&nbsp;<img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="301" height="49">				
                <p><font face="TH Baijam">&nbsp;</font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=x%></span></font></p>
        </td>
        <td width="207" height="49">
            <p><font face="TH Baijam">&nbsp;</font></p>
        </td>
        <td width="56" height="49">                <p align="right"><font face="TH Baijam">&nbsp;</font></p>
        </td>
        <td width="174" height="49">                <p align="right"><font face="TH Baijam" color="black"><span style="font-size:16pt;"><INPUT type="Button" Onclick="location.href='showallmember.asp?id=<%=rs("idmember")%>'"  style="font-family:Tahoma; font-size:12px; border-width:1; border-style:solid; cursor:hand;" value="แสดงข้อมูลทั้งหมด"></span></font></p>
        </td>
        <td width="73" height="49">                <p align="right"><font face="TH Baijam" color="white"><span style="font-size:18pt;"><INPUT type="Button" Onclick="location.href='editmember.asp?id=<%=rs("idmember")%>'"  style="font-family:Tahoma; font-size:12px; border-width:1; border-style:solid; cursor:hand;" value="แก้ไข"></span></font></p>
        </td>
        <td width="50" height="49">                <p align="right"><font face="TH Baijam" color="white"><span style="font-size:18pt;"><INPUT type="Button" Onclick="location.href='delmember.asp?id=<%=rs("idmember")%>'"  style="font-family:Tahoma; font-size:12px; border-width:1; border-style:solid; cursor:hand;" value="ลบ"></span></font></p>
        </td>
    </tr>
    <tr>
        <td width="220" colspan="2" height="47">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รหัสสมาชิก</b></span></font></p>
        </td>
        <td width="77" height="47">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="301" height="47">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs("idmember")%></span></font></p>
        </td>
        <td width="207" height="47">                        <p align="right"><font face="TH Baijam">&nbsp;</font></p>
        </td>
        <td width="56" height="47">                            <p align="center"><font face="TH Baijam">&nbsp;</font></p>
        </td>
        <td width="297" colspan="3" height="47">
            <p><font face="TH Baijam" color="#990033"><span style="font-size:16pt;">&nbsp;</span></font></p>
        </td>
    </tr>
    <tr>
        <td width="220" colspan="2" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ชื่อ-นามสกุล</b></span></font></p>
        </td>
        <td width="77" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="301" height="45">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs("pname")%> 
            <%=rs("fname")%> &nbsp;<%=rs("lname")%></span></font></p>
        </td>
        <td width="207" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>อายุ</b></span></font></p>
        </td>
        <td width="56" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="297" colspan="3" height="45">
            <p><font face="TH Baijam" color="#990033"><span style="font-size:16pt;">&nbsp;</span></font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("age")%></span></font></p>
        </td>
    </tr>
    <tr>
        <td width="220" colspan="2" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>อีเมล์</b></span></font></p>
        </td>
        <td width="77" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="301" height="45">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs("email")%></span></font></p>
        </td>
        <td width="207" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>ตำแหน่ง</b></span></font></p>
        </td>
        <td width="56" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
			<%
				idposition= rs("idposition")


				sql4="SELECT * FROM position  WHERE idposition ='"&idposition&"' order by idposition;"

				Set conn4 =Server.CreateObject("ADODB.Connection")
				conn4.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

				Set rs4 = Server.CreateObject("ADODB.Recordset")
				rs4.Open sql4,conn4,1,3
	
				%>
        <td width="297" colspan="3" height="45">
            <p><font face="TH Baijam" color="#990033"><span style="font-size:16pt;">&nbsp;</span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=rs4("positionname")%>&nbsp;[</span></font><font color="#CC0000" face="TH Baijam"><span style="font-size:16pt;"><%=rs4("position")%></span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;">]</span></font></p>
        </td>
    </tr>
    <tr>
        <td width="220" colspan="2" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>หน่วยงาน/ผู้ประกอบการ</b></span></font></p>
        </td>
        <td width="77" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="301" height="45">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs("agency")%></span></font></p>
        </td>
        <td width="207" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>โทรสาร</b></span></font><font face="TH Baijam">&nbsp;</font></p>
        </td>
        <td width="56" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="297" colspan="3" height="45">
            <p><font face="TH Baijam" color="#990033"><span style="font-size:16pt;">&nbsp;</span></font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("fax")%></span></font></p>
        </td>
    </tr>
    <tr>
        <td width="220" colspan="2" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>วันที่ทำการบันทึก</b></span></font></p>
        </td>
        <td width="77" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="301" height="45">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;</span></font><font color="#A43931" face="Angsana New"><span style="font-size:18pt;"><%=formatdateTime(rs("datesave"))%></span></font></p>
        </td>
        <td width="207" height="45">            <p align="right"><font face="TH Baijam">&nbsp;</font></p>
        </td>
        <td width="56" height="45">                            <p align="center">&nbsp;</p>
        </td>
        <td width="297" colspan="3" height="45">
            <p>&nbsp;</p>
        </td>
    </tr>
    <tr bgcolor="#FFCC99">
        <td width="220" colspan="2" height="50">                        <p align="right"><font face="TH Baijam">&nbsp;</font></p>
        </td>
        <td width="77" height="50">            <p align="center"><font face="TH Baijam">&nbsp;</font></p>
        </td>
        <td width="301" height="50">
            <p><font face="TH Baijam">&nbsp;</font></p>
        </td>
        <td width="207" height="50">                        <p align="right"><font face="TH Baijam">&nbsp;</font></p>
        </td>
        <td width="56" height="50">            <p align="center"><font face="TH Baijam">&nbsp;</font></p>
        </td>
        <td width="297" colspan="3" height="50">
            <p><font face="TH Baijam">&nbsp;</font></p>
        </td>
    </tr>
	<%
x=x+1
rs.movenext 
Loop
%>
</table>


        <p align="center"><font face="TH Baijam"><input type="submit" name="ลบข้อมูล" value="ลบข้อมูล" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:20; color:purple; text-align:center; background-color:rgb(255,204,255); border-top-color:rgb(0,0,0); border-right-color:rgb(0,0,0); border-bottom-color:rgb(0,0,0);">&nbsp;</font></FORM>
</body>

</html>
