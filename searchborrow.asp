<html>

<head>
<title>ค้นหาข้อมูลใบยืมเครื่องมือ</title>
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
</font><a href="inputborrow.asp"><font face="TH Baijam" color="navy">เพิ่มข้อมูลใบยืมเครื่องมือ</font></a><font face="TH Baijam" color="navy"> 
</font><a href="searchborrow.asp"><font face="TH Baijam" color="#DA4453">ค้นหาข้อมูลใบยืมเครื่องมือ</font></a></b></span></p>
<p align="center">&nbsp;<font face="TH Baijam" color="#CC0000"><span style="font-size:28pt;"><b><u>ค้นหาข้อมูลใบยืมเครื่องมือ</u></b></span></font></p>
<form method="post" action="searchborrow.asp">
<table align="center" width="603" cellpadding="0" cellspacing="0">
    <tr>
        <td width="176" height="47">            <p align="right"><font face="TH Baijam" color="purple"><span style="font-size:16pt;"><b>ชื่อสมาชิก</b></span></font></p>
        </td>
        <td width="66" height="47">                                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="361" height="47">
            <p align="left"><font face="TH Baijam">&nbsp;</font><font face="TH KoHo" color="fuchsia"><span style="font-size:18pt;"><select name="searchtype1" size="1" style="font-family:'TH Baijam'; font-weight:normal; font-size:16pt; color:black; background-color:white; border-color:white; border-style:solid;">
                <option value="999" selected>โปรดเลือก</option>
				<%
				'sql="SELECT * FROM tbsport2021  WHERE id ='"&id&"' order by  id;"
				sql2="SELECT * FROM member order by  fname;"
				
				Set conn =Server.CreateObject("ADODB.Connection")
				conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
				Set rs2 = Server.CreateObject("ADODB.Recordset")
				rs2.Open sql2,conn,1,3

				Do While Not rs2.eof
				%>
				        <option value="<%=rs2("idmember")%>"><%=rs2("pname")%>&nbsp;<%=rs2("fname")%>&nbsp;<%=rs2("lname")%></option>
				<%
				rs2.movenext
				Loop
				%>
            </select></span></font></p>
        </td>
    </tr>
    <tr>
        <td width="176" height="53">            <p align="right"><font face="TH Baijam" color="purple"><span style="font-size:16pt;"><b>เครื่องมือ</b></span></font></p>
        </td>
        <td width="66" height="53">                                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="361" height="53">
            <p align="left"><font face="TH Baijam">&nbsp;</font><font face="TH KoHo" color="fuchsia"><span style="font-size:18pt;"><select name="searchtype2" size="1" style="font-family:'TH Baijam'; font-weight:normal; font-size:16pt; color:black; background-color:white; border-color:white; border-style:none;">
                <option value="999" selected>โปรดเลือก</option>
				<%
				sql3="SELECT * FROM tool order by  toolname;"
				
				Set conn =Server.CreateObject("ADODB.Connection")
				conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
				Set rs3 = Server.CreateObject("ADODB.Recordset")
				rs3.Open sql3,conn,1,3

				Do While Not rs3.eof
				%>
				     <option value="<%=rs3("idtool")%>">&nbsp;<%=rs3("toolname")%>&nbsp;รุ่น<%=rs3("model")%>[<%=rs3("color")%>]</option>
				<%
				rs3.movenext
				Loop
				%>
            </select></span></font></p>
        </td>
    </tr>
</table>
<p align="center"><font face="TH Baijam">&nbsp;<input type="submit" name="ค้นหาข้อมูล" value="ค้นหาข้อมูล" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:20; color:rgb(102,0,0); text-align:center; background-color:rgb(255,102,204); border-top-color:rgb(0,0,0); border-right-color:rgb(0,0,0); border-bottom-color:rgb(0,0,0);"></font></p>
</FORM>
<FORM METHOD=POST ACTION="del2borrow.asp" name="frmMain" OnSubmit="return onDelete();">

<table align="center" width="1158" cellpadding="0" cellspacing="0">
    <tr bgcolor="#CC00FF">
        <td width="1158" colspan="9">            <p align="left">

<font face="TH Baijam" color="#990033"><span style="font-size:16pt;"><input name="CheckAll" type="checkbox" id="CheckAll" value="Y" onClick="ClickCheckAll(this);"><b>เลือกทั้งหมด</b></span></font>
	 
  </p>
        </td>
    </tr>
<%
 searchtype1= request.Form("searchtype1")
 searchtype2= request.Form("searchtype2")

 If searchtype1 <> 999 And searchtype2 = 999  Then
	sql="SELECT * FROM borrow WHERE idmember = '"&searchtype1&"' order by  idmember;"
ElseIf searchtype2 <> 999 And searchtype1 = 999  Then
	sql="SELECT * FROM borrow  WHERE idtool = '"&searchtype2&"' order by idtool;"
Else
	sql="SELECT * FROM borrow order by idmember,idtool;"
End if 
	Set conn =Server.CreateObject("ADODB.Connection")
	conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql,conn,1,3
x=1 
Do While Not rs.eof
%>
    <tr bgcolor="#FFCCFF">
        <td width="53" height="49">				<p align="left">
<font face="TH Baijam" color="black"><span style="font-size:16pt;"><INPUT TYPE="checkbox" name="dele"  value="<%=rs("idborrow")%>" id="chk"></span></font>
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
        <td width="174" height="49">                <p align="right"><font face="TH Baijam" color="black"><span style="font-size:16pt;"><INPUT type="Button" Onclick="location.href='showallborrow.asp?id=<%=rs("idborrow")%>'"  style="font-family:Tahoma; font-size:12px; border-width:1; border-style:solid; cursor:hand;" value="แสดงข้อมูลทั้งหมด"></span></font>&nbsp;</p>
        </td>
        <td width="73" height="49">                <p align="right"><font face="TH Baijam" color="white"><span style="font-size:18pt;"><INPUT type="Button" Onclick="location.href='editborrow.asp?id=<%=rs("idborrow")%>'"  style="font-family:Tahoma; font-size:12px; border-width:1; border-style:solid; cursor:hand;" value="แก้ไข"></span></font></p>
        </td>
        <td width="50" height="49">                <p align="right"><font face="TH Baijam" color="white"><span style="font-size:18pt;"><INPUT type="Button" Onclick="location.href='delborrow.asp?id=<%=rs("idborrow")%>'"  style="font-family:Tahoma; font-size:12px; border-width:1; border-style:solid; cursor:hand;" value="ลบ"></span></font></p>
        </td>
    </tr>
    <tr>
        <td width="220" colspan="2" height="47">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>รหัสใบยืมเครื่องมือ</b></span></font></p>
        </td>
        <td width="77" height="47">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="301" height="47">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs("idborrow")%></span></font></p>
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
        <td width="220" colspan="2" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>สมาชิก</b></span></font></p>
        </td>
        <td width="77" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
		<%
				idmember= rs("idmember")


				sql4="SELECT * FROM member  WHERE idmember ='"&idmember&"' order by idmember;"

				Set conn4 =Server.CreateObject("ADODB.Connection")
				conn4.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

				Set rs4 = Server.CreateObject("ADODB.Recordset")
				rs4.Open sql4,conn4,1,3
	
				%>
        <td width="301" height="45">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs4("pname")%> 
            <%=rs4("fname")%> &nbsp;<%=rs4("lname")%> &nbsp;<%=rs4("agency")%> โทรสาร[<%=rs4("fax")%>]</span></font></p>
        </td>
        <td width="207" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>เครื่องมือ</b></span></font></p>
        </td>
        <td width="56" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
		<%
				idtool= rs("idtool")


				sql5="SELECT * FROM tool  WHERE idtool ='"&idtool&"' order by idtool;"

				Set conn5 =Server.CreateObject("ADODB.Connection")
				conn5.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

				Set rs5 = Server.CreateObject("ADODB.Recordset")
				rs5.Open sql5,conn5,1,3
	
				%>
        <td width="297" colspan="3" height="45">
            <p><font face="TH Baijam" color="#990033"><span style="font-size:16pt;">&nbsp;</span></font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">
            <%=rs5("toolname")%> &nbsp;รุ่น<%=rs5("model")%>&nbsp;สี[<%=rs5("color")%>]</span></font></p>
        </td>
    </tr>
    <tr>
        <td width="220" colspan="2" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>วันที่ยืม</b></span></font></p>
        </td>
        <td width="77" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="301" height="45">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=Day(rs("bdate"))%> 
		<%monthh=month(rs("bdate"))

select case  monthh
						                case "1" 
										          mm ="มกราคม"
										case "2" 
										          mm ="กุมภาพันธ์"
										case "3" 
										          mm ="มีนาคม"
			 							 case "4" 
		 								          mm ="เมษายน"		  		  
										 case "5" 
										          mm ="พฤษภาคม"
									     case "6" 
										          mm ="มิถุนายน"
									     case "7" 
										          mm ="กรกฎาคม"
									     case "8" 
										          mm ="สิงหาคม"
									     case "9" 
										          mm ="กันยายน"
									     case "10" 
										          mm ="ตุลาคม"
								         case "11" 
										          mm ="พฤศจิกายน"
									      case "12" 
										          mm ="ธันวาคม"
					            end select
%>
                &nbsp; 
            <%=mm%> &nbsp; <%=year(rs("bdate"))%></span></font></p>
        </td>
        <td width="207" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>วันที่คืน</b></span></font></p>
        </td>
        <td width="56" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
		
        <td width="297" colspan="3" height="45">
            <p><font face="TH Baijam" color="#990033"><span style="font-size:16pt;">&nbsp;</span></font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=Day(rs("edate"))%> 
			<%monthh2=month(rs("edate"))

select case  monthh2
						                case "1" 
										          mm ="มกราคม"
										case "2" 
										          mm ="กุมภาพันธ์"
										case "3" 
										          mm ="มีนาคม"
			 							 case "4" 
		 								          mm ="เมษายน"		  		  
										 case "5" 
										          mm ="พฤษภาคม"
									     case "6" 
										          mm ="มิถุนายน"
									     case "7" 
										          mm ="กรกฎาคม"
									     case "8" 
										          mm ="สิงหาคม"
									     case "9" 
										          mm ="กันยายน"
									     case "10" 
										          mm ="ตุลาคม"
								         case "11" 
										          mm ="พฤศจิกายน"
									      case "12" 
										          mm ="ธันวาคม"
					            end select
%>
                &nbsp; 
            <%=mm%> &nbsp; <%=year(rs("edate"))%></span></font></p>
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
        <td width="207" height="45">            <p align="right">&nbsp;</p>
        </td>
        <td width="56" height="45">                            <p align="center">&nbsp;</p>
        </td>
		
        <td width="297" colspan="3" height="45">
            <p><font face="TH Baijam" color="#990033"><span style="font-size:16pt;">&nbsp;</span></font></p>
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
