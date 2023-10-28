<html>

<head>
<title>ระบบการยืมเครื่องมือ</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red" background="window-instrumento-workshop-wallpaper-preview_4.jpg">
<div align="right">
    <table border="1" width="139">
        <tr>
            <td width="129">
                <p align="center"><a href="login.asp"><span style="font-size:20pt;"><b><font face="TH Baijam" color="white">เข้าสู่ระบบ</font></b></span></a></p>
            </td>
        </tr>
    </table>
</div>
<p align="center"><span style="font-size:16pt;"><span style="font-size:48pt;"><font color="white" face="TH Baijam"><b>&nbsp;ระบบการยืมเครื่องมือ 
</b></font></span></span></p>
<FORM METHOD=POST ACTION="" name="frmMain" OnSubmit="return onDelete();">
    <table align="center" width="1175" cellpadding="0" cellspacing="0" bgcolor="white">
        <tr>
            <td width="107" height="62" background="BG-Post-Sub.png">
                <p align="right">&nbsp;<font face="TH Baijam" color="purple"><span style="font-size:16pt;"><b>ชื่อสมาชิก</b></span></font></p>
            </td>
            <td width="69" height="62" background="BG-Post-Sub.png">
<p align="center"><img src="icons8-wrapped-gift-48 (1).png" width="35" height="35" border="0"></p>
            </td>
            <td width="318" colspan="2" height="62" background="BG-Post-Sub.png">
            <p align="left"><font face="TH KoHo" color="fuchsia"><span style="font-size:18pt;"><select name="searchtype1" size="1" style="font-family:'TH Mali Grade 6'; font-weight:normal; font-size:16pt; color:black; background-color:rgb(134,253,232); border-color:white; border-style:none;">
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
            <td width="153" height="62" background="BG-Post-Sub.png">
<p align="right"><font face="TH Baijam" color="purple"><span style="font-size:16pt;"><b>ชื่อเครื่องมือ</b></span></font></p>
            </td>
            <td width="54" height="62" background="BG-Post-Sub.png">
                <p align="center"><img src="icons8-wrapped-gift-48 (1).png" width="35" height="35" border="0"></p>
            </td>
            <td width="318" colspan="2" height="62" background="BG-Post-Sub.png">
            <p align="left"><font face="TH KoHo" color="fuchsia"><span style="font-size:18pt;"><select name="searchtype2" size="1" style="font-family:'TH Mali Grade 6'; font-weight:normal; font-size:16pt; color:black; background-color:rgb(134,253,232); border-color:white; border-style:none;">
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
            <td width="156" height="62" background="pngtree-abstract-background-white-and-gray-geometric-square-with-shadow-image_1386541.jpg">
                <p align="center">&nbsp;<font face="TH Baijam"><input type="submit" name="ค้นหาข้อมูล" value="ค้นหาข้อมูล" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:20; color:rgb(102,0,0); text-align:center; background-color:rgb(234,175,200); border-top-color:black; border-right-color:black; border-bottom-color:black;"></font></p>
            </td>
        </tr>
        <tr bgcolor="#5D26C1">
            <td width="1019" colspan="8" height="41" bgcolor="#EAAFC8">
                <p align="center">&nbsp;</p>
            </td>
            <td width="156" height="41">
                <p>&nbsp;</p>
            </td>
        </tr>
        <tr>
            <td width="107" height="41">
                <p align="center">&nbsp;<font face="TH Baijam" color="#FF0099"><span style="font-size:16pt;"><b>ลำดับที่</b></span></font></p>
            </td>
            <td width="202" height="41" colspan="2"><p align="left"><font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><b>สมาชิก</b></span></font></p>
            </td>
            <td width="338" height="41" colspan="2"><p align="left"><font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><b>เครื่องมือ</b></span></font></p>
            </td>
            <td width="214" height="41" colspan="2"><p align="left"><font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><b>วันที่ยืม</b></span></font></p>
            </td>
            <td width="158" height="41"><p align="left"><font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><b>วันที่คืน</b></span></font></p>
            </td>
            <td width="156" height="41">
                <p>&nbsp;</p>
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
<tr>
            <td width="107" height="41">
                <p align="center">&nbsp;<font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><%=x%></span></font></p>
            </td>
			<%
				idmember= rs("idmember")


				sql4="SELECT * FROM member  WHERE idmember ='"&idmember&"' order by idmember;"

				Set conn4 =Server.CreateObject("ADODB.Connection")
				conn4.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

				Set rs4 = Server.CreateObject("ADODB.Recordset")
				rs4.Open sql4,conn4,1,3
	
				%>
            <td width="202" height="41" colspan="2">                            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs4("pname")%> 
            <%=rs4("fname")%> &nbsp;<%=rs4("lname")%></span></font></p>
            </td>
			<%
				idtool= rs("idtool")


				sql5="SELECT * FROM tool  WHERE idtool ='"&idtool&"' order by idtool;"

				Set conn5 =Server.CreateObject("ADODB.Connection")
				conn5.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

				Set rs5 = Server.CreateObject("ADODB.Recordset")
				rs5.Open sql5,conn5,1,3
	
				%>
            <td width="338" height="41" colspan="2">                            <p>
            <font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs5("toolname")%> &nbsp;รุ่น<%=rs5("model")%>&nbsp;สี[<%=rs5("color")%>]</span></font></p>
            </td>
            <td width="214" height="41" colspan="2">                            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=Day(rs("bdate"))%> 
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
            <td width="158" height="41">                            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=Day(rs("edate"))%> 
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
            <td width="156" height="41">                <p align="center"><font face="TH Baijam" color="black"><span style="font-size:16pt;"><INPUT type="Button" Onclick="location.href='showallborrow2.asp?id=<%=rs("idborrow")%>'"  style="font-family:Tahoma; font-size:12px; border-width:1; border-style:solid; cursor:hand;" value="แสดงข้อมูลทั้งหมด"></span></font></p>
            </td>
        </tr>
		<%
x=x+1
rs.movenext 
Loop
%>
    </table>
</form>
<FORM METHOD=POST ACTION="del2borrow.asp" name="frmMain" OnSubmit="return onDelete();">

        <p align="center">&nbsp;</FORM>
</body>

</html>
