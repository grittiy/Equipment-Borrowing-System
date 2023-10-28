
<%
	
idmember=request("idmember")
	sql="SELECT * FROM member  WHERE idmember ='"&idmember&"' order by idmember;"

Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3
%>

<html>

<head>
<title>ระบบการยืมเครื่องมือ</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red" background="window-instrumento-workshop-wallpaper-preview_1.jpg">
<div align="right">
    <table width="302" cellpadding="0" cellspacing="0">
        <tr>
			<td width="117"><p align="center"><span style="font-size:26pt;"><a href="showallmember2.asp?idmember=<%=idmember%>"><font color="white" face="TH Baijam"><b>โปรไฟล์</b></font></a></span></p>
            </td>
            <td width="185"><p align="center"><span style="font-size:26pt;"><a href="main_page.asp"><font color="white" face="TH Baijam"><b>ออกจากระบบ</b></font></a></span></p>
            </td>
        </tr>
    </table>
</div>
<p>&nbsp;</p>
<%
	sql1="SELECT * FROM member  WHERE fname ='"&fname&"'AND lname ='"&lname&"' ;"


Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql1,conn,1,3
%>
<p align="center"><span style="font-size:16pt;"><span style="font-size:48pt;"><font color="white" face="TH Baijam"><b>ยินดีต้อนรับคุณ <%=rs("fname")%> &nbsp;<%=rs("lname")%>
</b></font></span></span></p>
<FORM METHOD=POST ACTION="" name="frmMain" OnSubmit="return onDelete();">

<table align="center" width="527" bgcolor="white" cellpadding="0" cellspacing="0">
    <tr>
	
	
        <td width="259" height="44" bgcolor="#FFCCCC"><p align="center" border-top-color:#000><font face="TH Baijam" color="#DA4453"><span style="font-size:18pt;"><b>การยืมเครื่องมือ</b></span></font></p>
        </td>
    </tr>
	
    <tr>
        <td width="259" height="61">
            <p align="center">&nbsp;<a href="inputborrow2.asp?idmember=<%=idmember%>"><img src="icons8-plus-96.png" width="49" height="49" border="0"></a></p>
        </td>
    </tr>
</table>
<p>&nbsp;</p>
</FORM>
</body>

</html>
