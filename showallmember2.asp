<%idmember=request("idmember")
sql="SELECT * FROM member  WHERE idmember ='"&idmember&"' order by idmember;"

Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3
%>
<html>

<head>
<title>�ʴ���������Ҫԡ������</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>
<p align="left"><a href="user.asp?idmember=<%=idmember%>"><span style="font-size:18pt;"><b><font face="TH Baijam" color="navy">˹����ѡ</font></b></span></a><span style="font-size:18pt;"><b><font face="TH Baijam" color="navy"> 
</font><a href="showallmember2.asp?idmember=<%=idmember%>"><font face="TH Baijam" color="#DA4453">�ʴ���������Ҫԡ</font></a><font face="TH Baijam" color="navy"> 
</font><a href="editmember2.asp?idmember=<%=idmember%>"><font face="TH Baijam" color="navy">��䢢�������Ҫԡ</font></a></b></span></p>

<FORM METHOD=POST ACTION="" name="frmMain" enctype="multipart/form-data">

    <font color="#003333" face="TH Baijam"><%
sql = "SELECT * FROM  member WHERE idmember='"+request("idmember")+"';"



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

 
    
    </font><p align="center"><font color="#003333" face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="#2BC0E4"><span style="font-size:28pt;"><b>�ʴ�</b></span></font><font face="TH Baijam" color="#614385"><span style="font-size:28pt;"><b>���</b></span></font><font face="TH Baijam" color="#FF8008"><span style="font-size:28pt;"><b>���</b></span></font><font face="TH Baijam" color="#EB3349"><span style="font-size:28pt;"><b>�</b></span></font><font face="TH Baijam" color="#FF512F"><span style="font-size:28pt;"><b>��</b></span></font><font face="TH Baijam" color="#003333"><span style="font-size:28pt;"><b>�ԡ</b></span></font><font face="TH Baijam" color="#FFC837"><span style="font-size:28pt;"><b>���</b></span></font><font face="TH Baijam" color="#1D976C"><span style="font-size:28pt;"><b>���</b></span></font></p>


    <table align="center" width="982" cellpadding="0" cellspacing="0">
        <tr>
            <td width="337" height="41" bgcolor="#FFCCCC">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>������Ҫԡ</b></span></font></p>
            </td>
            <td width="58" height="41" bgcolor="#FFCCCC">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="408" height="41" bgcolor="#FFCCCC">
                <p><font color="#330066" face="TH Baijam"><b>&nbsp;</b><span style="font-size:16pt;"><b><%=rs("idmember")%></b></span></font></p>
            </td>
            <td width="179" height="123" rowspan="3">
<p align="right"><font color="#330066" face="TH Baijam"><b><img src="showpicprofile.asp" style="border : solid #6BA7C4 2px;"></b></font>
					</p>
				
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>����-���ʡ��</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="408" height="41">
                <p><font color="#330066" face="TH Baijam"><b>&nbsp;</b><span style="font-size:16pt;"><b><%=rs("pname")%> 
            <%=rs("fname")%> &nbsp;<%=rs("lname")%></b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>����</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="408" height="41">
                <p><font color="#330066" face="TH Baijam"><b>&nbsp;</b><span style="font-size:16pt;"><b><%=rs("age")%> 
                ��</b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>��</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#330066" face="TH Baijam"><b>&nbsp;</b></font><font face="TH Baijam" color="#330066"><span style="font-size:16pt;"><b><%=rs("sex")%></b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>���˹�</b></span></font></p>
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
                <p><font face="TH Baijam" color="#330066"><span style="font-size:16pt;"><b>&nbsp;</b></span></font><font color="#330066" face="TH Baijam"><span style="font-size:16pt;"><b><%=rs("positionname")%>&nbsp;[<%=rs("position")%>]</b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>˹��§ҹ/����Сͺ���</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#330066" face="TH Baijam"><b>&nbsp;</b><span style="font-size:16pt;"><b><%=rs2("agency")%></b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font color="#000066" face="TH Baijam">&nbsp;</font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#330066" face="TH Baijam"><b>&nbsp;</b></font><font face="TH Baijam" color="#330066"><span style="font-size:16pt;"><b><%=rs2("person")%></b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�������</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#330066" face="TH Baijam"><b>&nbsp;</b></font><font face="TH Baijam" color="#330066"><span style="font-size:16pt;"><b><%=rs2("address")%></b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>������</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#330066" face="TH Baijam"><b>&nbsp;</b><span style="font-size:16pt;"><b><%=rs2("email")%></b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�����Ţ���Ѿ��</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#330066" face="TH Baijam"><b>&nbsp;</b></font><font face="TH Baijam" color="#330066"><span style="font-size:16pt;"><b><%=rs2("phone")%></b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�����</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2">
                <p><font color="#330066" face="TH Baijam"><b>&nbsp;</b><span style="font-size:16pt;"><b><%=rs2("fax")%></b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41" bgcolor="#FFC837">                            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�ѹ�Դ</b></span></font></p>
            </td>
            <td width="58" height="41" bgcolor="#FFC837">
                <p align="center"><font color="#003333" face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="587" height="41" colspan="2" bgcolor="#FFC837">
                <p><font face="TH Baijam" color="#330066"><span style="font-size:16pt;"><b>&nbsp;</b></span></font><font color="#330066" face="TH Baijam"><span style="font-size:16pt;"><b><%=rs2("bdate")%></b></span></font></p>
            </td>
        </tr>
    </table>

    <p align="center">&nbsp;</FORM>
    <p align="center">
	&nbsp;</form>
</body>

</html>
