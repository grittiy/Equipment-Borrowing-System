<html>

<head>
<title>�ʴ����������˹�ҷ�������</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<FORM METHOD=POST ACTION="savepicoffice.asp" name="frmMain" enctype="multipart/form-data">

    <font color="#003333"><%
sql = "SELECT * FROM  office WHERE idoffice='"+request("id")+"';"



Set conn = Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql,conn,1,3

Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.Open sql,conn,1,3

session("idoffice")=rs("idoffice")
status="admin"
%>

<INPUT TYPE="hidden" NAME="idoffice"  value="<%=rs("idoffice")%>">

 
    </font><font face="TH Baijam">&nbsp;</font>
    <p align="center"><font face="TH Baijam" color="navy"><span style="font-size:28pt;"><b>�ʴ����������˹�ҷ�������</b></span></font></p>

    <table align="center" width="734" cellpadding="0" cellspacing="0">
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�������˹�ҷ��</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="351" height="41">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs("idoffice")%></span></font></p>
            </td>
            <td width="144" height="123" rowspan="3">
<p align="right"><font color="#003333"><img src="showpicprofile.asp" style="border : solid #6BA7C4 2px;"></font>
					</p>
				
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>����-���ʡ��</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="351" height="41">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("pname")%> 
            <%=rs("fname")%> &nbsp;<%=rs("lname")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>����</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="351" height="41">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("age")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="495" height="41" colspan="2">
                <p><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><%=rs("sex")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>���˹�</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
			<%
			idgenre= rs1("idgenre")
		sql="SELECT * FROM genre  WHERE idgenre ='"&idgenre&"' order by idgenre, genre;"

		Set conn =Server.CreateObject("ADODB.Connection")
		conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql,conn,1,3
		%>
            <td width="495" height="41" colspan="2">
            <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=rs("genre")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�������</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="495" height="41" colspan="2">
                <p><font face="TH Baijam" color="#003333"><span style="font-size:16pt;"><%=rs2("address")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>������</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="495" height="41" colspan="2">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs2("email")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="43">
                <p align="right"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�����Ţ���Ѿ��</b></span></font></p>
            </td>
            <td width="58" height="43">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="495" height="43" colspan="2">
                <p><font face="TH Baijam">&nbsp;</font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs2("phone")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="43">
                <p align="right"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�ѹ�������Ժѵԧҹ</b></span></font></p>
            </td>
            <td width="58" height="43">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="495" height="43" colspan="2">
                <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs2("sdate")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="43">
                <p align="right"><font face="TH Baijam">&nbsp;</font><font face="TH Sarabun New" color="#003333"><span style="font-size:16pt;"><b>�ٻ�Ҿ</b></span></font></p>
            </td>
            <td width="58" height="43">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="495" height="43" colspan="2">
                <p><font face="TH KoHo" color="#003333"><span style="font-size:18pt;"><input type="file" name="pict" maxlength="50" size="20" style="font-family:'TH Mali Grade 6'; font-weight:bolder; font-size:16pt; color:rgb(0,0,153); background-color:rgb(153,255,255); border-color:maroon; border-style:none;"></span></font></p>
            </td>
        </tr>
    </table>


    <p align="center"><span style="font-size:16pt;"><font color="#003333"><input type="submit" name="��ŧ" value="���/����¹�ٻ�Ҿ" style="font-family:'TH Mali Grade 6'; font-size:16; color:black; background-color:rgb(153,255,255);"></font></span></FORM>
    <p align="center"><font color="#003333"><input type="submit" name="�ʴ�����������" value="�ʴ�����������" style="color:white; background-color:green;"></font>
	</body>

</html>
