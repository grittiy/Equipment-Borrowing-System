<%idmember=request("idmember")
sql="SELECT * FROM member  WHERE idmember ='"&idmember&"' order by idmember;"

Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3
%>


<html>

<head>
<title>��䢢�������Ҫԡ</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<p align="left"><a href="user.asp?idmember=<%=idmember%>"><span style="font-size:18pt;"><b><font face="TH Baijam" color="navy">˹����ѡ</font></b></span></a><span style="font-size:18pt;"><b><font face="TH Baijam" color="navy"> 
</font><a href="showallmember2.asp?idmember=<%=idmember%>"><font face="TH Baijam" color="#000097">�ʴ���������Ҫԡ</font></a><font face="TH Baijam" color="navy"> 
</font><a href="editmember2.asp?idmember=<%=idmember%>"><font face="TH Baijam" color="#F28853">��䢢�������Ҫԡ</font></a></b></span></p>

<p align="center">&nbsp;<font face="TH Sarabun New" color="navy"><span style="font-size:28pt;"><b>��䢢�������Ҫԡ</b></span></font></p>
	<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT * FROM  member WHERE idmember='"+request("idmember")+"';" ,conn,1,3
%>
<form method="post" action="showeditmember2.asp?idmember=<%=idmember%>">
    <table align="center" width="982" cellpadding="0" cellspacing="0">
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>������Ҫԡ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH SarabunPSK" color="#6600CC"><span style="font-size:16pt;"><input  type="hidden" name="idmember" maxlength="13" size="15" style="font-family:SOV_Thanamas; font-size:20; color:blue; background-color:silver; border-style:outset;" value='<%=rs("idmember")%>'
		><%=rs("idmember")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>����-���ʡ��</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><select name="pname" size="1" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);">
            <option value="���" <%If rs("pname")="���"then%>selected<%End if%>>���</option>
            <option value="�ҧ" <%If rs("pname")="�ҧ"then%>selected<%End if%>>�ҧ</option>
			<option value="�ҧ���" <%If rs("pname")="�ҧ���"then%>selected<%End if%>>�ҧ���</option>
            </select> &nbsp;<input type="text" name="fname" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value='<%=rs("fname")%>'> 
                &nbsp;&nbsp;<input type="text" name="lname" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value='<%=rs("lname")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>����</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="age" maxlength="2" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="5" value='<%=rs("age")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>��</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><input type="radio" name="sex" value="���" <%If rs("sex")="���"then%>checked<%End if%>> 
            <b>��� &nbsp;&nbsp;&nbsp; 
            <input type="radio" name="sex" value="˭ԧ" <%If rs("sex")="˭ԧ"then%>checked<%End if%>> ˭ԧ</b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>���˹�</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="Angsana New"><span style="font-size:20pt;"><select name="idposition" size="1" type="text" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,51); border-style:outset;">
		<%
			sql1="SELECT * FROM position order by idposition;"

			Set conn =Server.CreateObject("ADODB.Connection")
			conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

			Set rs1 = Server.CreateObject("ADODB.Recordset")
			rs1.Open sql1,conn,1,3
				
			Do While Not rs1.eof
			
		%>

            <option value="<%=rs1("idposition")%>" <%if CInt(rs1("idposition"))=rs("idposition")  then%>selected<%End if%>><%=rs1("positionname")%>&nbsp;[<%=rs1("position")%>]</CENTER>
			</option>

			<%
			rs1.movenext
			Loop
			%>


            </select></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>˹��§ҹ/����Сͺ���</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="agency" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value='<%=rs("agency")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right">&nbsp;</p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><input type="radio" name="person" value="�Ҥ�Ѱ" <%If rs("person")="�Ҥ�Ѱ"then%>checked<%End if%>><b>�Ҥ�Ѱ 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="person" value="�͡��" <%If rs("person")="�͡��"then%>checked<%End if%>>�͡��</b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>�������</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="address" maxlength="225" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value='<%=rs("address")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>������</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="email" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value='<%=rs("email")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>���ʼ�ҹ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="password" name="password" maxlength="50" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value='<%=rs("password")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>�����Ţ���Ѿ��</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="phone" maxlength="10" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value='<%=rs("phone")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>�����</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="TH Sarabun New"><input type="text" name="fax" maxlength="9" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:black; background-color:rgb(255,204,51);" size="20" value='<%=rs("fax")%>'></font></p>
            </td>
        </tr>
        <tr>
            <td width="337" height="41">                            <p align="right"><font face="TH Sarabun New" color="navy"><span style="font-size:16pt;"><b>�ѹ�Դ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></p>
            </td>
            <td width="587" height="41">
                <p>&nbsp;<font face="SOV_Thanamas" color="blue"><span style="font-size:16pt;"><select name="dayy" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,51);">
                <option value="1">
		<%
		x=1 
		Do While x<=31
		%>

                    </option>
<option value="<%=x%>" <%if x=Day(rs("bdate"))then%>selected<%end if%>><%=x%></option>
		<%
		x=x+1
		loop
		%></option>
                </select> <select name="monthh" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,51);">
                 <option value="01" <%If Cint(month(rs("bdate")))=1 then%>selected<%End if%>>���Ҥ�</option>
    <option value="02" <%If Cint(month(rs("bdate")))=2 then%>selected<%End if%>>����Ҿѹ��</option>
    <option value="03" <%If Cint(month(rs("bdate")))=3 then%>selected<%End if%>>�չҤ�</option>
    <option value="04" <%If Cint(month(rs("bdate")))=4 then%>selected<%End if%>>����¹</option>
    <option value="05" <%If Cint(month(rs("bdate")))=5 then%>selected<%End if%>>����Ҥ�</option>
    <option value="06" <%If Cint(month(rs("bdate")))=6 then%>selected<%End if%>>�Զع�¹</option>
    <option value="07" <%If Cint(month(rs("bdate")))=7 then%>selected<%End if%>>�á�Ҥ�</option>
    <option value="08" <%If Cint(month(rs("bdate")))=8 then%>selected<%End if%>>�ԧ�Ҥ�</option>
    <option value="09" <%If Cint(month(rs("bdate")))=9 then%>selected<%End if%>>�ѹ��¹</option>
    <option value="10" <%If Cint(month(rs("bdate")))=10 then%>selected<%End if%>>���Ҥ�</option>
    <option value="11" <%If Cint(month(rs("bdate")))=11 then%>selected<%End if%>>��Ȩԡ�¹</option>
    <option value="12" <%If Cint(month(rs("bdate")))=12 then%>selected<%End if%>>�ѹ�Ҥ�</option>
                </select> <select name="yearr" type="text" size="1" style="font-family:'TH Mali Grade 6'; font-size:20; color:black; background-color:rgb(255,204,51);">
                <option value="2018"><%
		y=1000 
		Do While y<=9999
		%>

                    </option>
<option value="<%=y%>" <%if y=CInt(Year(rs("bdate"))) then%>selected<%End if%>><%=y%></option>
		<%
		y=y+1
		loop
		%></option>
                </select></span></font></p>
            </td>
        </tr>
    </table>
<p align="center"><font face="TH Sarabun New"><input type="submit" name="��䢢�����" value="��䢢�����" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);"></font></p>
</FORM>
</body>

</html>
