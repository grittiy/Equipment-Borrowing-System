<html>

<head>
<title>���Ң���������ͧ���</title>
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
<p align="left"><a href="menuborrow2565.asp"><span style="font-size:18pt;"><b><font face="TH Baijam" color="navy">˹����ѡ</font></b></span></a><span style="font-size:18pt;"><b><font face="TH Baijam" color="navy"> 
</font><a href="inputmember.asp"><font face="TH Baijam" color="navy">��������������ͧ���</font></a><font face="TH Baijam" color="navy"> 
</font><a href="searchmember.asp"><font face="TH Baijam" color="#DA4453">���Ң���������ͧ���</font></a></b></span></p>
<form method="post" action="searchtool.asp">
    <p align="center"><font face="TH Baijam" color="#CC0000"><span style="font-size:28pt;"><b><u>���Ң���������ͧ���</u></b></span></font></p>
<table align="center" width="467" cellpadding="0" cellspacing="0">
    <tr>
        <td width="175" height="47">            <p align="right"><font face="TH Baijam" color="purple"><span style="font-size:16pt;"><b>���Ң�����</b></span></font></p>
        </td>
        <td width="50" height="47">                                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="220" height="47">
            <p align="left"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="black"><span style="font-size:16pt;"><input  type="text" name="searchtext" maxlength="50" size="25" style="font-family:'TH Mali Grade 6'; font-size:20; color:rgb(0,0,153); background-color:rgb(255,204,0); border-style:outset;"></span></font></p>
        </td>
    </tr>
    <tr>
        <td width="175" height="53">            <p align="right"><font face="TH Baijam" color="purple"><span style="font-size:16pt;"><b>���͡���Ң����ŷ���ͧ���</b></span></font></p>
        </td>
        <td width="50" height="53">                                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="220" height="53">
            <p align="left"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="black"><span style="font-size:16pt;"><select name="searchtype" size="1" type="text" style="font-family:'TH Mali Grade 6'; font-size:20; color:rgb(0,0,153); background-color:rgb(255,204,0); border-style:outset;">
                <option value="9">--�ô���͡--</option>
                <option value="1">��������ͧ���</option>
                <option value="2">��������ͧ���</option>
                <option value="3">���</option>
            </select></span></font></p>
        </td>
    </tr>
</table>
<p align="center"><font face="TH Baijam">&nbsp;<input type="submit" name="���Ң�����" value="���Ң�����" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:20; color:rgb(102,0,0); text-align:center; background-color:rgb(255,102,204); border-top-color:rgb(0,0,0); border-right-color:rgb(0,0,0); border-bottom-color:rgb(0,0,0);"></font></p>
</FORM>
<FORM METHOD=POST ACTION="del2tool.asp" name="frmMain" OnSubmit="return onDelete();">

<table align="center" width="1158" cellpadding="0" cellspacing="0">
    <tr bgcolor="#CC00FF">
        <td width="1158" colspan="9">            <p align="left">

<font face="TH Baijam" color="#990033"><span style="font-size:16pt;"><input name="CheckAll" type="checkbox" id="CheckAll" value="Y" onClick="ClickCheckAll(this);"><b>���͡������</b></span></font>
	 
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
	sql="SELECT * FROM tool  WHERE idtool like '%"&searchtext&"%' order by idtool;"
elseif searchtype=2 then
	sql="SELECT * FROM tool  WHERE toolname like '%"&searchtext&"%' order by idtool ;"
elseif searchtype=3 then
	sql="SELECT * FROM tool  WHERE model like '%"&searchtext&"%' order by idtool ;"
elseif searchtype=4 Then
	sql="SELECT * FROM tool  WHERE agency ='"&searchtext&"' order by idtool;"
elseif searchtype=5 Then
	sql2="SELECT * FROM tool  WHERE idcategory2 like '%"&searchtext&"%' order by idtool ;"

rs2.Open sql2,conn,1,3

idcategory2=CInt(rs2("idcategory2"))

	sql="SELECT * FROM tool  WHERE idcategory2 ='"&idcategory2&"' order by idtool;"

elseif searchtype=0 Or searchtype=9 Or searchtext="" Then
	sql="SELECT * FROM tool order by idtool;"

end if

rs.Open sql,conn,1,3

x=1
Do While Not rs.eof 

%>
    <tr bgcolor="#FFCCFF">
        <td width="53" height="49">				<p align="left">
<font face="TH Baijam" color="black"><span style="font-size:16pt;"><INPUT TYPE="checkbox" name="dele"  value="<%=rs("idtool")%>" id="chk"></span></font>
				</p>
        </td>
        <td width="167" height="49">				
                <p align="right"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="purple"><span style="font-size:16pt;"><b>���</b></span></font></p>
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
        <td width="174" height="49">                <p align="right"><font face="TH Baijam" color="black"><span style="font-size:16pt;"><INPUT type="Button" Onclick="location.href='showalltool.asp?id=<%=rs("idtool")%>'"  style="font-family:Tahoma; font-size:12px; border-width:1; border-style:solid; cursor:hand;" value="�ʴ������ŷ�����"></span></font></p>
        </td>
        <td width="73" height="49">                <p align="right"><font face="TH Baijam" color="white"><span style="font-size:18pt;"><INPUT type="Button" Onclick="location.href='edittool.asp?id=<%=rs("idtool")%>'"  style="font-family:Tahoma; font-size:12px; border-width:1; border-style:solid; cursor:hand;" value="���"></span></font></p>
        </td>
        <td width="50" height="49">                <p align="right"><font face="TH Baijam" color="white"><span style="font-size:18pt;"><INPUT type="Button" Onclick="location.href='deltool.asp?id=<%=rs("idtool")%>'"  style="font-family:Tahoma; font-size:12px; border-width:1; border-style:solid; cursor:hand;" value="ź"></span></font></p>
        </td>
    </tr>
    <tr>
        <td width="220" colspan="2" height="47">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��������ͧ���</b></span></font></p>
        </td>
        <td width="77" height="47">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="301" height="47">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs("idtool")%></span></font></p>
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
        <td width="220" colspan="2" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��������ͧ���</b></span></font></p>
        </td>
        <td width="77" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="301" height="45">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs("toolname")%> 
                &nbsp;</span></font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>���</b></span></font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"> 
            <%=rs("model")%></span></font></p>
        </td>
        <td width="207" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��������´</b></span></font></p>
        </td>
        <td width="56" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="297" colspan="3" height="45">
            <p><font face="TH Baijam" color="#990033"><span style="font-size:16pt;">&nbsp;</span></font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("details")%></span></font></p>
        </td>
    </tr>
    <tr>
        <td width="220" colspan="2" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��</b></span></font></p>
        </td>
        <td width="77" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="301" height="45">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs("color")%></span></font></p>
        </td>
        <td width="207" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��Ǵ����</b></span></font></p>
        </td>
        <td width="56" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
			<%
				idcategory2= rs("idcategory2")


				sql4="SELECT * FROM category2  WHERE idcategory2 ='"&idcategory2&"' order by idcategory2;"

				Set conn4 =Server.CreateObject("ADODB.Connection")
				conn4.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

				Set rs4 = Server.CreateObject("ADODB.Recordset")
				rs4.Open sql4,conn4,1,3
	
				%>
        <td width="297" colspan="3" height="45">
            <p><font face="TH Baijam" color="#990033"><span style="font-size:16pt;">&nbsp;</span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=rs4("category2")%>&nbsp;[</span></font><font color="#CC0000" face="TH Baijam"><span style="font-size:16pt;"><%=rs4("brand")%></span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;">]</span></font></p>
        </td>
    </tr>
    <tr>
        <td width="220" colspan="2" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�Ҥҵ��˹���</b></span></font></p>
        </td>
        <td width="77" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="301" height="45">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=rs("unitprice")%> 
                </span></font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�ҷ</b></span></font></p>
        </td>
        <td width="207" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�ӹǹ㹤�ѧ</b></span></font></p>
        </td>
        <td width="56" height="45">                            <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
        </td>
        <td width="297" colspan="3" height="45">
            <p><font face="TH Baijam" color="#990033"><span style="font-size:16pt;">&nbsp;</span></font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("quantity")%></span></font></p>
        </td>
    </tr>
    <tr>
        <td width="220" colspan="2" height="45">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�ѹ���ӡ�úѹ�֡</b></span></font></p>
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


        <p align="center"><font face="TH Baijam"><input type="submit" name="ź������" value="ź������" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:20; color:purple; text-align:center; background-color:rgb(255,204,255); border-top-color:rgb(0,0,0); border-right-color:rgb(0,0,0); border-bottom-color:rgb(0,0,0);">&nbsp;</font></FORM>
</body>

</html>
