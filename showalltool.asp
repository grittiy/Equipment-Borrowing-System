<html>

<head>
<title>�ʴ�����������ͧ��ͷ�����</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<FORM METHOD=POST ACTION="savepictool.asp" name="frmMain" enctype="multipart/form-data">

    <font color="#003333"><%
sql = "SELECT * FROM  tool WHERE idtool='"+request("id")+"';"



Set conn = Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql,conn,1,3

Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.Open sql,conn,1,3

session("idtool")=rs("idtool")

%>

<INPUT TYPE="hidden" NAME="idtool"  value="<%=rs("idtool")%>">

 
    </font><font face="TH Baijam">&nbsp;</font>

<p align="center">&nbsp;<font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:28pt;"><b>�ʴ�����������ͧ��ͷ�����</b></span></font></p>

    <table align="center" width="667" cellpadding="0" cellspacing="0">
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��������ͧ���</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="259" height="41">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("idtool")%></span></font></p>
            </td>
            <td width="169" height="123" rowspan="3">
<p align="right"><font color="#003333"><img src="showpicprofile.asp" style="border : solid #6BA7C4 2px;"></font>
					</p>
				
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��������ͧ���</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="259" height="41">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("toolname")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�������</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="259" height="41">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("model")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��Ǵ��������ͧ���</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
				<%
				idcategory2= rs("idcategory2")


				sql4="SELECT * FROM category2  WHERE idcategory2 ='"&idcategory2&"' order by idcategory2;"

				Set conn4 =Server.CreateObject("ADODB.Connection")
				conn4.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

				Set rs4 = Server.CreateObject("ADODB.Recordset")
				rs4.Open sql4,conn4,1,3
	
				%>
            <td width="428" height="41" colspan="2">
            <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=rs4("category2")%>&nbsp;[</span></font><font color="#CC0000" face="TH Baijam"><span style="font-size:16pt;"><%=rs4("brand")%></span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;">]</span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��Ҵ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="428" height="41" colspan="2">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("size")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>���˹ѡ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="428" height="41" colspan="2">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("weight")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="428" height="41" colspan="2">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("color")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">                            <p align="right"><font face="TH Sarabun New" color="#003333"><span style="font-size:16pt;"><b>�ٻ�Ҿ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="428" height="41" colspan="2">
                <p><font face="TH KoHo" color="#003333"><span style="font-size:18pt;"><input type="file" name="pict" maxlength="50" size="20" style="font-family:'TH Mali Grade 6'; font-weight:bolder; font-size:16pt; color:rgb(0,0,153); background-color:rgb(204,153,255); border-color:maroon; border-style:none;"></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��������´</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="428" height="41" colspan="2">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("details")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�Ҥҵ��˹���</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="428" height="41" colspan="2">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("unitprice")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>&nbsp;�ӹǹ㹤�ѧ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="428" height="41" colspan="2">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("quantity")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="43">
                <p align="right"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�ѹ�����Ҥ�ѧ</b></span></font></p>
            </td>
            <td width="58" height="43">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="428" height="43" colspan="2">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("idate")%></span></font></p>
            </td>
        </tr>
    </table>
    <p align="center"><span style="font-size:16pt;"><font color="#003333"><input type="submit" name="��ŧ" value="���/����¹�ٻ�Ҿ" style="font-family:'TH Mali Grade 6'; font-size:16; color:black; background-color:rgb(204,153,255);"></font></span></FORM>
    <p align="center"><font color="#003333"><input type="submit" name="�ʴ�����������" value="�ʴ�����������" style="color:white; background-color:green;"></font>
	</body>

</html>
