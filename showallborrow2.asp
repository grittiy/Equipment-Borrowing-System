<html>

<head>
<title>������������������ͧ���</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<FORM METHOD=POST ACTION="main_page.asp" name="frmMain" enctype="multipart/form-data">

    <font color="#003333"><%
sql = "SELECT * FROM  borrow WHERE idborrow='"+request("id")+"';"



Set conn = Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql,conn,1,3

Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.Open sql,conn,1,3

session("idborrow")=rs("idborrow")
%><INPUT TYPE="hidden" NAME="idborrow"  value="<%=rs("idborrow")%>">

 
   
</font><p align="center">&nbsp;<font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:28pt;"><b>��������������ͧ��ͷ�����</b></span></font></p>
    <table align="center" width="637" cellpadding="0" cellspacing="0">
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>������������ͧ���</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="borrow/icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p> &nbsp;<font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("idborrow")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��Ҫԡ</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="borrow/icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
			<%
			idmember= rs("idmember")

				sql4="SELECT * FROM member  WHERE idmember ='"&idmember&"' order by idmember;"

				Set conn =Server.CreateObject("ADODB.Connection")
				conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

				Set rs4 = Server.CreateObject("ADODB.Recordset")
				rs4.Open sql4,conn,1,3
	
				%>
            <td width="366" height="41">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs4("pname")%> 
            <%=rs4("fname")%> &nbsp;<%=rs4("lname")%> (����<%=rs4("age")%>) 
                <%=rs4("agency")%> �����[<%=rs4("fax")%>]</span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>����ͧ���</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="borrow/icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
			<%
				idtool= rs("idtool")


				sql5="SELECT * FROM tool  WHERE idtool ='"&idtool&"' order by idtool;"

				Set conn5 =Server.CreateObject("ADODB.Connection")
				conn5.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

				Set rs5 = Server.CreateObject("ADODB.Recordset")
				rs5.Open sql5,conn5,1,3
	
				%>
            <td width="366" height="41">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs5("idtool")%> 
            <%=rs5("toolname")%> &nbsp;���<%=rs5("model")%>&nbsp;��[<%=rs5("color")%>]</span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>���˹�ҷ��</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="borrow/icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
				<%
				idmember= rs("idmember")


				sql6="SELECT * FROM member  WHERE idmember ='"&idmember&"' order by idmember;"

				Set conn6 =Server.CreateObject("ADODB.Connection")
				conn6.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

				Set rs6 = Server.CreateObject("ADODB.Recordset")
				rs6.Open sql6,conn6,1,3
	
				%>
            <td width="366" height="41">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs6("pname")%> 
            <%=rs6("fname")%> &nbsp;<%=rs6("lname")%> (����<%=rs6("age")%>) 
                <%=rs6("agency")%> ���Ѿ��[<%=rs6("phone")%>]</span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�ѹ������</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="borrow/icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
            <p><font color="#990033" face="TH Baijam"><span style="font-size:16pt;">&nbsp;<%=Day(rs("bdate"))%> 
		<%monthh=month(rs("bdate"))

select case  monthh
						                case "1" 
										          mm ="���Ҥ�"
										case "2" 
										          mm ="����Ҿѹ��"
										case "3" 
										          mm ="�չҤ�"
			 							 case "4" 
		 								          mm ="����¹"		  		  
										 case "5" 
										          mm ="����Ҥ�"
									     case "6" 
										          mm ="�Զع�¹"
									     case "7" 
										          mm ="�á�Ҥ�"
									     case "8" 
										          mm ="�ԧ�Ҥ�"
									     case "9" 
										          mm ="�ѹ��¹"
									     case "10" 
										          mm ="���Ҥ�"
								         case "11" 
										          mm ="��Ȩԡ�¹"
									      case "12" 
										          mm ="�ѹ�Ҥ�"
					            end select
%>
                &nbsp; 
            <%=mm%> &nbsp; <%=year(rs("bdate"))%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�ӹǹ����ͧ��ͷ�����</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="borrow/icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p> <font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("quantity")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�ѹ���׹</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="borrow/icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
            <p><font face="TH Baijam" color="#990033"><span style="font-size:16pt;">&nbsp;</span></font><font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=Day(rs("edate"))%> 
			<%monthh2=month(rs("edate"))

select case  monthh2
						                case "1" 
										          mm ="���Ҥ�"
										case "2" 
										          mm ="����Ҿѹ��"
										case "3" 
										          mm ="�չҤ�"
			 							 case "4" 
		 								          mm ="����¹"		  		  
										 case "5" 
										          mm ="����Ҥ�"
									     case "6" 
										          mm ="�Զع�¹"
									     case "7" 
										          mm ="�á�Ҥ�"
									     case "8" 
										          mm ="�ԧ�Ҥ�"
									     case "9" 
										          mm ="�ѹ��¹"
									     case "10" 
										          mm ="���Ҥ�"
								         case "11" 
										          mm ="��Ȩԡ�¹"
									      case "12" 
										          mm ="�ѹ�Ҥ�"
					            end select
%>
                &nbsp; 
            <%=mm%> &nbsp; <%=year(rs("edate"))%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�ӹǹ�Թ</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam"><img src="borrow/icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p> <font color="#990033" face="TH Baijam"><span style="font-size:16pt;"><%=rs("amount")%></span></font></p>
            </td>
        </tr>
    </table>
<p align="center"><font face="TH Baijam"><input type="submit" name="��Ѻ�" value="��Ѻ�" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);">&nbsp;&nbsp;</font>&nbsp;</p>
</FORM>
    <div align="right">
        <table cellpadding="0" cellspacing="0" width="154" bordercolordark="black" bordercolorlight="black">
            <tr>
                <td width="154" height="97">
                    <p align="center">&nbsp;<a href="login.asp"><img src="icons8-plus-96.png" width="83" height="83" border="0"></a></p>
                </td>
            </tr>
        </table>
    </div>
    <p align="center">&nbsp;</p>
</body>

</html>
