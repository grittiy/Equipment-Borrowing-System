<html>

<head>
<title>������������������ͧ���</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<FORM METHOD=POST ACTION="saveborrow.asp">
    <font color="#000066"><%
idborrow=request.Form("idborrow")
idmember=CDbl(request.Form("idmember"))
idofficer=CDbl(request.Form("idofficer"))

idtool=request.Form("idtool")

quantity=request.Form("quantity")
amount=request.Form("amount")

dayy=(request.Form("dayy"))
monthh=(request.Form("monthh"))
yearr=(request.Form("yearr"))

dayy2=(request.Form("dayy2"))
monthh2=(request.Form("monthh2"))
yearr2=(request.Form("yearr2"))
%>
<INPUT TYPE="hidden" NAME="idborrow" value="<%=idborrow%>">

<INPUT TYPE="hidden" NAME="idmember" value="<%=idmember%>">
<INPUT TYPE="hidden" NAME="idtool" value="<%=idtool%>">
<INPUT TYPE="hidden" NAME="idofficer" value="<%=idofficer%>">

<INPUT TYPE="hidden" NAME="quantity" value="<%=quantity%>">
<INPUT TYPE="hidden" NAME="amount" value="<%=amount%>">

<INPUT TYPE="hidden" NAME="dayy"		value="<%=dayy%>">
<INPUT TYPE="hidden" NAME="monthh"		value="<%=monthh%>">
<INPUT TYPE="hidden" NAME="yearr"		value="<%=yearr%>">

<input type="Hidden" name="bdate" value="<%=yearr%>/<%=monthh%>/<%=dayy%>">

<INPUT TYPE="hidden" NAME="dayy2"		value="<%=dayy2%>">
<INPUT TYPE="hidden" NAME="monthh2"		value="<%=monthh2%>">
<INPUT TYPE="hidden" NAME="yearr2"		value="<%=yearr2%>">

<input type="Hidden" name="edate" value="<%=yearr2%>/<%=monthh2%>/<%=dayy2%>">


    </font><p align="center"><font color="#000066">&nbsp;</font><font face="TH Baijam" color="#000066">&nbsp;<span style="font-size:28pt;"><b>������������������ͧ���</b></span></font></p>

    <table align="center" width="637" cellpadding="0" cellspacing="0">
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>������������ͧ���</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam" color="#000066"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"><%=idborrow%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>��Ҫԡ</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam" color="#000066"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
			<%
		sql="SELECT * FROM member  WHERE idmember ='"&idmember&"' order by idmember;"

		Set conn =Server.CreateObject("ADODB.Connection")
		conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql,conn,1,3
		%>
            <td width="366" height="41">
                <p><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"><%=rs("pname")%> 
                <%=rs("fname")%>&nbsp;<%=rs("lname")%>&nbsp;</span></font><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>(���� </b></span></font><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"><%=rs("age")%></span></font><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>) 
                </b></span></font><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"><%=rs("agency")%> 
                </span></font><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�����[</b></span></font><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"><%=rs("fax")%></span></font><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>]</b></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>����ͧ���</b></span></font></p>
            </td>
				<%
		sql="SELECT * FROM tool  WHERE idtool ='"&idtool&"' order by idtool;"

		Set conn =Server.CreateObject("ADODB.Connection")
		conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql,conn,1,3
		%>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam" color="#000066"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"><%=rs("idtool")%><%=rs("toolname")%>&nbsp;���<%=rs("model")%> 
                ��<%=rs("color")%> �Ҥҵ��˹���<%=rs("unitprice")%> �ҷ</span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>���˹�ҷ��</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam" color="#000066"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
			<%
			sql="SELECT * FROM office order by idoffice;"

			Set conn =Server.CreateObject("ADODB.Connection")
			conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql,conn,1,3
				
		%>

            <td width="366" height="41">
                <p><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"><%=rs("pname")%><%=rs("fname")%> 
                <%=rs("lname")%> (���� <%=rs("age")%>) �������Ѿ�� [<%=rs("phone")%>]</span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�ѹ������</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam" color="#000066"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�ѹ��� 
                </b></span></font><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"><%=dayy%>&nbsp;</span></font><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>��͹</b></span></font><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"> 
                 
                <% select case  monthh						                
				case "01" 
										          mm ="���Ҥ�"
										case "02" 
										          mm ="����Ҿѹ��"
										case "03" 
										          mm ="�չҤ�"
			 							 case "04" 
		 								          mm ="����¹"		  		  
										 case "05" 
										          mm ="����Ҥ�"
									     case "06" 
										          mm ="�Զع�¹"
									     case "07" 
										          mm ="�á�Ҥ�"
									     case "08" 
										          mm ="�ԧ�Ҥ�"
									     case "09" 
										          mm ="�ѹ��¹"
									     case "10" 
										          mm ="���Ҥ�"
								         case "11" 
										          mm ="��Ȩԡ�¹"
									      case "12" 
										          mm ="�ѹ�Ҥ�"
					            end select	%> <%=mm%> 
                </span></font><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�.� 
                </b><%=yearr%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�ӹǹ����ͧ��ͷ�����</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam" color="#000066"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"><%=quantity%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�ѹ���׹</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam" color="#000066"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�ѹ��� 
                </b></span></font><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"><%=dayy2%>&nbsp;</span></font><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>��͹</b></span></font><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"> 
                 
                <% select case  monthh2						                
				case "01" 
										          mm ="���Ҥ�"
										case "02" 
										          mm ="����Ҿѹ��"
										case "03" 
										          mm ="�չҤ�"
			 							 case "04" 
		 								          mm ="����¹"		  		  
										 case "05" 
										          mm ="����Ҥ�"
									     case "06" 
										          mm ="�Զع�¹"
									     case "07" 
										          mm ="�á�Ҥ�"
									     case "08" 
										          mm ="�ԧ�Ҥ�"
									     case "09" 
										          mm ="�ѹ��¹"
									     case "10" 
										          mm ="���Ҥ�"
								         case "11" 
										          mm ="��Ȩԡ�¹"
									      case "12" 
										          mm ="�ѹ�Ҥ�"
					            end select	%> <%=mm%> 
                </span></font><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�.� 
                </b><%=yearr2%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="214" height="41">            <p align="right"><font face="TH Baijam" color="#000066"><span style="font-size:16pt;"><b>�ӹǹ�Թ</b></span></font></p>
            </td>
            <td width="57" height="41">
                <p align="center"><font face="TH Baijam" color="#000066"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="366" height="41">
                <p><font color="#000066" face="TH Baijam"><span style="font-size:16pt;"><%=amount%></span></font></p>
            </td>
        </tr>
    </table>

<p align="center"><font face="TH Sarabun New" color="#000066"><input type="submit" name="�ѹ�֡������" value="�ѹ�֡������" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);"></font></p>
</FORM>
</body>

</html>
