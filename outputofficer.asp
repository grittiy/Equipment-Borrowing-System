<html>

<head>
<title>�к��ҹ���������˹�ҷ��</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red">
<FORM METHOD=POST ACTION="saveoffice.asp">
    <font face="TH Baijam"><%
pname=request.Form("pname")
fname=request.Form("fname")
lname=request.Form("lname")
sex=request.Form("sex")
age=request.Form("age")
address=request.Form("address")
phone=request.Form("phone")
email=(request.Form("email"))
password=request.Form("password")

idgenre=request.Form("idgenre")

dayy=CDbl(request.Form("dayy"))
monthh=(request.Form("monthh"))
yearr=CDbl(request.Form("yearr"))
status="admin"

%>
<INPUT TYPE="hidden" NAME="pname" value="<%=pname%>">
<INPUT TYPE="hidden" NAME="fname" value="<%=fname%>">
<INPUT TYPE="hidden" NAME="lname" value="<%=lname%>">
<INPUT TYPE="hidden" NAME="sex" value="<%=sex%>">
<INPUT TYPE="hidden" NAME="age" value="<%=age%>">
<INPUT TYPE="hidden" NAME="address" value="<%=address%>">
<INPUT TYPE="hidden" NAME="phone" value="<%=phone%>">
<INPUT TYPE="hidden" NAME="email" value="<%=email%>">

<INPUT TYPE="hidden" NAME="password" value="<%=password%>">


<INPUT TYPE="hidden" NAME="idgenre" value="<%=idgenre%>">


<INPUT TYPE="hidden" NAME="dayy"		value="<%=dayy%>">
<INPUT TYPE="hidden" NAME="monthh"		value="<%=monthh%>">
<INPUT TYPE="hidden" NAME="yearr"		value="<%=yearr%>">

<input type="Hidden" name="sdate" value="<%=yearr%>/<%=monthh%>/<%=dayy%>">

    </font><p align="center"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:28pt;"><b>�������������˹�ҷ��</b></span></font></p>
    <table align="center" width="633" cellpadding="0" cellspacing="0">
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>����-���ʡ��</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="394" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=pname%><%=fname
%> &nbsp;&nbsp;<%=lname%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>����</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="394" height="41">
<p align="left"><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=age%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>��</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="394" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=sex%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>���˹�</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
			<%
		sql="SELECT * FROM genre  WHERE idgenre ='"&idgenre&"' order by idgenre, genre;"

		Set conn =Server.CreateObject("ADODB.Connection")
		conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql,conn,1,3
		%>
            <td width="394" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=rs("genre")%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�������</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="394" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=address
%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>������</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="394" height="41">
                <p><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=email
%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="41">            <p align="right"><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>���ʼ�ҹ</b></span></font></p>
            </td>
            <td width="58" height="41">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="394" height="41">
                <p>**********************</p>
            </td>
        </tr>
        <tr>
            <td width="181" height="43">
                <p align="right"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�����Ţ���Ѿ��</b></span></font></p>
            </td>
            <td width="58" height="43">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="394" height="43">
                <p><font face="TH Baijam">&nbsp;</font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><%=phone%></span></font></p>
            </td>
        </tr>
        <tr>
            <td width="181" height="43">
                <p align="right"><font face="TH Baijam">&nbsp;</font><font face="TH Baijam" color="navy"><span style="font-size:16pt;"><b>�ѹ�������Ժѵԧҹ</b></span></font></p>
            </td>
            <td width="58" height="43">
                <p align="center"><font face="TH Baijam"><img src="icons8-leaf-fluttering-in-wind-48.png" width="38" height="38" border="0"></font></p>
            </td>
            <td width="394" height="43">
                <p><font face="TH Baijam" color="#3300CC"><span style="font-size:16pt;"><b>�ѹ��� 
                </b></span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"><% =dayy%>&nbsp;</span></font><font face="TH Baijam" color="#3300CC"><span style="font-size:16pt;"><b>��͹</b></span></font><font color="#3300CC" face="TH Baijam"><span style="font-size:16pt;"> 
                 
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
                </span></font><font face="TH Baijam" color="#3300CC"><span style="font-size:16pt;"><b>�.� 
                </b><%=yearr%></span></font><font face="TH Baijam">&nbsp;</font></p>
            </td>
        </tr>
    </table>
<p align="center"><font face="TH Baijam"><input type="submit" name="�ѹ�֡������" value="�ѹ�֡������" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(204,0,0); background-color:rgb(255,102,204);"></font></p>
</FORM>
</body>

</html>
