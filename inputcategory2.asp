<html>

<head>
<title>เพิ่มหมวดหมู่เครื่องมือ</title>
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
<p align="left"><a href="menuborrow2565.asp"><span style="font-size:18pt;"><b><font face="TH Baijam" color="navy">หน้าหลัก</font></b></span></a><span style="font-size:18pt;"><b><font face="TH Baijam" color="navy"> 
&nbsp;</font><a href="intputposition.asp"><font face="TH Baijam" color="navy">เพิ่มข้อมูลประเภทตำแหน่ง</font></a><font face="TH Baijam" color="navy"> 
</font><a href="inputgenre.asp"><font face="TH Baijam" color="navy">เพิ่มข้อมูลประเภทเจ้าหน้าท</font></a><font face="TH Baijam" color="navy">ี่</font><font face="TH Baijam" color="#DA4453">&nbsp;</font><a href="inputcategory2.asp"><font face="TH Baijam" color="#DA4453">เพิ่มหมวดหมู่เครื่องมือ</font></a><font face="TH Baijam" color="#DA4453"> 
</font></b></span></p>
<form method="post" action="savecategory2.asp">
    <p align="center"><font face="TH Baijam" color="maroon"><span style="font-size:28pt;"><b><u>เพิ่มหมวดหมู่เครื่องมือ</u></b></span></font></p>
</form>
<FORM METHOD=POST ACTION="savecategory2.asp">
    <table align="center" cellpadding="0" cellspacing="0" width="522">
        <tr>
            <td width="173" height="53">            <p align="right"><font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><b>&nbsp;หมวดหมู่เครื่องมือ</b></span></font></p>
            </td>
            <td width="88" height="53">
                <p align="center"><img src="Lovepik_com-401708332-playing-cards.png" width="44" height="44" border="0"></p>
            </td>
            <td width="261" height="53">
                <p>&nbsp;<input type="text" name="category2" maxlength="50" size="20" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:navy; background-color:rgb(204,153,255);"></p>
            </td>
        </tr>
        <tr>
            <td width="173" height="53">            <p align="right"><font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><b>ยี่ห้อเครื่องมือ</b></span></font></p>
            </td>
            <td width="88" height="53">
                <p align="center"><img src="Lovepik_com-401708332-playing-cards.png" width="44" height="44" border="0"></p>
            </td>
            <td width="261" height="53">
                <p>&nbsp;<input type="text" name="brand" maxlength="50" size="20" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:navy; background-color:rgb(204,153,255);"></p>
            </td>
        </tr>
    </table>

<p align="center"><input type="submit" name="เพิ่มข้อมูล" value="เพิ่มข้อมูล" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(51,0,153); background-color:rgb(204,153,255);"> 
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="ยกเลิก" value="ยกเลิก" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(102,0,102); background-color:rgb(153,204,255);"></p>
</FORM>

<FORM METHOD=POST ACTION="del2category2.asp" name="frmMain" OnSubmit="return onDelete();">
    <table align="center" width="792" cellpadding="0" cellspacing="0">
        <tr>
            <td width="792" colspan="6"><p align="left">

<font face="SOV_monomon" color="white"><span style="font-size:20pt;"><font face="SOV_Thanamas" color="blue"><span style="font-size:16pt;"><input name="CheckAll" type="checkbox" id="CheckAll" value="Y" onClick="ClickCheckAll(this);" style="font-family:'TH Mali Grade 6'; color:rgb(102,0,102); background-color:rgb(255,204,255);"></span></font></span></font><font face="TH Baijam" color="#330099"><span style="font-size:16pt;"><b>เลือกทั้งหมด</b></span></font>
	 
  </p>        </td>
        </tr>
        <tr>
            <td width="105">
                <p>&nbsp;</p>
            </td>
            <td width="39"><p align="center"><font face="TH Baijam" color="#FF0099"><span style="font-size:16pt;"><b>ที่</b></span></font></p>
            </td>
            <td width="217"><p align="left"><font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><b>หมวดหมู่เครื่องมือ</b></span></font></p>
            </td>
            <td width="291"><p align="left"><font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><b>ยี่ห้อเครื่องมือ</b></span></font></p>
            </td>
            <td width="140" colspan="2"><p align="left">&nbsp;</p>
            </td>
        </tr>
		<%
sql="SELECT * FROM category2 order by brand,category2,idcategory2;"


Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql,conn,1,3

x=1
Do While Not rs.eof 

%>
        <tr>
            <td width="105">
                <p>&nbsp;<font face="Angsana New" color="white"><span style="font-size:16pt;"><INPUT TYPE="checkbox" name="dele"  value="<%=rs("idcategory2")%>" id="chk" style="font-family:'TH Mali Grade 6'; color:rgb(102,0,102); background-color:rgb(255,204,255);"></span></font></p>
            </td>
            <td width="39"><p align="center"><font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><%=x%></span></font></p>
            </td>
            <td width="217">                <p><font face="TH Baijam" color="#FF0099"><span style="font-size:16pt;"><%=rs("category2")%></span></font></p>
            </td>
            <td width="291">                <p><font face="TH Baijam" color="#CC0066"><span style="font-size:16pt;"><%=rs("brand")%></span></font></p>
            </td>
            <td width="71">                <p>&nbsp;<font face="SOV_monomon" color="white"><span style="font-size:18pt;"><INPUT type="Button" Onclick="location.href='editcategory2.asp?id=<%=rs("idcategory2")%>'"  style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bolder; font-size:18px; color:rgb(51,0,0); background-color:rgb(204,153,255); border-width:1; border-style:none; cursor:hand;" value="แก้ไข"></span></font></p>
            </td>
            <td width="69"><p align="center"><INPUT type="Button" Onclick="location.href='delcategory.asp?id=<%=rs("idcategory2")%>'"  style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bolder; font-size:18px; color:rgb(102,0,102); background-color:rgb(255,204,255); border-width:1; border-style:none; cursor:hand;" value="ลบ"></p>
            </td>
        </tr>
		<%
x=x+1
rs.movenext 
Loop
%>
    </table>
    <p align="center">&nbsp;<input type="submit" name="ลบข้อมูล" value="ลบข้อมูล" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:20; color:navy; text-align:center; background-color:rgb(204,153,255); border-top-color:black; border-right-color:black; border-bottom-color:black;"></p>
</form>
</body>

</html>
