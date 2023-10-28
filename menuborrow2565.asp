<%
id=request("id")
%>
<html>

<head>
<title>ระบบการยืมเครื่องมือ</title>
<meta name="generator" content="Namo WebEditor v5.0">
</head>

<body bgcolor="white" text="black" link="blue" vlink="purple" alink="red" background="window-instrumento-workshop-wallpaper-preview.jpg">
<div align="right">
    <table width="196" cellpadding="0" cellspacing="0">
        <tr>
            <td width="196"><p align="right"><span style="font-size:26pt;"><a href="main_page.asp"><font color="white" face="TH Baijam"><b>ออกจากระบบ</span></b></font></a></p>
            </td>
        </tr>
    </table>
</div>
<p align="center"><span style="font-size:16pt;"><span style="font-size:16pt;">&nbsp;</span></span></p>
<p align="center"><span style="font-size:16pt;"><span style="font-size:48pt;"><font color="white" face="TH Baijam"><b>&nbsp;ระบบการยืมเครื่องมือ 
</span></b></font></span></p>
<table align="center" width="1080" cellspacing="0" bordercolordark="black" bordercolorlight="black" bgcolor="white" cellpadding="0" style="text-align:center;" background="marble-bg-purple.jpg">
    <tr bgcolor="white">
        <td width="1074" colspan="8" height="46">
            <p align="center"><span style="font-size:20pt;"><font face="TH Baijam" color="#CB356B"><b>ตา</b></font><font face="TH Baijam" color="#3A1C71"><b>ราง</b></font><font face="TH Baijam" color="#45A247"><b>หลัก 
            </b></font></span></p>
        </td>
    </tr>
    <tr>
        <td width="262" colspan="2" bgcolor="#FDB99B" height="47">
            <p align="center"><span style="font-size:18pt;"><font face="TH Baijam" color="#004E92"><b>แฟ้ม</b></font><font face="TH Baijam" color="#42275A"><b>ข้อมูล</b></font><font face="TH Baijam" color="#734B6D"><b>สมาชิก</b></font></span></p>
        </td>
        <td width="268" colspan="2" bgcolor="#A1FFCE" height="47">
            <p align="center"><span style="font-size:18pt;"><font face="TH Baijam" color="#1D4350"><b>แฟ้มข้อมูล</b></font><font face="TH Baijam" color="#A43931"><b>เจ้าหน้าที่</b></font></span></p>
        </td>
        <td width="252" colspan="2" bgcolor="#F8B500" height="47">
            <p align="center"><span style="font-size:18pt;"><font face="TH Baijam" color="#603813"><b>แฟ้มข้อมูล</b></font><font face="TH Baijam" color="#753A88"><b>เครื่อง</b></font><font face="TH Baijam" color="#4E4376"><b>มือ</b></font></span><span style="font-size:16pt;">&nbsp;</span></p>
        </td>
        <td width="280" colspan="2" bgcolor="#EF473A" height="47">
            <p align="center"><span style="font-size:18pt;"><font face="TH Baijam" color="#F7F8F8"><b>แฟ้มข้อมูล</b></font><font face="TH Baijam" color="#CCCC99"><b>ใบ</b></font><font face="TH Baijam" color="#86FDE8"><b>ยืม</b></font><font face="TH Baijam" color="#FFE000"><b>เครื่อง</b></font><font face="TH Baijam" color="#FFA751"><b>มือ</b></font></span></p>
        </td>
    </tr>
    <tr>
        <td width="129" height="71" bgcolor="#FFFFCC">
            <p align="center"><span style="font-size:16pt;"><a href="inputmember.asp?id=<%=id%>" "><img src="icons8-book-64 (1).png" width="54" height="54" border="0"></a></span></p>
        </td>
        <td width="129" height="71" bgcolor="#FFFFCC">
            <p align="center"><span style="font-size:16pt;"><a href="searchmember.asp?id=<%=id%>"><img src="icons8-search-more-100.png" width="54" height="54" border="0"></a></span></p>
        </td>
        <td width="132" height="71">
            <p align="center"><span style="font-size:16pt;"><a href="inputofficer.asp?id=<%=id%>"><img src="icons8-book-64 (1).png" width="54" height="54" border="0"></a></span></p>
        </td>
        <td width="132" height="71">
            <p align="center"><span style="font-size:16pt;"><a href="searchofficer.asp?id=<%=id%>"><img src="icons8-search-more-100.png" width="54" height="54" border="0"></a></span></p>
        </td>
        <td width="124" height="71" bgcolor="#FFDDE1">
            <p align="center"><span style="font-size:16pt;"><a href="inputtool.asp?id=<%=id%>"><img src="icons8-book-64 (1).png" width="54" height="54" border="0"></a></span></p>
        </td>
        <td width="124" height="71" bgcolor="#FFDDE1">
            <p align="center"><span style="font-size:16pt;"><a href="searchtool.asp?id=<%=id%>"><img src="icons8-search-more-100.png" width="54" height="54" border="0"></a></span></p>
        </td>
        <td width="131" height="71">
            <p align="center"><span style="font-size:16pt;"><a href="inputborrow.asp?id=<%=id%>"><img src="icons8-book-64 (1).png" width="54" height="54" border="0"></a></span></p>
        </td>
        <td width="145" height="71">
            <p align="center"><span style="font-size:16pt;"><a href="searchborrow.asp?id=<%=id%>"><img src="icons8-search-more-100.png" width="54" height="54" border="0"></a></span></p>
        </td>
    </tr>
</table>
<p><span style="font-size:16pt;"><span style="font-size:16pt;">&nbsp;</span></span></p>
<table align="center" width="1045" cellspacing="0" bordercolordark="black" bordercolorlight="black" bgcolor="#EE9CA7" cellpadding="0">
    <tr bgcolor="#EE9CA7">
        <td width="1045" colspan="3" height="49" bgcolor="white" background="marble-bg-purple.jpg">
            <p align="center"><span style="font-size:20pt;"><font face="TH Baijam" color="#654EA3"><b>ตาราง</b></font><font face="TH Baijam" color="#8A2387"><b>ประ</b></font><font face="TH Baijam" color="#1E9600"><b>กอบ</b></font></span></p>
        </td>
    </tr>
    <tr bgcolor="white">
        <td width="320" height="87" bgcolor="#99F2C8">
            <p align="center"><span style="font-size:16pt;"><a href="intputposition.asp"><img src="icons8-quill-with-ink-100.png" width="60" height="60" border="0"></a></span></p>
            <p align="center"><span style="font-size:16pt;"><font face="TH Baijam" color="#93291E"><b>แฟ้มข้อมูลประเภทตำแหน่ง</b></font></span></p>
        </td>
        <td width="385" height="87" bgcolor="#FBD786">
            <p align="center"><span style="font-size:16pt;"><a href="inputgenre.asp?id=<%=id%>"><img src="icons8-quill-with-ink-100.png" width="60" height="60" border="0"></a></span></p>
            <p align="center"><span style="font-size:16pt;"><font face="TH Baijam" color="#0083B0"><b>แฟ้มข้อมูลประเภทเจ้าหน้าที่</b></font></span></p>
        </td>
        <td width="340" height="87" bgcolor="#FFCCCC">
            <p align="center"><span style="font-size:16pt;"><a href="inputcategory2.asp?id=<%=id%>"><img src="icons8-quill-with-ink-100.png" width="60" height="60" border="0"></a></span></p>
            <p align="center"><span style="font-size:16pt;"><font face="TH Baijam" color="#5D26C1"><b>แฟ้มข้อมูลหมวดหมู่เครื่องมือ</b></font></span></p>
        </td>
    </tr>
</table>
<p align="center"><span style="font-size:16pt;"><span style="font-size:16pt;">&nbsp;</span></span></p>
</body>

</html>
