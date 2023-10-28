<FORM METHOD=POST ACTION="saveeditgenre.asp">

<%

sql = "select *  from  genre WHERE idgenre='"&request("id")&"' ;"
Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")

rs.Open  sql,conn,1,3
%>

<INPUT TYPE="hidden" NAME="idgenre" value="<%=rs("idgenre")%>">



    <p align="center"><font face="TH Baijam" color="#FF0099"><span style="font-size:28pt;"><b><u>แก้ไขข้อมูลประเภทเจ้าหน้าที่</u></b></span></font></p>

    <table align="center" cellpadding="0" cellspacing="0" width="522">
        <tr>
            <td width="173" height="53">            <p align="right"><font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><b>ประเภท</b></span></font></p>
            </td>
            <td width="88" height="53">
                <p align="center"><img src="Lovepik_com-401708332-playing-cards.png" width="44" height="44" border="0"></p>
            </td>
            <td width="261" height="53">
                <p>&nbsp;<input type="text" name="genre" maxlength="50" size="20" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:navy; background-color:rgb(204,153,255);" value='<%=rs("genre")%>'></p>
            </td>
        </tr>
    </table>



<p align="center"><input type="submit" name="แก้ไขข้อมูล" value="แก้ไขข้อมูล" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(102,0,0); text-align:center; background-color:rgb(255,153,0); border-top-color:black; border-right-color:black; border-bottom-color:black;"></p>
</FORM>
