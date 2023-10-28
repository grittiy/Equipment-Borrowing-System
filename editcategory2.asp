<FORM METHOD=POST ACTION="saveeditcategory2.asp">

<%

sql = "select *  from  category2 WHERE idcategory2='"&request("id")&"' ;"
Set conn =Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=borrow2565;UID=root;PWD=;OPTION=3"
Set rs = Server.CreateObject("ADODB.Recordset")

rs.Open  sql,conn,1,3
%>

<INPUT TYPE="hidden" NAME="idcategory2" value="<%=rs("idcategory2")%>">


    <p align="center"><font face="TH Baijam" color="maroon"><span style="font-size:28pt;"><b><u>แก้ไขหมวดหมู่เครื่องมือ</u></b></span></font></p>

    <table align="center" cellpadding="0" cellspacing="0" width="522">
        <tr>
            <td width="173" height="53">            <p align="right"><font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><b>&nbsp;หมวดหมู่เครื่องมือ</b></span></font></p>
            </td>
            <td width="88" height="53">
                <p align="center"><img src="Lovepik_com-401708332-playing-cards.png" width="44" height="44" border="0"></p>
            </td>
            <td width="261" height="53">
                <p>&nbsp;<input type="text" name="category2" maxlength="50" size="20" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:navy; background-color:rgb(204,153,255);" value='<%=rs("category2")%>'></p>
            </td>
        </tr>
        <tr>
            <td width="173" height="53">            <p align="right"><font face="TH Baijam" color="#000099"><span style="font-size:16pt;"><b>ยี่ห้อเครื่องมือ</b></span></font></p>
            </td>
            <td width="88" height="53">
                <p align="center"><img src="Lovepik_com-401708332-playing-cards.png" width="44" height="44" border="0"></p>
            </td>
            <td width="261" height="53">
                <p>&nbsp;<input type="text" name="brand" maxlength="50" size="20" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16pt; color:navy; background-color:rgb(204,153,255);" value='<%=rs("brand")%>'></p>
            </td>
        </tr>
    </table>

<p align="center"><input type="submit" name="แก้ไขข้อมูล" value="แก้ไขข้อมูล" style="font-family:'TH Mali Grade 6'; font-style:normal; font-weight:bold; font-size:16; color:rgb(51,0,153); background-color:rgb(204,153,255);"></p>
</FORM>

