<%
'--1. TextToBinary
Function TextToBinary(text)
       for i=1 to len(text)
	           character = mid(text,i,1)
			   TextToBinary = TextToBinary &chrB(Asc(character))
	   next
End Function

'--2. BinaryToText
Function BinaryToText(binary)
       BinaryToText = ""
       for i=1 to LenB(binary)
	           character = midB(binary,i,1)
			   BinaryToText = BinaryToText & chr(AscB(character))
	   next
End Function

'--3. สร้างออบเจ็กต์ Dictionary ตัวแรก ---
Set uploaddata = CreateObject("Scripting.Dictionary")

'--4. เก็บข้อมูลเข้าตัวแปร data 
data = Request.BinaryRead(Request.TotalBytes)
posend = InStrB(1,data,TextToBinary(Chr(13)))
header = MidB(data,1,posend-1)

'--5. หาตำแหน่งของเฮดเดอร์ปิดท้าย
endheader = header&TextToBinary("-")
pos_header = 1
pos_endheader = InStrB(1,data,endheader)

'--6. วนรอบการทำงานจนกว่าจะครบทุกเฮดเดอร์
Do While pos_header <> pos_endheader
      
	  '--7.  สร้างอ็อบเจ็กต์ Dictionary ตัวที่ 2
      Set sub_uploaddata = CreateObject("Scripting.Dictionary")
	  
      '--8. ตัดชื่อของช่องกรอกข้อมูล
	  pos_name = InStrB(pos_header,data,TextToBinary("name="))
	  pos_namebegin = pos_name+6
	  pos_nameend = InStrB(pos_namebegin,data,TextToBinary(Chr(34)))
	  name = BinaryToText(MidB(data,pos_namebegin,pos_nameend-pos_namebegin))
	  
	  '--9. สร้างตัวแปรเพื่อเช็คว่า ข้อมูลที่ได้รับเป็นไฟล์หรือไม่
	  pos_file = InStrB(pos_nameend,data,TextToBinary("filename="))
	  enddata = InStrB(pos_nameend,data,header)
	  if (pos_file<>0) and (pos_file < enddata) then
	  
	        '--10. ถ้าใช่ ก็ให้ตัดชื่อไฟล์
	        pos_filebegin = pos_file + 10
			pos_fileend = InStrB(pos_filebegin,data,TextToBinary(Chr(34)))
			filename = BinaryToText(MidB(data,pos_filebegin,pos_fileend-pos_filebegin))
			
			'--11. เก็บข้อมูลชื่อไฟล์ไว้ในอ็อบเจ็กต์ Dictionary
			sub_uploaddata.Add "filename",filename
			
			'--12. ตัด Content-type
			pos_content = InStrB(pos_fileend,data,TextToBinary("Content-Type:"))
			pos_contentbegin = pos_content + 14
			pos_contentend = InStrB(pos_contentbegin,data,TextToBinary(Chr(13)))
			contenttype = BinaryToText(MidB(data,pos_contentbegin,pos_contentend-pos_contentbegin))
			
			'--13. เก็บข้อมูล Content-Type เอาไว้ในอ็อบเจ็กต์ Dictionary
			sub_uploaddata.Add "contenttype",contenttype
			
			'--14. ตัดข้อมูลในไฟล์
			pos_valuebegin = pos_contentend + 4
			pos_valueend = InStrB(pos_valuebegin,data,header)-2
			value = BinaryToText(MidB(data,pos_valuebegin,pos_valueend-pos_valuebegin))
			
			'--15. เก็บข้อมูลไฟล์เอาไว้ในอ็อบเจ็กต์ Dictionary
			sub_uploaddata.Add "value",value
			
	  '--16. ถ้า ไม่ใช่ ก็ตัดข้อมูลที่กรอก
	  else
	          pos_valuebegin = pos_nameend + 4
			  pos_valueend = InStrB(pos_valuebegin,data,header)-2
			  value = BinaryToText(MidB(data,pos_valuebegin,pos_valueend-pos_valuebegin))
			  
		'--17.  เก็บข้อมูลไฟล์เอาไว้ในอ็อบเจ็กต์ Dictionary
			  sub_uploaddata.Add "value",value
	  end if
       '--18. เก็บข้อมูลทั้งหมดไว้ใน Dictionary ชื่อ uploaddata
	   uploaddata.Add name,sub_uploaddata
	   
	   '--19. เปลี่ยนตำแหน่งเฮดเดอร์ไปยังข้อมูลชุดถัดไป
	   pos_header = InStrB(pos_header+LenB(header),data,header)
Loop
%>