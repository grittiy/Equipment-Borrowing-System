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

'--3. ���ҧ�ͺ�硵� Dictionary ����á ---
Set uploaddata = CreateObject("Scripting.Dictionary")

'--4. �红�������ҵ���� data 
data = Request.BinaryRead(Request.TotalBytes)
posend = InStrB(1,data,TextToBinary(Chr(13)))
header = MidB(data,1,posend-1)

'--5. �ҵ��˹觢ͧ�δ����Դ����
endheader = header&TextToBinary("-")
pos_header = 1
pos_endheader = InStrB(1,data,endheader)

'--6. ǹ�ͺ��÷ӧҹ�����ҨФú�ء�δ����
Do While pos_header <> pos_endheader
      
	  '--7.  ���ҧ��ͺ�硵� Dictionary ��Ƿ�� 2
      Set sub_uploaddata = CreateObject("Scripting.Dictionary")
	  
      '--8. �Ѵ���ͧ͢��ͧ��͡������
	  pos_name = InStrB(pos_header,data,TextToBinary("name="))
	  pos_namebegin = pos_name+6
	  pos_nameend = InStrB(pos_namebegin,data,TextToBinary(Chr(34)))
	  name = BinaryToText(MidB(data,pos_namebegin,pos_nameend-pos_namebegin))
	  
	  '--9. ���ҧ�������������� �����ŷ�����Ѻ������������
	  pos_file = InStrB(pos_nameend,data,TextToBinary("filename="))
	  enddata = InStrB(pos_nameend,data,header)
	  if (pos_file<>0) and (pos_file < enddata) then
	  
	        '--10. ����� �����Ѵ�������
	        pos_filebegin = pos_file + 10
			pos_fileend = InStrB(pos_filebegin,data,TextToBinary(Chr(34)))
			filename = BinaryToText(MidB(data,pos_filebegin,pos_fileend-pos_filebegin))
			
			'--11. �红����Ū������������ͺ�硵� Dictionary
			sub_uploaddata.Add "filename",filename
			
			'--12. �Ѵ Content-type
			pos_content = InStrB(pos_fileend,data,TextToBinary("Content-Type:"))
			pos_contentbegin = pos_content + 14
			pos_contentend = InStrB(pos_contentbegin,data,TextToBinary(Chr(13)))
			contenttype = BinaryToText(MidB(data,pos_contentbegin,pos_contentend-pos_contentbegin))
			
			'--13. �红����� Content-Type ���������ͺ�硵� Dictionary
			sub_uploaddata.Add "contenttype",contenttype
			
			'--14. �Ѵ����������
			pos_valuebegin = pos_contentend + 4
			pos_valueend = InStrB(pos_valuebegin,data,header)-2
			value = BinaryToText(MidB(data,pos_valuebegin,pos_valueend-pos_valuebegin))
			
			'--15. �红�����������������ͺ�硵� Dictionary
			sub_uploaddata.Add "value",value
			
	  '--16. ��� ����� ��Ѵ�����ŷ���͡
	  else
	          pos_valuebegin = pos_nameend + 4
			  pos_valueend = InStrB(pos_valuebegin,data,header)-2
			  value = BinaryToText(MidB(data,pos_valuebegin,pos_valueend-pos_valuebegin))
			  
		'--17.  �红�����������������ͺ�硵� Dictionary
			  sub_uploaddata.Add "value",value
	  end if
       '--18. �红����ŷ��������� Dictionary ���� uploaddata
	   uploaddata.Add name,sub_uploaddata
	   
	   '--19. ����¹���˹��δ������ѧ�����Ūش�Ѵ�
	   pos_header = InStrB(pos_header+LenB(header),data,header)
Loop
%>