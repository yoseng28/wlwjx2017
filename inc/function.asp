<%
'**************************************************
'��������index_content
'��  �ã���ҳ���ݵ���
'��  ����imgname ---- Ҫ�����ַ�
'        classname����������
'        titlenum����ʾ����
'        linenum����ʾ����
'        bigimg������ͼƬ��ÿ�б���
'����ֵ����ʾ���õ�����
'**************************************************
Function index_content(imgname,classname,titlenum,linenum,bigimg,ProductPrice2)
if classname<>"" then
	sql="select *  from content where bigclassname='" & classname & "' order by id desc "
else
	sql="select *  from content order by id desc "
end if
	Set rs= Server.CreateObject("ADODB.recordset")
	rs.Open sql,conn,1,3
	if NOT rs.EOF then 
		for i=1 to linenum
			if NOT rs.EOF then
				Response.Write ("<TABLE height=25 cellSpacing=0 cellPadding=0 width=100% align=center border=0>")
				Response.Write ("<TBODY>")
				Response.Write ("<TR>")
				Response.Write ("<TD height=25 width=15 background="&bigimg&"><IMG src="&imgname&"></TD>")
   				Response.Write ("<TD height=25 background="&bigimg&">&nbsp;<a href=view.asp?id="&rs("id")&" target=_blank>"&left(rs("title"),titlenum)&"</a>")
				Response.Write ("</TD>")
    			Response.Write ("</TR>")
  				Response.Write ("</TBODY>")
				Response.Write ("</TABLE>")
		rs.MoveNEXT
		end if
		next
	else Response.Write ("��������.") end if
	rs.close 
	Set rs = Nothing
end function
%>

<%
'list,class
'**************************************************
'��������class_bigclassname
'��  �ã���Ŀ��������
'��  ����bigclass ����������
'        smallclass��С������
'����ֵ������bigclass����smallclassֵ
'**************************************************
Function class_bigclassname(bigclass,smallclass)
bigclass=request("type")
smallclass=request("type2")
bigclass=replace(bigclass,"'","")
smallclass=replace(smallclass,"'","")
	if bigclass="" then '
  		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�</font></p>"
		response.end
	end if
	if len(request("type"))>15 then 
		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�</font></p>"
		response.end
	end if
	if smallclass<>"" and len(request("type2"))>15 then
		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�</font></p>"
		response.end
	end if
end function
%>

<%
'class��Ŀ����
'**************************************************
'��������class_smallclassname
'��  �ã�class.asp�ļ�����Ŀ����
'��  ����bigclass ����������
'        smallclass��С������
'����ֵ������bigclass����smallclassֵ
'**************************************************
Function class_smallclassname(class_imgname,class_bigclass,bigimg) 'class_imgname��Ŀǰͼ�꣬class_bigclass�������ƣ�bigimg����ͼƬ
set rs=server.CreateObject("adodb.recordset")
rs.open "Select * From SmallClass Where BigClassName='" & class_bigclass & "'",conn,1,1
if not(rs.bof and rs.eof) then
do while not rs.eof
    Response.Write ("<TABLE width=100% height=25 border=0 cellPadding=0 cellSpacing=0 id=table32>")
        Response.Write ("<TBODY>")
                        Response.Write ("<TR> ")
                          Response.Write ("<TD background="&bigimg&">&nbsp; <font color=#85eb0><strong>��</strong></font> <a  href=list.asp?type="&class_bigclass&"&type2="&rs("SmallClassName")&"><FONT color=#caf48f>"&rs("SmallClassName")&"</FONT></a></TD>")
                        Response.Write ("</TR>")
                      Response.Write ("</TBODY>")
    Response.Write ("</TABLE>")
rs.movenext
loop
end if
rs.close
set rs=nothing	
end function
%>
<%
Function class_content(imgname1,imgname2,bigclass,bigimg1,bigimg2,bigimg3,linenum) 'imgname1:С������ǰͼ�꣬imgname�����ݱ���ǰͼƬ,bigclass����������,bigimg1������ͼƬ����,bigimg2�����ݱ���,bigimg3����С�������ͼƬ����,linenum��ʾ����
set rs=server.CreateObject("adodb.recordset")
rs.open "Select * From SmallClass Where BigClassName='" & bigclass & "'",conn,1,1
if not(rs.bof and rs.eof) then
	if rs("BigClassName")<>bigclass then
  		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�</font></p>"
		response.end
	end if
do while not rs.eof
SmallClassName=rs("SmallClassName")

Response.Write ("<TABLE cellSpacing=0 cellPadding=0 width=100% bgColor=#f8f6e9 border=0>")
  Response.Write ("<TBODY>")
    Response.Write ("<TR>") 
      Response.Write ("<TD width=3% height=30 background="&bigimg1&" align=center><IMG height=9 src="&imgname1&"></TD>")
      Response.Write ("<TD width=39% background="&bigimg1&"><B><a href=list.asp?type="&bigclass&"&type2="&rs("SmallClassName")&"><SPAN style=FONT-WEIGHT: bold; FONT-SIZE: 14px><font color=#ffffff>"&rs("SmallClassName")&"</font></SPAN></a></B></TD>")
      Response.Write ("<TD width=51% align=middle background="&bigimg1&">&nbsp;</TD>")
      Response.Write ("<TD width=7% background="&bigimg1&"><a href=list.asp?type="&bigclass&"&type2="&rs("SmallClassName")&"><IMG src=images/cw_more.gif border=0></a></TD>")
    Response.Write ("</TR>")
  Response.Write ("</TBODY>")
Response.Write ("</TABLE>")
Response.Write ("<TABLE cellSpacing=0 cellPadding=0 width=100% border=0>")
  Response.Write ("<TBODY>")
    Response.Write ("<TR>")
      Response.Write ("<TD height=61 rowSpan=2 align=middle vAlign=top>")

sql2="select *  from content where BigClassName='"&bigclass&"' and SmallClassName='"&SmallClassName&"' order by id desc "
Set oRS= Server.CreateObject("ADODB.recordset")
oRS.Open sql2,conn,1,3
	if NOT oRS.EOF then 
		for i=1 to 5
  		if NOT oRS.EOF then
		Response.Write ("<TABLE width=98% border=0 align=center cellPadding=0 cellSpacing=0>")
          Response.Write ("<TBODY>")
            Response.Write ("<TR>")
              Response.Write ("<TD height=25 background="&bigimg2&">&nbsp;<font color=#85eb0><strong>��</strong></font>&nbsp;<a href=view.asp?id="&oRS("id")&" target=_blank>")
			  Response.Write ("<font color=#85eb0d>"&oRS("title")&"</font>")
                Response.Write ("</a> <font color=#CCCCCC> &nbsp; ")
                Response.Write (""&oRS("infotime")&"")
                Response.Write ("</font></TD>")
            Response.Write ("</TR>")
          Response.Write ("</TBODY>")
       Response.Write (" </TABLE>")
	   
oRS.MoveNEXT
end if
next
else Response.Write ("��������") end if
oRS.close 
Set oRS = Nothing

        Response.Write ("</TD>")
    Response.Write ("</TR>")
  Response.Write ("</TBODY>")
Response.Write ("</TABLE>")

			rs.movenext
		loop
else 
Response.Write ("<DIV align=center>")
page=clng(request("page"))		 
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from content where BigClassName='"&bigclass&"' order by id desc"
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
response.Write("��������")
else 

Response.Write ("<TABLE cellSpacing=0 cellPadding=0 width=100% border=0>")

      Response.Write ("<TR>")
        Response.Write ("<TD background="&bigimg3&" width=74% height=30>&nbsp;<font color=#ffffff>����</font></TD>")
        Response.Write ("<TD align=middle width=6% background="&bigimg3&"><font color=#ffffff>�Ķ�</font></TD>")
        Response.Write ("<TD align=middle width=20% background="&bigimg3&"><font color=#ffffff>ʱ��</font></TD>")
      Response.Write ("</TR>")

rs.PageSize=linenum
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.AbsolutePage=page  
for j=1 to rs.PageSize 

      Response.Write ("<TR height=24>")
          if rs("firstImageName")<>"" then 
		  Response.Write ("<TD height=25 background="&bigimg2&">&nbsp;<font color=#85eb0><strong>��</strong></font><img src='images/news.gif' border=0 alt='ͼƬ����'>&nbsp;<a href=view.asp?id="& RS("id")&" target=_blank><font color=#85eb0d>"& RS("TITLE")&"</font></a></TD>") 
		  else
		  Response.Write ("<TD height=25 background="&bigimg2&">&nbsp;<font color=#85eb0><strong>��</strong></font>&nbsp;<a href=view.asp?id="& RS("id")&" target=_blank><font color=#85eb0d>"& RS("TITLE")&"</font></a></TD>") 
		  end if
        Response.Write ("<TD align=middle background="&bigimg2&">"& RS("hits")&"</TD>")
        Response.Write ("<TD align=middle background="&bigimg2&">"& RS("infotime")&"</TD>")
      Response.Write ("</TR>")

rs.movenext
if rs.eof then exit for
next

    Response.Write ("</TBODY>")
  Response.Write ("</TABLE>")
  Response.Write ("<CENTER>")
    Response.Write ("<table width=100% border=0>")
     Response.Write ("<tr>") 
        Response.Write ("<td height=40><div align=center><FONT color=#60a060>") 
           
		   
if Page<2 then      
    response.write "<FONT color=#60a060>��ҳ ��һҳ</FONT>&nbsp;"
  else
    response.write "<a href=class.asp?type="&bigclass&"&type2="&smallclass&"&page=1><FONT color=#ff0000>��ҳ</FONT></a>&nbsp;"
    response.write "<a href=class.asp?type="&bigclass&"&type2="&smallclass&"&page=" & Page-1 & "><FONT color=#ff0000>��һҳ</FONT></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "<FONT color=#60a060>��һҳ βҳ</FONT>"
  else
    response.write "<a href=class.asp?type="&bigclass&"&type2="&smallclass&"&page=" & (page+1) & ">"
    response.write "<FONT color=#ff0000>��һҳ</FONT></a> <a href=class.asp?type="&bigclass&"&type2="&smallclass&"&page="&rs.pagecount&"><FONT color=#ff0000>βҳ</FONT></a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#FF0000'>"&rs.recordcount&"</font></b>����Ϣ <b>"&rs.pagesize&"</b>����Ϣ/ҳ"

            Response.Write ("</FONT></div></td>")
      Response.Write ("</tr>")
    Response.Write ("</table>")
  Response.Write ("</CENTER>")
Response.Write ("</DIV>")


end if
rs.close
set rs=nothing


	  end if
	  rs.close
	  set rs=nothing	
end function
%>

<%
function list_content(imgname,bigclass,SmallClass,bigimg1,bigimg2,linenum)
page=clng(request("page"))		 
Set rs=Server.CreateObject("ADODB.RecordSet") 
if bigclass<>"" and smallclass <>"" then
sql="select * from content where BigClassName='"&bigclass&"' and SmallClassName='"&smallclass&"' order by id desc"
rs.Open sql,conn,1,1
elseif bigclass<>"" then
sql="select * from content where BigClassName='"&bigclass&"' order by id desc"
rs.Open sql,conn,1,1
end if
if rs.eof and rs.bof then
response.Write("��������")
else 

Response.Write ("<TABLE cellSpacing=0 cellPadding=0 width=100% border=0>")

      Response.Write ("<TR>")
        Response.Write ("<TD background="&bigimg1&" width=74% height=30>&nbsp;<font color=#ffffff>����</font></TD>")
        Response.Write ("<TD align=middle width=6% background="&bigimg1&"><font color=#ffffff>�Ķ�</font></TD>")
        Response.Write ("<TD align=middle width=20% background="&bigimg1&"><font color=#ffffff>ʱ��</font></TD>")
      Response.Write ("</TR>")

rs.PageSize=linenum
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.AbsolutePage=page  
for j=1 to rs.PageSize 

      Response.Write ("<TR height=24>")
          if rs("firstImageName")<>"" then 
		  Response.Write ("<TD height=25 background="&bigimg2&">&nbsp;<font color=#85eb0><strong>��</strong></font>&nbsp;<img src='images/news.gif' border=0 alt='ͼƬ����'>&nbsp;<a href=view.asp?id="& RS("id")&" target=_blank><font color=#8DDB5D>"& RS("TITLE")&"</font></a></TD>") 
		  else
		  Response.Write ("<TD height=25 background="&bigimg2&">&nbsp;<font color=#85eb0><strong>��</strong></font>&nbsp;<a href=view.asp?id="& RS("id")&" target=_blank><font color=#8DDB5D>"& RS("TITLE")&"</font></a></TD>") 
		  end if
        Response.Write ("<TD align=middle background="&bigimg2&">"& RS("hits")&"</TD>")
        Response.Write ("<TD align=middle background="&bigimg2&">"& RS("infotime")&"</TD>")
      Response.Write ("</TR>")

rs.movenext
if rs.eof then exit for
next

    Response.Write ("</TBODY>")
  Response.Write ("</TABLE>")
  Response.Write ("<CENTER>")
    Response.Write ("<table width=100% border=0>")
     Response.Write ("<tr>") 
        Response.Write ("<td height=40><div align=center><FONT color=#60a060>") 
           
		   
if Page<2 then      
    response.write "<FONT color=#60a060>��ҳ ��һҳ</FONT>&nbsp;"
  else
    response.write "<a href=list.asp?type="&bigclass&"&type2="&smallclass&"&page=1><FONT color=#ff0000>��ҳ</FONT></a>&nbsp;"
    response.write "<a href=list.asp?type="&bigclass&"&type2="&smallclass&"&page=" & Page-1 & "><FONT color=#ff0000>��һҳ</FONT></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "<FONT color=#60a060>��һҳ βҳ</FONT>"
  else
    response.write "<a href=list.asp?type="&bigclass&"&type2="&smallclass&"&page=" & (page+1) & ">"
    response.write "<FONT color=#ff0000>��һҳ</FONT></a> <a href=list.asp?type="&bigclass&"&type2="&smallclass&"&page="&rs.pagecount&"><FONT color=#ff0000>βҳ</FONT></a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#FF0000'>"&rs.recordcount&"</font></b>����Ϣ <b>"&rs.pagesize&"</b>����Ϣ/ҳ"

            Response.Write ("</FONT></div></td>")
      Response.Write ("</tr>")
    Response.Write ("</table>")
  Response.Write ("</CENTER>")
Response.Write ("</DIV>")


end if
rs.close
set rs=nothing


	  
	  rs.close
	  set rs=nothing	
end function
%>


<%
'=============================================================================================================
'view.asp ��id�����Ĺ���
function view_content(title,infotime,hits,content,user)
id=trim(request("id"))
id=replace(id,"'","")
id=left(id,8)
	if id="" then '
  		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�</font></p>"
		response.end
	end if	
	if  not isNumeric(request("id")) then
  		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�</font></p>"
		response.end
	end if
    if len(request("id"))>8 then 
		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�</font></p>"
		response.end
	end if
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="update content set hits=hits+1 where id="&id
conn.execute sql
sql="select * from content where id="&id
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�</font></p>"
		response.end
else
title=rs("title")
infotime= rs("infotime")
hits= rs("hits")
content=rs("content")
end if
rs.close
set rs=nothing
end function
%>

<%
'=============================================================================================================
'view.asp ��id�����Ĺ���
function about_content(title,content,bigclass)
bigclass=request("type")
titlename=request("type2")
	if len(request("type"))>18 then 
		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�1</font></p>"
		response.end
	end if
Set rs=Server.CreateObject("ADODB.RecordSet") 
if titlename<>"" then
sql="select title,content,bigclassname from about where title='"&titlename&"'and bigclassname='"&bigclass&"'"
else
sql="select title,content,bigclassname from about where bigclassname='"&bigclass&"'"
end if
rs.Open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�2</font></p>"
		response.end
	else
		title=rs("title")
		content=rs("content")
		bigclass=rs("bigclassname")
	end if
rs.close
set rs=nothing
end function
%>