<!--#include file="adminconn.asp"-->
<%
Dim DownName,DownBigClass,DownSmallClass,Downlink,Downdaxiao,DownUser,ok,sContent,i
DownName=request("DownName")
DownBigClass=request("BigClassName")
DownSmallClass=request("SmallClassName")
Downdaxiao=request("Downdaxiao")
Downlink=request("firstImageName")
DownUser=request("user")
ok=request("ok")
	sContent = ""
	For i = 1 To Request.Form("d_content").Count
		sContent = sContent & Request.Form("d_content")(i)
	Next

'�ж����������Ƿ�����
if DownName="" or DownBigClass="" or sContent="" then  
response.write"<div align=center><font color=red size=+3><br><br><strong>�������ˡ�</strong></font>"
response.write"<br><br><font color=green>�����������,û����ӳɹ���</font>"
response.write"<br><br>����<font color=red>*</font>�ĵط��Ƿ�ȫ����д��"
response.write"<br><br><a href=javascript:history.back()>���ؼ�����</a></div>"  
response.end 
end if

set rs=server.createobject("adodb.recordset")
sql="select * from download where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("DownName")=DownName
rs("DownBigClass")=DownBigClass
rs("DownSmallClass")=DownSmallClass
rs("Downlink")=Downlink
rs("Downdaxiao")=Downdaxiao
rs("DownUser")=DownUser
if ok<>"" then rs("ok") = ok
rs("DownContent")=sContent
rs.update
rs.close
set rs=nothing
conn.close  
set conn=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('�ύ�ɹ���');" & Chr(13)
		response.write "window.document.location.href='download_info.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
%>