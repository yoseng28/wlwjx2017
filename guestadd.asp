<!--#include file="conn.asp"-->
<%
ComeUrl=lcase(trim(request.ServerVariables("HTTP_REFERER")))
if ComeUrl="" then
	response.write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ��������ֱ�������ַ���ʱ�ϵͳ�ĺ�̨����ҳ�档</font></p>"
	response.end
else
	cUrl=trim("http://" & Request.ServerVariables("SERVER_NAME"))
	if mid(ComeUrl,len(cUrl)+1,1)=":" then
		cUrl=cUrl & ":" & Request.ServerVariables("SERVER_PORT")
	end if
	cUrl=lcase(cUrl & request.ServerVariables("SCRIPT_NAME"))
	if lcase(left(ComeUrl,instrrev(ComeUrl,"/")))<>lcase(left(cUrl,instrrev(cUrl,"/"))) then
		response.write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ����������ⲿ���ӵ�ַ���ʱ�ϵͳ�ĺ�̨����ҳ�档</font></p>"
		response.end
	end if
end if
	'************add.asp************
	a_ip=request.servervariables("REMOTE_HOST")
	a_guestcontent = trim(Request("TxbContent"))
	a_guestcontent = Server.HTMLEncode(Request("TxbContent"))
	a_guestcontent = Replace( a_guestcontent, vbCrLf, "<br>")
	a_guestcontent = Replace( a_guestcontent, "'", "''")
	
	a_name = Server.HTMLEncode(Request("name"))
	a_homepage = Server.HTMLEncode(Request("homepage"))
	a_email = Server.HTMLEncode(Request("email"))
	a_oicq = Server.HTMLEncode(Request("oicq"))
	a_from01 = Server.HTMLEncode(Request("from01"))
	classname = Server.HTMLEncode(Request("classname"))
	
	
%>
<%
	SQLcmd = "Insert Into guestbook(name,guestcontent,email,homepage,oicq,from01,reply,time01,ip,classname) Values"
	SQLcmd = SQLcmd & "('" & a_name & "','" & a_guestcontent & "','" & a_email & "','" & a_homepage & "','" & a_oicq & "','" & a_from01 & "'," & reply & ",now,'" & a_ip & "','" & classname & "')"
	conn.Execute SQLcmd
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('�����Ϣ�Ѿ��ύ�ɹ���');" & Chr(13)
		response.write "window.document.location.href='guestbook.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
%>
