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
	ip=request.servervariables("REMOTE_HOST")
	content = trim(Request("TxbContent"))
	content = Server.HTMLEncode(Request("TxbContent"))
	content = Replace( content, vbCrLf, "<br>")
	content = Replace( content, "'", "''")
	
	webname = Server.HTMLEncode(Request("name"))
	weburl = Server.HTMLEncode(Request("homepage"))
	email = Server.HTMLEncode(Request("email"))
	oicq = Server.HTMLEncode(Request("oicq"))
	
	
%>
<%
	SQLcmd = "Insert Into frendlink(webname,content,email,weburl,oicq,time01,ip) Values"
	SQLcmd = SQLcmd & "('" & webname & "','" & content & "','" & email & "','" & weburl & "','" & oicq & "',now,'" & ip & "')"
	conn.Execute SQLcmd
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('����ɹ����ȴ�����Ա��ˣ�');" & Chr(13)
		response.write "window.document.location.href='index.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
%>
