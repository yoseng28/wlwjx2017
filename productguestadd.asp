<!--#include file="conn.asp"-->
<%
	'************add.asp************
	a_ip=request.servervariables("REMOTE_HOST")
	a_name = trim(Request("name"))						'取得表单姓名
	a_email= trim(Request("email"))
	a_homepage = trim(Request("homepage"))
	a_oicq = trim(Request("oicq"))
	a_from01 = trim(Request("from01"))
	lifang = trim(Request("lifang"))
    ruzhu = trim(Request("ruzhu"))
    fangshi = trim(Request("fangshi"))
    zhifu = trim(Request("zhifu "))
	a_guestcontent = trim(Request("TxbContent"))					' 取得愿望内容
	a_guestcontent = Server.HTMLEncode(Request("TxbContent"))
	a_guestcontent = Replace( a_guestcontent, vbCrLf, "<br>")
	a_guestcontent = Replace( a_guestcontent, "'", "''")
	
	a_name = Server.HTMLEncode(Request("name"))
	a_homepage = Server.HTMLEncode(Request("homepage"))
	a_email = Server.HTMLEncode(Request("email"))
	a_oicq = Server.HTMLEncode(Request("oicq"))
	lifang = Server.HTMLEncode(Request("lifang"))
	ruzhu = Server.HTMLEncode(Request("ruzhu"))
	fangshi = Server.HTMLEncode(Request("fangshi"))
	zhifu = Server.HTMLEncode(Request("zhifu"))
	a_from01 = Server.HTMLEncode(Request("from01"))
	biaotiid = Server.HTMLEncode(Request("biaotiid"))
	classname = Server.HTMLEncode(Request("classname"))
	
	
%>
<%
	SQLcmd = "Insert Into guestbook(name,guestcontent,email,homepage,lifang,ruzhu,fangshi,jianshu,zhifu,oicq,from01,reply,time01,ip,biaotiid,classname) Values"
	SQLcmd = SQLcmd & "('" & a_name & "','" & a_guestcontent & "','" & ruzhu & "','" & a_email & "','" & lifang & "','" & jianshu & "','" & fangshi & "','" & zhifu & "','" & a_homepage & "','" & a_oicq & "','" & a_from01 & "'," & reply & ",now,'" & a_ip & "','" & biaotiid & "','" & classname & "')"
	conn.Execute SQLcmd
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('你的信息已经发布成功，我们会尽快和你联系！');" & Chr(13)
		response.write "window.document.location.href='productview.asp?id="&biaotiid&"';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
%>
