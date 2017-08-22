<!--#include file="adminconn.asp"-->
<%
  Dim id
%>
<html>
<head>
<title>留言管理-删除</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="5">
<%
	id = trim(Request("id"))
	Set rs= Server.CreateObject("ADODB.recordset")
	
	sql = "delete from guestbook where id="&id
	Response.Write("<SCRIPT language=JavaScript>alert('提示：你已经成功删除留言操作成功！');this.location.href='"&request.ServerVariables("HTTP_REFERER")&"';</SCRIPT>")
	conn.Execute sql
%>
 
</body>
</html>