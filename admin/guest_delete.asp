<!--#include file="adminconn.asp"-->
<%
  Dim id
%>
<html>
<head>
<title>���Թ���-ɾ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="5">
<%
	id = trim(Request("id"))
	Set rs= Server.CreateObject("ADODB.recordset")
	
	sql = "delete from guestbook where id="&id
	Response.Write("<SCRIPT language=JavaScript>alert('��ʾ�����Ѿ��ɹ�ɾ�����Բ����ɹ���');this.location.href='"&request.ServerVariables("HTTP_REFERER")&"';</SCRIPT>")
	conn.Execute sql
%>
 
</body>
</html>