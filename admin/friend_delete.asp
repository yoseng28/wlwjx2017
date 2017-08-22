<!--#include file="adminconn.asp"-->
<%
Dim id
%>
<html>
<head>
<title>友情链接-删除</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="5">
<%
	id = trim(Request("id"))
	Set rs= Server.CreateObject("ADODB.recordset")
	sql = "delete from frendlink where id="&id
	response.write "<table width='100%' border=0 cellspacing=0 cellpadding=0 height='100%'><tr><td>"
	response.write "<table width=250 border=1 cellspacing=2 cellpadding=2 align=center bordercolorlight=7285CF bordercolordark=7285CF height=90 bgcolor=F2F2F2>"
	response.write "  <tr bordercolor=F2F2F2 bgcolor=F2F2F2 align=center> "
	response.write "<td>删除成功！</td>"
	response.write "</tr>"
	response.write "<tr bordercolor=F2F2F2 bgcolor=F2F2F2 align=center> "
	response.write "<td>您已经成功的将此条数据删除！<br><br><a href=friend_admin.asp>点击返回</a></td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "</td></tr></table>"
	conn.Execute sql
%>
 
</body>
</html>