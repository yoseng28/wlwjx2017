<!--#include file="adminconn.asp"-->
<%
  Dim edit,types,title,id,ok
%>
<%
if edit=ok then
types=request.form("types")
title=server.htmlencode(request.form("title"))
id=request.form("id")
if title<>"" then
conn.execute("update V_title set types="&types&",title='"&title&"' where id="&id&"")
response.write "<script>alert('�޸�ͶƱ�ɹ���');location.href='vote_admin.asp';</script>"
response.end
end if
end if
%>
<html>
<head>
<title>�޸�ͶƱ</title>
<link href="images/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<br>
<%
id=request.querystring("id")
if cint(id)<>"" then
set rs=server.createobject("adodb.recordset")
rs.open "select id,title,types from V_title where id="&id&"",conn,1,1
if not rs.eof then
%>
<form method="post" action="vote_edit.asp?edit=ok">
  <table width="60%" border="1" align="center" cellpadding="4" cellspacing="1" bordercolorlight="#376ed0" bordercolordark="#376ed0" class="inputt">
    <tr>
      <td colspan="2" bgcolor="#f7f7f7" height="25" background="images/b1.gif"><div align="center"><strong><font color="#FFFFFF">�޸�ͶƱ</font></strong></div></td>
</tr>
<tr>
<td>ͶƱ����</td>
<td>
<select name="types" class="inputt">
<%
select case rs("types")
case "1" '��ѡ
response.write "<option value='1' selected>��ѡ</option><option value='2'>��ѡ</option>"
case "2" '��ѡ
response.write "<option value='1'>��ѡ</option><option value='2' selected>��ѡ</option>"
end select
%>
</td>
</tr>
<tr>
<td>ͶƱ����</td>
<td><input type="text" size="40" name="title" value="<%=rs("title")%>" class="inputt">
<input type="hidden" name="id" value="<%=id%>">
</td>
</tr>
<tr>
<td bgcolor="#f7f7f7" colspan="2" height="30" align="center"><input type="submit" name="submit" value="�޸�" class="inputt">&nbsp;&nbsp;<input type="reset" name="reset" value="�ָ�ԭֵ" class="inputt"></td>
</tr>
</table>
</form>
<%
end if
rs.close
end if
%>

</body>
</html>