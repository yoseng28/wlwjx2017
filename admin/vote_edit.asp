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
response.write "<script>alert('修改投票成功！');location.href='vote_admin.asp';</script>"
response.end
end if
end if
%>
<html>
<head>
<title>修改投票</title>
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
      <td colspan="2" bgcolor="#f7f7f7" height="25" background="images/b1.gif"><div align="center"><strong><font color="#FFFFFF">修改投票</font></strong></div></td>
</tr>
<tr>
<td>投票类型</td>
<td>
<select name="types" class="inputt">
<%
select case rs("types")
case "1" '单选
response.write "<option value='1' selected>单选</option><option value='2'>多选</option>"
case "2" '多选
response.write "<option value='1'>单选</option><option value='2' selected>多选</option>"
end select
%>
</td>
</tr>
<tr>
<td>投票名称</td>
<td><input type="text" size="40" name="title" value="<%=rs("title")%>" class="inputt">
<input type="hidden" name="id" value="<%=id%>">
</td>
</tr>
<tr>
<td bgcolor="#f7f7f7" colspan="2" height="30" align="center"><input type="submit" name="submit" value="修改" class="inputt">&nbsp;&nbsp;<input type="reset" name="reset" value="恢复原值" class="inputt"></td>
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