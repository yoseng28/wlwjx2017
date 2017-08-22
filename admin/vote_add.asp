<!--#include file="adminconn.asp"-->
<%
  Dim types,title
%>
<html>
<head>
<title>添加投票</title>
<link href="images/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<br>
<%
types=request.form("types")
title=server.htmlencode(request.form("title"))
if types<>"" then
conn.execute("insert into V_title (types,title) values ("&types&",'"&title&"')")
response.write "<script>alert('投票添加成功，请给新的投票添加项目');location.href='vote_admin.asp';</script>"
response.end
end if
%>


<form method="post" action="<%=request.servervariables("script_name")%>">
  <table width="70%" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0" bordercolorlight="#376ed0" bordercolordark="#376ed0" class="inputt">
    <tr> 
      <td height="25" colspan="2" background="images/b1.gif" bgcolor="#f7f7f7">
<div align="center"><strong><font color="#FFFFFF">添加新的投票</font></strong></div></td>
</tr>
<tr>
      <td height="30">投票类型</td>
      <td height="30">
<select name="types" class="inputt">
	<option value="1" selected>单选</option>
	<option value="2">多选</option>
	</select>
</td>
</tr>
<tr>
      <td height="30">投票标题</td>
      <td height="30">
<input type="text" name="title" size="30" maxlength="150" class="inputt"></td>
</tr>
<tr>
      <td colspan="2" align="center" height="30">
<input type="submit" name="submit" value="添加新的投票" class="inputt">
</td>
</tr>
</table>
</form>
</body>
</html>