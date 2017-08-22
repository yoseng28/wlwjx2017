<!--#include file="adminconn.asp"-->
<html>
<head>
<title>管理投票</title>
<link href="images/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<br>
<table width="90%" border="1" align="center" cellpadding="2" cellspacing="1" bordercolorlight="#376ed0" bordercolordark="#376ed0" class="css">
  <tr> 
    <td height="25" colspan="4" background="images/b1.gif" bgcolor="#efefef">
<div align="center"><strong><font color="#FFFFFF">查看已有的投票</font></strong></div></td>
</tr>
<tr>
    <td width="10%" height="25" bgcolor="#f7f7f7">图例</td>
    <td width="36%" bgcolor="#f7f7f7">投票名称 (<font color="#0000FF">点击名称查看修改投票选项</font>)</td>
<td width="21%" bgcolor="#f7f7f7">操作</td>
<td width="33%" bgcolor="#f7f7f7">显示情况</td>
</tr>

<%
set rs=server.createobject("adodb.recordset")
rs.open "select id,title,xindex from V_title",conn,1,1
if not rs.eof then
do while not rs.eof
%>
<tr>
    <td height="30"><img src="images/isvote.gif" width="13" height="15" border="0"></td>
<td><a href="vote_item.asp?id=<%=rs("id")%>"><%=rs("title")%></a></td>
<td>
<a href="vote_edit.asp?id=<%=rs("id")%>">修改</a>
<a href="#" onclick="if (confirm('您真的要删除该投票吗？\n\n注意：该操作不可恢复'))location.href='vote_del.asp?id=<%=rs("id")%>'">删除</a>
</td>
<td><% if rs("xindex")="Yes" then %>
  <span class="style1">已显示</span>　　<a href="vote_set.asp?id=<% =rs("id") %>&setindex=No"><font color="#FF0000">取消显示</font></a>
  <% else %>
  <span class="style2">未显示</span>　　<a href="vote_set.asp?id=<% =rs("id") %>&setindex=Yes"><font color="#0000FF">设置显示</font></a>
  <% end if %></td>
</tr>
<%
rs.movenext
loop
end if
rs.close
%>
</table>

</body>
</html>