<!--#include file="adminconn.asp"-->
<html>
<head>
<title>����ͶƱ</title>
<link href="images/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<br>
<table width="90%" border="1" align="center" cellpadding="2" cellspacing="1" bordercolorlight="#376ed0" bordercolordark="#376ed0" class="css">
  <tr> 
    <td height="25" colspan="4" background="images/b1.gif" bgcolor="#efefef">
<div align="center"><strong><font color="#FFFFFF">�鿴���е�ͶƱ</font></strong></div></td>
</tr>
<tr>
    <td width="10%" height="25" bgcolor="#f7f7f7">ͼ��</td>
    <td width="36%" bgcolor="#f7f7f7">ͶƱ���� (<font color="#0000FF">������Ʋ鿴�޸�ͶƱѡ��</font>)</td>
<td width="21%" bgcolor="#f7f7f7">����</td>
<td width="33%" bgcolor="#f7f7f7">��ʾ���</td>
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
<a href="vote_edit.asp?id=<%=rs("id")%>">�޸�</a>
<a href="#" onclick="if (confirm('�����Ҫɾ����ͶƱ��\n\nע�⣺�ò������ɻָ�'))location.href='vote_del.asp?id=<%=rs("id")%>'">ɾ��</a>
</td>
<td><% if rs("xindex")="Yes" then %>
  <span class="style1">����ʾ</span>����<a href="vote_set.asp?id=<% =rs("id") %>&setindex=No"><font color="#FF0000">ȡ����ʾ</font></a>
  <% else %>
  <span class="style2">δ��ʾ</span>����<a href="vote_set.asp?id=<% =rs("id") %>&setindex=Yes"><font color="#0000FF">������ʾ</font></a>
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