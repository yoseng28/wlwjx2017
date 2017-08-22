<!--#include file="adminconn.asp"-->
<%
dim iid,cont,id,titles,modi,ok,vcount,i
if modi=ok then
iid=request.form("iid")
id=request.form("id")
cont=server.htmlencode(request.form("cont"))
vcount=request.form("vcount")
if iid<>"" and cont<>"" then
conn.execute("update V_vote set cont='"&cont&"',vcount="&vcount&" where id=cint('"&iid&"')")
response.write"<SCRIPT language=JavaScript>alert('修改成功！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
end if
%>



<html>
<head>
<title>投票内容</title>
<link href="images/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<br>

<table width="70%" border="1" align="center" cellpadding="2" cellspacing="1" bordercolorlight="#376ed0" bordercolordark="#376ed0" class="css">
  <tr> 
    <td height="25" colspan="4" background="images/b1.gif" bgcolor="#efefef">
<div align="center"><b><font color="#FFFFFF">修改、添加投票选项</font></b></div></td>
</tr>
<tr>
    <td height="30">编号</td>
    <td height="30">投票选项</td>
    <td height="30">投票数</td>
    <td height="30">操作</td>
</tr>

<%
id=request.querystring("id")
if cint(id)<>"" then
set rs=server.createobject("adodb.recordset")
rs.open "select * from V_vote where lid=cint('"&id&"')",conn,1,1
if not rs.eof then
do while not rs.eof
i=i+1
%>
<tr>
    <td height="30" align="center"><%=i%></td>
<form method="post" action="vote_item.asp?modi=ok">
      <td height="30">
<input type="text" name="cont" class="inputt" size="35" value="<%=rs("cont")%>"><input type="hidden" name="iid" value="<%=rs("id")%>"><input type="hidden" name="id" value="<%=id%>"><input type="hidden" name="titles" value="<%=titles%>"></td>
      <td height="30">
<input type="text" name="vcount" class="inputt" size="3" value="<%=rs("vcount")%>">
      <td height="30">
<input type="submit" name="submit" value="修改" class="inputt">&nbsp;&nbsp;<input type="button" name="button" value="删除" class="inputt" onclick="if (confirm('您真的要删除该选项吗？\n\n注意：该操作不可恢复'))location.href='vote_delitem.asp?iid=<%=rs("id")%>&id=<%=id%>'"></td>
</form>
</tr>
<%
rs.movenext
loop
end if
rs.close
end if
%>
<form method="post" action="vote_additem.asp">
<tr>
      <td height="30" colspan="4">添加选项&nbsp; 
        <input type="text" name="cont" class="inputt"><input type="hidden" name="id" value="<%=id%>">
<input type="submit" name="submit" value="添加" class="inputt">
</td>
</tr>
</form>
</table>
</body>
</html>