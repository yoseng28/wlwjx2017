<!--#include file="adminconn.asp" -->
<%
Dim id,password,aleave
if session("aleave")="check" then
response.write"<SCRIPT language=JavaScript>alert('对不起，你没有这个权限！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
%>
<%
id=request.QueryString("id")
set rs=server.createobject("adodb.recordset")
sql="select * from admin where id="&id
rs.open sql,conn,1,1
if rs.eof then
response.write"<SCRIPT language=JavaScript>alert('服务器出错，请联系管理员！');"
response.write"javascript:history.go(-1)</SCRIPT>"
else
admin=rs("admin")
password=rs("password")
aleave=rs("aleave")
%>
<LINK href="images/style.css" type=text/css rel=stylesheet>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<table width="90%" border="1" align="center" cellpadding="4" cellspacing="1" bordercolor="#376ed0">
  <form method="POST" action="admin_adminSave.asp?id=<%=id%>">
          
    <tr> 
      <td width="100%" height="25" colspan=2 align=center background="images/b1.gif"><b>修 
        改 管 理 员 资 料</b></td>
</tr>
    <tr> 
      <td width="30%" height="22" align="right" bgcolor="#f9f9f9">用户名：</td>
      <td width="70%" bgcolor="#f9f9f9"> 
        <input type="text" name="admin" value="<%=admin%>" size="20" class="input"></td>
</tr>
    <tr> 
      <td width="30%" height="22" align="right" bgcolor="#f9f9f9">密码：</td>
      <td width="70%" bgcolor="#f9f9f9"> 
        <input type="text" name="password" value="<%=decrypt(rs("password"))%>" size="20" class="input"></td>
</tr>
    <tr> 
      <td width="30%" height="22" align="right" bgcolor="#f9f9f9">权限：</td>
      <td width="70%" height="22" bgcolor="#f9f9f9"> 
        <select name="aleave" style="font-size:9pt" class="input">
<option value=super<%if aleave="super" then%> selected<%end if%>>超级管理员</option>
<option value=check<%if aleave="check" then%> selected<%end if%>>普通管理员</option>
</select>
</td>
</tr>
    <tr align="center" height="24"> 
      <td height="30" colspan=2 bgcolor="#f9f9f9"> 
        <input type="hidden" value="edit" name="act">
<input name="cmdok" type="submit" id="cmdok" value=" 修 改 " class="input">
        &nbsp;
<input name="cmdcance" type="reset" id="cmdcance" value=" 清 除 " class="input">
</td>
</tr>
</form>
</table>
</body>
</html>
<%
end if
rs.close
set rs=nothing
%>
