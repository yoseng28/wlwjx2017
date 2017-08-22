<!--#include file="adminconn.asp" -->
<%
Dim id,password,aleave
if request("xg")="yes" then
password=request.form("password")
if password="" then
response.write"<SCRIPT language=JavaScript>alert('密码不能为空！');"
response.write"javascript:history.go(-1)</SCRIPT>"
Response.End
end if
sql="select * from admin where admin='"&session("admin")&"'"
set rs=Server.CreateObject("ADODB.recordset") 
rs.open sql,conn,1,3
rs("password")=md5(password,32)
rs.update
rs.close

Response.Write"<script language=JavaScript>"
Response.Write"alert(""恭喜.密码修改成功，你的新密码是“"&password&"”,请重新登陆"");"
Response.Write"window.location='admin_adminmodify_check.asp'"
Response.Write"</script>"
end if
%>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from admin where admin='"&session("admin")&"'"
rs.open sql,conn,1,1
if rs.eof then
response.write"<SCRIPT language=JavaScript>alert('服务器出错，请联系管理员！');"
response.write"javascript:history.go(-1)</SCRIPT>"
else
admin=rs("admin")
id=rs("id")
%>
<LINK href="images/style.css" type=text/css rel=stylesheet>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<table width="90%" border="1" align="center" cellpadding="4" cellspacing="1" bordercolor="#376ed0">
  <form method="POST" action="admin_adminmodify_check.asp?xg=yes">
    <tr> 
      <td height="25" colspan=2 align=center background="images/b1.gif"><b>修 改 
        管 理 员 密 码</b></td>
    </tr>
    <tr> 
      <td width="21%" height="22" align="right" bgcolor="#f9f9f9">用户名：</td>
      <td width="79%" bgcolor="#f9f9f9"> <input type="text" name="admin" value="<%=admin%>" size="20" class="input">
        提醒：用户名不能修改。<font color="#FF0000"><strong>即使输入也无效</strong></font>。</td>
    </tr>
    <tr> 
      <td width="21%" height="22" align="right" bgcolor="#f9f9f9">密码：</td>
      <td width="79%" bgcolor="#f9f9f9"> <p> 
          <input type="text" name="password" size="20" class="input">
          提醒：请输入你要修改的密码。</p></td>
    </tr>
    <tr align="center" height="24"> 
      <td height="30" colspan=2 bgcolor="#f9f9f9"> 
        <input name="cmdok" type="submit" id="cmdok" value=" 修 改 " class="input"> 
        &nbsp; <input name="cmdcance" type="reset" id="cmdcance" value=" 清 除 " class="input"> 
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
