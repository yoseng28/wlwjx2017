<!--#include file="adminconn.asp" -->
<%
Dim id,password,aleave
if request("xg")="yes" then
password=request.form("password")
if password="" then
response.write"<SCRIPT language=JavaScript>alert('���벻��Ϊ�գ�');"
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
Response.Write"alert(""��ϲ.�����޸ĳɹ�������������ǡ�"&password&"��,�����µ�½"");"
Response.Write"window.location='admin_adminmodify_check.asp'"
Response.Write"</script>"
end if
%>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from admin where admin='"&session("admin")&"'"
rs.open sql,conn,1,1
if rs.eof then
response.write"<SCRIPT language=JavaScript>alert('��������������ϵ����Ա��');"
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
      <td height="25" colspan=2 align=center background="images/b1.gif"><b>�� �� 
        �� �� Ա �� ��</b></td>
    </tr>
    <tr> 
      <td width="21%" height="22" align="right" bgcolor="#f9f9f9">�û�����</td>
      <td width="79%" bgcolor="#f9f9f9"> <input type="text" name="admin" value="<%=admin%>" size="20" class="input">
        ���ѣ��û��������޸ġ�<font color="#FF0000"><strong>��ʹ����Ҳ��Ч</strong></font>��</td>
    </tr>
    <tr> 
      <td width="21%" height="22" align="right" bgcolor="#f9f9f9">���룺</td>
      <td width="79%" bgcolor="#f9f9f9"> <p> 
          <input type="text" name="password" size="20" class="input">
          ���ѣ���������Ҫ�޸ĵ����롣</p></td>
    </tr>
    <tr align="center" height="24"> 
      <td height="30" colspan=2 bgcolor="#f9f9f9"> 
        <input name="cmdok" type="submit" id="cmdok" value=" �� �� " class="input"> 
        &nbsp; <input name="cmdcance" type="reset" id="cmdcance" value=" �� �� " class="input"> 
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
