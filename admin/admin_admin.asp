<!--#include file="adminconn.asp" -->
<%
Dim aleave
%>
<html>
<head>
<title>管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="images/style.css" type="text/css">
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="50" valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="8">&nbsp;</td>
        </tr>
      </table>
      <table width="90%" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
        <tr align="center"> 
          <td width="15%" height="25" background="images/b1.gif"> ID</td>
          <td width="15%" background="images/b1.gif">用户</td>
          <td width="20%" background="images/b1.gif">密码(<font color="#666666">加密</font>)</td>
          <td width="20%" background="images/b1.gif">权限</td>
          <td width="15%" background="images/b1.gif">修改</td>
          <td width="15%" background="images/b1.gif">删除</td>
        </tr>
        <%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from admin order by id" 
rs.Open sql,conn,1,1 
while not rs.eof
if rs("aleave")="super" then aleave="超级管理员" end if
if rs("aleave")="check" then aleave="普通管理员" end if
%>
        <tr align="center" bgcolor="#f9f9f9"> 
          <td height="22"><%=rs("id")%></td>
          <td><%=rs("admin")%></td>
          <td><%=rs("password")%></td>
          <td><%=aleave%></td>
          <td><a href="admin_AdminModify.asp?id=<%=rs("id")%>">修改</a></td>
          <td><a href="admin_AdminDel.asp?id=<%=rs("id")%>">删除</a></td>
        </tr>
        <%
rs.movenext
wend
rs.close
set rs=nothing
%>
      </table> 
      <br>
<table width="300" border="1" align="center" cellpadding="4" cellspacing="1" bordercolor="#376ed0">
        <tr> 
          <td height="25" colspan="2" background="images/b1.gif" bgcolor="#CCCCCC"><strong>添加管理员：</strong></td> 
        </tr>
        <form name="add" method="post" action="admin_adminsave.asp">
          <tr bgcolor="#f9f9f9"> 
            <td height="22" align="right">管理帐号：</td>
            <td> 
              <input type="text" name="admin" class="form"> </td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td align="right">管理权限：</td> 
            <td> 
              <select name="aleave" class="input">
<option value=super selected>超级管理员</option> 
<option value=check>普通管理员</option>  
</select>
</td>
</tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="22" align="right">管理密码：</td>
            <td> 
              <input type="password" name="password" class="form"> </td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="确 定"> 
              <input type="hidden" name="act" value="add">
            </td>
          </tr>
        </form>
      </table>
      <br>
    </td>
  </tr>
</table>
</body>
</html>
