<!--#include file="adminconn.asp" -->
<%
Dim aleave
%>
<html>
<head>
<title>����</title>
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
          <td width="15%" background="images/b1.gif">�û�</td>
          <td width="20%" background="images/b1.gif">����(<font color="#666666">����</font>)</td>
          <td width="20%" background="images/b1.gif">Ȩ��</td>
          <td width="15%" background="images/b1.gif">�޸�</td>
          <td width="15%" background="images/b1.gif">ɾ��</td>
        </tr>
        <%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from admin order by id" 
rs.Open sql,conn,1,1 
while not rs.eof
if rs("aleave")="super" then aleave="��������Ա" end if
if rs("aleave")="check" then aleave="��ͨ����Ա" end if
%>
        <tr align="center" bgcolor="#f9f9f9"> 
          <td height="22"><%=rs("id")%></td>
          <td><%=rs("admin")%></td>
          <td><%=rs("password")%></td>
          <td><%=aleave%></td>
          <td><a href="admin_AdminModify.asp?id=<%=rs("id")%>">�޸�</a></td>
          <td><a href="admin_AdminDel.asp?id=<%=rs("id")%>">ɾ��</a></td>
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
          <td height="25" colspan="2" background="images/b1.gif" bgcolor="#CCCCCC"><strong>��ӹ���Ա��</strong></td> 
        </tr>
        <form name="add" method="post" action="admin_adminsave.asp">
          <tr bgcolor="#f9f9f9"> 
            <td height="22" align="right">�����ʺţ�</td>
            <td> 
              <input type="text" name="admin" class="form"> </td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td align="right">����Ȩ�ޣ�</td> 
            <td> 
              <select name="aleave" class="input">
<option value=super selected>��������Ա</option> 
<option value=check>��ͨ����Ա</option>  
</select>
</td>
</tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="22" align="right">�������룺</td>
            <td> 
              <input type="password" name="password" class="form"> </td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="ȷ ��"> 
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
