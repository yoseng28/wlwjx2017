<%
if request("logout")<>"" then 
session("admin")=""
session("password")=""
session("aleave")=""
response.redirect "adminlogin.asp"
end if
%>
<html>
<head>
<LINK href="images/style.css" type=text/css rel=stylesheet>
<title>��̨����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<p align="center"><br>
  <div align="center">
<table width="1003" border="0" cellpadding="0" cellspacing="1" background="log.png" height="438">
  <form method=post action="chklogin.asp" target="_top">
    <tr> 
      <td width="100%" height="274" align="center">
		<p style="margin-top: 0; margin-bottom: 0">��</td>
    </tr>
    <tr> 
      <td width="100%" align="center">
		<p style="margin-top: 0; margin-bottom: 0"><font face="΢���ź�">
		<font color="#FFFFFF"><b>�û�����</b> </font> 
        <input name="adminname" class="input" id="adminname" style="color: #0064A8; background-color: #FFFFFF; border: 1px solid #000000" size="12"></font></td>
    </tr>
    <tr> 
      <td width="100%" height="25" align="center">
		<p style="margin-top: 0; margin-bottom: 0"><font face="΢���ź�">
		<font color="#FFFFFF"><b>�ܡ��룺</b> </font> 
        <input name="adminpassword" type="password" class="input" id="adminpassword" style="color: #0064A8; background-color: #FFFFFF; border: 1px solid #000000" size="12"></font></td>
    </tr>
    <tr> 
      <td width="100%" height="37" align="center"> 
        <p style="margin-top: 0; margin-bottom: 0"> 
        &nbsp;&nbsp; 
        <input class="input" style="color: #FFFFFF; background-color: #0099CC; border: 1px solid #FFFFFF" type="submit" value="ȷ ��" name="Submit"> 
        &nbsp;&nbsp;&nbsp; <input  class="input" style="color: #FFFFFF; background-color: #0099CC; border: 1px solid #FFFFFF" type="reset" value="ȡ ��" name="Submit2"> 
      </td>
    </tr>
    <tr> 
      <td width="100%" height="60" align="center"> 
        <p style="margin-top: 0; margin-bottom: 0">��</td>
    </tr>
  </form>
</table>
</div>
<br>
<br><br>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0" height="38">
  <tr>
    <td align="center">��</td>
  </tr>
</table>
</body>
</html>
