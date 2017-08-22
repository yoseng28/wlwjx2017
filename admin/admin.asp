<!--#include file="adminconn.asp" -->
<%
  if session("aleave")="" then
      response.redirect "adminlogin.asp"
	  response.end
  end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
  <title>二货帮网站后台 </title>
  <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="renderer" content="ie-comp"> 
<meta name="renderer" content="ie-stand">

</head>
	<frameset cols="185,8,*" framespacing="0" frameborder="0" border="false" id="frame" scrolling="yes">
	<frame name="left" scrolling="yes" marginwidth="0" marginheight="0" src="admin_left.asp">
	<frame src="center.htm" scrolling="no">
	<frame name="main" scrolling="auto" marginwidth="0" marginheight="0" src="admin_center.asp">
	</frameset>
<noframes><body>
<font color=red style=font-size:14px><b>你的浏览器不支持帧页面，请升级！</b></font> 
</body></noframes>
</html>