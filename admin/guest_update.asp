<!--#include file="adminconn.asp"-->
<%
  Dim replycontent,ors
%>
<%
if request("modi")="yes" then
replycontent=request.form("replycontent")
sql="select * from guestbook where id="& request.QueryString("id") &" "
set ors=Server.CreateObject("ADODB.recordset") 
ors.open sql,conn,1,3
ors("reply")="-1"
ors("replycontent")=replycontent
ors.update
ors.close


Response.Write"<script language=JavaScript>"
Response.Write"alert(""恭喜.留言回复成功，现在返回留言管理页面成功"");"
Response.Write"window.location='guest_admin.asp'"
Response.Write"</script>"
end if
%>
<html>
<head>
<title>留言回复</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>BODY {
	COLOR: #666666; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; FONT-SIZE: 11px
}
TD {
	COLOR: #666666; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; FONT-SIZE: 11px
}
TH {
	COLOR: #666666; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; FONT-SIZE: 11px
}
BODY 
.bg {
	BACKGROUND-COLOR: #ffffff; BACKGROUND-REPEAT: no-repeat; BORDER-BOTTOM: #cfedf7 1px solid; BORDER-LEFT: #cfedf7 1px solid; BORDER-RIGHT: #cfedf7 1px solid; BORDER-TOP: #cfedf7 1px solid; COLOR: #4b4b4b; FONT-SIZE: 12px
}
.style13 {
	COLOR: #ff0000
}
.distance {
	LINE-HEIGHT: 20px
}
A:link {
	COLOR: #333333; TEXT-DECORATION: none
}
A:visited {
	COLOR: #333333; TEXT-DECORATION: none
}
A:hover {
	COLOR: #ff6600; TEXT-DECORATION: none
}
A:active {
	COLOR: #ff6600; TEXT-DECORATION: none
}
.style14 {
	COLOR: #378cb5
}
.itable {
	background-position: top;
	table-layout:fixed;
	word-wrap:break-word;
	word-break:break-all;
}
</STYLE>

 <%
	sql ="select  * from guestbook where id="& request.QueryString("id") &" "
	set rs=Server.CreateObject("ADODB.recordset") 
	rs.open sql,conn,1,3
%> 
</head>
<body bgcolor="#FFFFFF" onload="window.focus();">
<center>
  <form method="POST" action="guest_update.asp?id=<%=rs("id")%>&modi=yes">
    <table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr> 
        <td height="1" bgcolor="7285CF"></td>
      </tr>
      <tr align="center" bgcolor="F2F2F2"> 
        <td height="30"><font color="#CCCCCC">■</font> <font color="#999999">■</font> 
          <font color="#666666">■</font>　留言回复　<font color="#666666">■</font> <font color="#999999">■ 
          </font><font color="#CCCCCC">■</font></td>
      </tr>
      <tr bgcolor="7285CF"> 
        <td height="1"></td>
      </tr>
    </table>
    <br>
    <table width="500" border="1" cellspacing="2" cellpadding="0" align="center" bordercolor="#FFFFFF">
      <tr> 
        <td bordercolor="7285CF" width="84" height="22" align="center" bgcolor="F2F2F2">留言内容</td>
        <td bordercolor="7285CF" width="12" bgcolor="F2F2F2">&nbsp;</td>
        <td bordercolor="7285CF" width="288" bgcolor="F2F2F2">&nbsp;<%=rs("guestcontent")%></td>
      </tr>
      <tr> 
        <td bordercolor="7285CF" width="84" height="22" align="center" bgcolor="F2F2F2">&nbsp;</td>
        <td bordercolor="7285CF" width="12" bgcolor="F2F2F2">&nbsp;</td>
        <td bordercolor="7285CF"width="288" bgcolor="F2F2F2">&nbsp; </td>
      </tr>
      <tr> 
        <td bordercolor="7285CF" width="84" height="22" align="center" bgcolor="F2F2F2">回复内容</td>
        <td bordercolor="7285CF" width="12" bgcolor="F2F2F2">&nbsp;</td>
        <td bordercolor="7285CF" width="288" bgcolor="F2F2F2"> 
          <textarea name="replycontent" cols="50" rows="10" class="editbox2" id="replycontent"><%=rs("replycontent")%></textarea> 
        </td>
      </tr>
      <tr> 
        <td height="30" colspan="3" bordercolor="7285CF"><div align="center"><font color="#FF0000"><strong>注意：回复后留言将自动通过审核！</strong></font></div></td>
      </tr>
      <tr>
        <td bordercolor="7285CF" colspan="3"> <div align="center">
            <input class=bottom name=B12 type=submit value="修改记录">
            　
<input class=bottom name=B22 type=reset value="重新添写">
          </div></td>
      </tr>
    </table>
    <br>
    <div align="center"></div>
    </form>   
</center>
</body>
</html>