<!--#include file="adminconn.asp"-->
<%
Dim currf,backf,backfy,objfso,fy

 if Request.QueryString("action")="back" then
 currf=request.form("currf")
 currf=server.mappath(currf)
  backf=request.form("backf")
backf=server.mappath(backf)
   backfy=request.form("backfy")
   on error resume next
 Set objfso = Server.CreateObject("Scripting.FileSystemObject")
if err then 
	err.clear
response.write "<script>alert(""不能建立fso对象，请确保你的空间支持fso:！"");history.back();</script>"
response.end
	end if
 if objfso.Folderexists(backf) then
 else
  Set fy=objfso.CreateFolder(backf)
  end if
 objfso.copyfile currf,backf& "\"& backfy
response.write "<script>alert(""备份数据库成功"");history.back();</script>"
end if 
%>
<head>
<title>管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/style.css" rel="stylesheet" type="text/css">
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
  <div align="center">
  <center>
    <table width="98%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>&nbsp;</td>
      </tr>
    </table>
    <table border=1 borderColor=#376ed0 cellPadding=4 cellSpacing=1 width=90% style="border-collapse: collapse" align=center>
      <form name="form1" method="POST" action="?action=back">
        <tr align="center" bgcolor="#f9f9f9"> 
          <td height=25 colspan="2"  background="images/b1.gif" class=tdc1> 
            <font color="#FFFFFF"><strong>备份数据库</strong></font></td>
        </tr>
        <tr align="center" bgcolor="#f9f9f9"> 
          <td class=tdc height=22 colspan="2">你的空间只有支持fso才可以进行如下操作，否则你只能手动备份</td>
        </tr>
        <tr bgcolor="#f9f9f9"> 
          <td width="30%" align="right" class=tdc>数据库路径：</td>
          <td class=tdc> <input type="text" name="currf" size="20" value="<%=admindataurl%>">
            请不要修改该项</td>
        </tr>
        <tr bgcolor="#f9f9f9"> 
          <td width="30%" align="right" class=tdc>备份数据目录：</td>
          <td class=tdc> <input type="text" name="backf" size="20" value="databack">
            请不要修改该项 </td>
        </tr>
        <tr bgcolor="#f9f9f9"> 
          <td width="30%" align="right" class=tdc>数据库名称：</td>
          <td class=tdc> <input type="text" name="backfy" size="20" value="<% =date() %>.mdb"> 
          </td>
        </tr>
        <tr bgcolor="#f9f9f9">
          <td width="30%" align="right" class=tdc>&nbsp;</td>
          <td bgcolor="#f9f9f9" class=tdc>注意：为了你的安全，在备份数据库后请到 <a href="http://127.0.0.1/admin/Edit_Admin_UploadFile.asp?sUploadDir=databack" target="main">备份数据库管理</a> 
            下载数据库，并删除备份数据库。</td>
        </tr>
        <tr bgcolor="#f9f9f9"> 
          <td width="30%" align="right" class=tdc>　</td>
          <td class=tdc> <input type="submit" name="Submit" value="提交" class=bdtj> 
            <input type="reset" name="Submit2" value="重置" class=bdtj> </td>
        </tr>
      </form>
    </table>
  </center>
  </div>
  </body>
</html>