<%
function renamefolder(sFolder,dFolder)     
on error resume next     
dim fso     
set fso = server.createobject("scripting.filesystemobject")     
if fso.folderexists(server.mappath(sFolder)) then     
fso.movefolder server.mappath(sFolder),server.mappath(dFolder)     
renamefolder = true     
else     
renamefolder = false     
set fso = nothing     
call alertbox("系统没有找到指定的路径[" & sFolder & "]!",2)     
end if     
set fso = nothing     
end function    
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>
<title>系统初始设置</title>
</head>
<BODY bgcolor="#f0f0f0" topMargin=0 leftmargin="0">
<%
modi=trim(request("modi"))
if modi="yes" then
datadir=trim(request("datadir"))
dataname=trim(request("dataname"))
admindir=trim(request("admindir"))
adminlo=trim(request("adminlo"))


Set fso = Server.CreateObject("Scripting.FileSystemObject") 
fso.movefile Server.MapPath("data\data.mdb"),Server.MapPath("data\"&dataname&"")'前面是要修改的文件名称，后面是修改后的名称

Set fso = Server.CreateObject("Scripting.FileSystemObject") 
fso.movefile Server.MapPath("admin\adminlogin.asp"),Server.MapPath("admin\"&adminlo&"")'前面是要修改的文件名称，后面是修改后的名称

call renamefolder("data",""&datadir&"")
call renamefolder("admin",""&admindir&"")

set fso = Server.CreateObject("Scripting.FileSystemObject")
set file1 = fso.OpenTextFile(Server.MapPath("inc/config.asp"))
content = file1.ReadAll()
set file1 = nothing
content = replace(content, "data/data.mdb", ""&datadir&"/"&dataname&"")
set file = fso.CreateTextFile(Server.Mappath("inc/config.asp"), true)
file.Write(content)

set fso = Server.CreateObject("Scripting.FileSystemObject")
set file1 = fso.OpenTextFile(Server.MapPath("后台地址.txt"))
content = file1.ReadAll()
set file1 = nothing
content = replace(content, "admin/adminlogin.asp", ""&admindir&"/"&adminlo&"")
set file = fso.CreateTextFile(Server.Mappath("后台地址.txt"), true)
file.Write(content)

set fso = createobject("scripting.filesystemobject")
fso.deletefile(server.mappath("setup.asp"))
set fso = nothing
response.write"<script language=javascript>alert('配置成功，点击进入后台，请牢记后台地址，同时后台地址已经在文本文件“后台地址.txt”中修改!');"
response.write"this.location.href='"&admindir&"/"&adminlo&"'</script>"
%>
<%
else
%>
<TABLE cellSpacing=0 cellPadding=0 width=980 align=center bgColor=#ffffff 
border=0>
  <TBODY>
    <TR> 
      <TD width="10">&nbsp;</TD>
      <TD width="960" height="25"><form name="form1" method="post" action="setup.asp">
          <p>&nbsp;</p>
          <table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#33CC99">
            <tr bgcolor="#FFFFFF"> 
              <td width="21%" height="30" align="right">数据库文件夹名称:</td>
              <td width="79%"> &nbsp; <input name="datadir" type="text" id="datadir" value="data">
                <font color="#FF0000">全部为字母，字母a - z 的组合,<font color="#0000FF">如：bijie520</font></font></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td height="30" align="right">数据库名称:</td>
              <td> &nbsp; <input name="dataname" type="text" id="dataname" value="data.asa">
                <font color="#FF0000">后缀名为<font color="#0000FF">asp</font>或<font color="#0000FF">asa</font>，必须是 
                <font color="#0000FF">.asp</font> 或 <font color="#0000FF">.asa 
                </font>结尾，如：<font color="#0000FF">gggsdrwes.asa</font></font></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td height="30" align="right">后台目录名称:</td>
              <td> &nbsp; <input name="admindir" type="text" id="admindir" value="admin">
                <font color="#FF0000">全部为字母格式，字母a_z 的组合，如：<font color="#0000FF">asdfasfaw</font></font></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td height="30" align="right">后台登陆文件:</td>
              <td> &nbsp; <input name="adminlo" type="text" id="adminlo" value="adminlogin.asp">
                <font color="#FF0000">后缀名为<font color="#0000FF">asp</font>，必须是 
                <font color="#0000FF">.asp</font> 结尾，如：<font color="#0000FF">weasdfasdf.asp</font></font></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td height="30">&nbsp;</td>
              <td> <input name="modi" type="hidden" id="modi" value="yes"> <input type="submit" name="Submit" value=" 确 定 ， 这 就 提 交  "> 
              </td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td height="47" colspan="2"><div align="center"><strong><font color="#0000FF">友情提醒：为了系统安全，建议修改相关名称，所有名称修改均方便你记忆，且不要修改为本示例名称，其中表格里面的是系统初始名称</font></strong></div></td>
            </tr>
          </table>
        <p>&nbsp;</p></form></TD>
      <TD width="10">&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
<%
end if
%>
 </TBODY>
</HTML>