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
call alertbox("ϵͳû���ҵ�ָ����·��[" & sFolder & "]!",2)     
end if     
set fso = nothing     
end function    
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>
<title>ϵͳ��ʼ����</title>
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
fso.movefile Server.MapPath("data\data.mdb"),Server.MapPath("data\"&dataname&"")'ǰ����Ҫ�޸ĵ��ļ����ƣ��������޸ĺ������

Set fso = Server.CreateObject("Scripting.FileSystemObject") 
fso.movefile Server.MapPath("admin\adminlogin.asp"),Server.MapPath("admin\"&adminlo&"")'ǰ����Ҫ�޸ĵ��ļ����ƣ��������޸ĺ������

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
set file1 = fso.OpenTextFile(Server.MapPath("��̨��ַ.txt"))
content = file1.ReadAll()
set file1 = nothing
content = replace(content, "admin/adminlogin.asp", ""&admindir&"/"&adminlo&"")
set file = fso.CreateTextFile(Server.Mappath("��̨��ַ.txt"), true)
file.Write(content)

set fso = createobject("scripting.filesystemobject")
fso.deletefile(server.mappath("setup.asp"))
set fso = nothing
response.write"<script language=javascript>alert('���óɹ�����������̨�����μǺ�̨��ַ��ͬʱ��̨��ַ�Ѿ����ı��ļ�����̨��ַ.txt�����޸�!');"
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
              <td width="21%" height="30" align="right">���ݿ��ļ�������:</td>
              <td width="79%"> &nbsp; <input name="datadir" type="text" id="datadir" value="data">
                <font color="#FF0000">ȫ��Ϊ��ĸ����ĸa - z �����,<font color="#0000FF">�磺bijie520</font></font></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td height="30" align="right">���ݿ�����:</td>
              <td> &nbsp; <input name="dataname" type="text" id="dataname" value="data.asa">
                <font color="#FF0000">��׺��Ϊ<font color="#0000FF">asp</font>��<font color="#0000FF">asa</font>�������� 
                <font color="#0000FF">.asp</font> �� <font color="#0000FF">.asa 
                </font>��β���磺<font color="#0000FF">gggsdrwes.asa</font></font></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td height="30" align="right">��̨Ŀ¼����:</td>
              <td> &nbsp; <input name="admindir" type="text" id="admindir" value="admin">
                <font color="#FF0000">ȫ��Ϊ��ĸ��ʽ����ĸa_z ����ϣ��磺<font color="#0000FF">asdfasfaw</font></font></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td height="30" align="right">��̨��½�ļ�:</td>
              <td> &nbsp; <input name="adminlo" type="text" id="adminlo" value="adminlogin.asp">
                <font color="#FF0000">��׺��Ϊ<font color="#0000FF">asp</font>�������� 
                <font color="#0000FF">.asp</font> ��β���磺<font color="#0000FF">weasdfasdf.asp</font></font></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td height="30">&nbsp;</td>
              <td> <input name="modi" type="hidden" id="modi" value="yes"> <input type="submit" name="Submit" value=" ȷ �� �� �� �� �� ��  "> 
              </td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td height="47" colspan="2"><div align="center"><strong><font color="#0000FF">�������ѣ�Ϊ��ϵͳ��ȫ�������޸�������ƣ����������޸ľ���������䣬�Ҳ�Ҫ�޸�Ϊ��ʾ�����ƣ����б���������ϵͳ��ʼ����</font></strong></div></td>
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