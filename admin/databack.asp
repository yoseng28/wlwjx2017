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
response.write "<script>alert(""���ܽ���fso������ȷ����Ŀռ�֧��fso:��"");history.back();</script>"
response.end
	end if
 if objfso.Folderexists(backf) then
 else
  Set fy=objfso.CreateFolder(backf)
  end if
 objfso.copyfile currf,backf& "\"& backfy
response.write "<script>alert(""�������ݿ�ɹ�"");history.back();</script>"
end if 
%>
<head>
<title>��������</title>
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
            <font color="#FFFFFF"><strong>�������ݿ�</strong></font></td>
        </tr>
        <tr align="center" bgcolor="#f9f9f9"> 
          <td class=tdc height=22 colspan="2">��Ŀռ�ֻ��֧��fso�ſ��Խ������²�����������ֻ���ֶ�����</td>
        </tr>
        <tr bgcolor="#f9f9f9"> 
          <td width="30%" align="right" class=tdc>���ݿ�·����</td>
          <td class=tdc> <input type="text" name="currf" size="20" value="<%=admindataurl%>">
            �벻Ҫ�޸ĸ���</td>
        </tr>
        <tr bgcolor="#f9f9f9"> 
          <td width="30%" align="right" class=tdc>��������Ŀ¼��</td>
          <td class=tdc> <input type="text" name="backf" size="20" value="databack">
            �벻Ҫ�޸ĸ��� </td>
        </tr>
        <tr bgcolor="#f9f9f9"> 
          <td width="30%" align="right" class=tdc>���ݿ����ƣ�</td>
          <td class=tdc> <input type="text" name="backfy" size="20" value="<% =date() %>.mdb"> 
          </td>
        </tr>
        <tr bgcolor="#f9f9f9">
          <td width="30%" align="right" class=tdc>&nbsp;</td>
          <td bgcolor="#f9f9f9" class=tdc>ע�⣺Ϊ����İ�ȫ���ڱ������ݿ���뵽 <a href="http://127.0.0.1/admin/Edit_Admin_UploadFile.asp?sUploadDir=databack" target="main">�������ݿ����</a> 
            �������ݿ⣬��ɾ���������ݿ⡣</td>
        </tr>
        <tr bgcolor="#f9f9f9"> 
          <td width="30%" align="right" class=tdc>��</td>
          <td class=tdc> <input type="submit" name="Submit" value="�ύ" class=bdtj> 
            <input type="reset" name="Submit2" value="����" class=bdtj> </td>
        </tr>
      </form>
    </table>
  </center>
  </div>
  </body>
</html>