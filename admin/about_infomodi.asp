<!--#include file="adminconn.asp"-->
<!--#include file="Fckeditor/Fckeditor.Asp" -->
<%
Dim id,ors,Content,title,sContent,i,bigclassname  '�������
if request("modi")="yes" then
id=request("id")
bigclassname=request("bigclassname")
title=request("title")
	sContent = ""
	For i = 1 To Request.Form("d_content").Count
		sContent = sContent & Request.Form("d_content")(i)
	Next
	
'�ж����������Ƿ�����
if title="" or sContent="" or bigclassname="" then  
response.write"<div align=center><font color=red size=+3><br><br><strong>�������ˡ�</strong></font>"
response.write"<br><br><font color=green>�����������,û����ӳɹ���</font>"
response.write"<br><br>����<font color=red>*</font>�ĵط��Ƿ�ȫ����д��"
response.write"<br><br><a href=javascript:history.back()>���ؼ�����</a></div>"  
response.end 
end if

set rs=server.createobject("adodb.recordset")
sql="select * from about where id="&id
rs.open sql,conn,1,3
rs("title")=title
rs("bigclassname")=bigclassname
rs("content")=sContent
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('�޸ĳɹ���');" & Chr(13)
		response.write "window.document.location.href='about_info.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/style.css" rel="stylesheet" type="text/css">
<title>�޸�����</title>
<SCRIPT language=JavaScript>
<!--
function CheckForm()
{
	if (document.addNEWS.title.value=="") {
		alert("ע�⣺���ⲻ��Ϊ��.");
		document.addNEWS.title.focus();
		return false;
	}
		if (document.addNEWS.d_content.value=="") {
		alert("ע�⣺���ݲ���Ϊ��.");
		document.addNEWS.d_content.focus();
		return false;
	}
	return true;
}
-->
</SCRIPT>
</head>

<body leftmargin="0" topmargin="0">
<% 
id=request("id")
Set ors=Server.CreateObject("ADODB.RecordSet") 
sql="select * from about where  id="&id
ors.Open sql,conn,1,1
if ors.eof and ors.bof then
response.Write("û�м�¼")
else
Content=ors("content")
%>
<table width="98%" height="63" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>
<table width="780" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <form name="addNEWS" method="post" action="about_infomodi.asp?modi=yes" onSubmit="return CheckForm();">
    <tr align="center" bgcolor="#FFFFEE"> 
      <td height="25" colspan="2" background="images/b1.gif"><font color="#FFFFFF"><strong>�޸�����</strong></font></td>
    </tr>
    <tr bgcolor="#f9f9f9"> 
      <td width="85" height="25" align="right"><font color="#FF0000">*</font>���⣺</td>
      <td width="678" valign="top" bgcolor="#f9f9f9"> 
        <input name="title" type="text" class="input" value="<%=ors("title")%>" size="20"> 
        <% if ors("abouttype")=1 then %>
        <strong><font color="#FF0000"> ע�⣺</font></strong>�޸ĸñ�����뵽<a href="ClassManage.asp" target="_blank"><font color="#FF0000">��Ŀ����</font></a>�����޸����ƣ�����ǰ̨�޷�������ʾ 
        <% end if %>
      </td>
    </tr>
    <tr bgcolor="#f9f9f9">
      <td width="85" height="25" align="right"><font color="#FF0000">*</font>�������ƣ�</td>
      <td valign="top" bgcolor="#f9f9f9">
        <select name="BigClassName" onChange="changelocation(document.addNEWS.BigClassName.options[document.addNEWS.BigClassName.selectedIndex].value)" size="1">
          <option selected value="<%=trim(ors("BigClassName"))%>"><%=trim(ors("BigClassName"))%></option>
          <%
		set rs=server.createobject("adodb.recordset")
        sql = "select * from BigClass where ClassType='about.asp' "
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "���������Ŀ��"
		else
		    selectclass=rs("BigClassName")
		    do while not rs.eof
			%>
          <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
          <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
        </select></td>
    </tr>
    <tr bgcolor="#f9f9f9"> 
      <td align="right" valign="top"><font color="#FF0000">*</font>���ݣ�</td>
      <td valign="top">
        <%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "FCKeditor/"
oFCKeditor.ToolbarSet = "bijie520"
oFCKeditor.Width = "100%"
oFCKeditor.Height = "400"
oFCKeditor.Value = Content
oFCKeditor.Create "d_content"
%>
      </td>
    </tr>
    <tr align="center" bgcolor="#f9f9f9"> 
      <td height="35" colspan="2"> <input type="submit" name="Submit" value="�ύ" class="input"> 
        <input type="hidden" name="id" value="<%=id%>">
        �� 
        <input type="reset" name="Submit2" value="����" class="input"> </td>
    </tr>
  </form>
</table>
<% End If
ors.close
set ors=nothing
 %>
</body>
</html>
