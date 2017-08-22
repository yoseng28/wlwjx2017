<!--#include file="adminconn.asp"-->
<!--#include file="Fckeditor/Fckeditor.Asp" -->
<%
Dim id,ors,Content,title,sContent,i,bigclassname  '定义变量
if request("modi")="yes" then
id=request("id")
bigclassname=request("bigclassname")
title=request("title")
	sContent = ""
	For i = 1 To Request.Form("d_content").Count
		sContent = sContent & Request.Form("d_content")(i)
	Next
	
'判断输入数据是否完整
if title="" or sContent="" or bigclassname="" then  
response.write"<div align=center><font color=red size=+3><br><br><strong>※出错了※</strong></font>"
response.write"<br><br><font color=green>你的输入有误,没有添加成功。</font>"
response.write"<br><br>请检查<font color=red>*</font>的地方是否全部填写。"
response.write"<br><br><a href=javascript:history.back()>返回继续填</a></div>"  
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
		response.write "alert('修改成功！');" & Chr(13)
		response.write "window.document.location.href='about_info.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/style.css" rel="stylesheet" type="text/css">
<title>修改内容</title>
<SCRIPT language=JavaScript>
<!--
function CheckForm()
{
	if (document.addNEWS.title.value=="") {
		alert("注意：标题不能为空.");
		document.addNEWS.title.focus();
		return false;
	}
		if (document.addNEWS.d_content.value=="") {
		alert("注意：内容不能为空.");
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
response.Write("没有记录")
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
      <td height="25" colspan="2" background="images/b1.gif"><font color="#FFFFFF"><strong>修改内容</strong></font></td>
    </tr>
    <tr bgcolor="#f9f9f9"> 
      <td width="85" height="25" align="right"><font color="#FF0000">*</font>标题：</td>
      <td width="678" valign="top" bgcolor="#f9f9f9"> 
        <input name="title" type="text" class="input" value="<%=ors("title")%>" size="20"> 
        <% if ors("abouttype")=1 then %>
        <strong><font color="#FF0000"> 注意：</font></strong>修改该标题后请到<a href="ClassManage.asp" target="_blank"><font color="#FF0000">栏目管理</font></a>里面修改名称，否则前台无法正常显示 
        <% end if %>
      </td>
    </tr>
    <tr bgcolor="#f9f9f9">
      <td width="85" height="25" align="right"><font color="#FF0000">*</font>大类名称：</td>
      <td valign="top" bgcolor="#f9f9f9">
        <select name="BigClassName" onChange="changelocation(document.addNEWS.BigClassName.options[document.addNEWS.BigClassName.selectedIndex].value)" size="1">
          <option selected value="<%=trim(ors("BigClassName"))%>"><%=trim(ors("BigClassName"))%></option>
          <%
		set rs=server.createobject("adodb.recordset")
        sql = "select * from BigClass where ClassType='about.asp' "
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
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
      <td align="right" valign="top"><font color="#FF0000">*</font>内容：</td>
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
      <td height="35" colspan="2"> <input type="submit" name="Submit" value="提交" class="input"> 
        <input type="hidden" name="id" value="<%=id%>">
        　 
        <input type="reset" name="Submit2" value="重置" class="input"> </td>
    </tr>
  </form>
</table>
<% End If
ors.close
set ors=nothing
 %>
</body>
</html>
