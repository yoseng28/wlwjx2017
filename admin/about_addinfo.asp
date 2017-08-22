<!--#include file="adminconn.asp"-->
<!--#include file="Fckeditor/Fckeditor.Asp" -->
<%
dim Content
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=X-UA-Compatible content=IE=EmulateIE7>
<meta http-equiv=“X-UA-Compatible” content=“chrome=1″ >

<meta name="renderer" content="ie-comp"> 
<meta name="renderer" content="ie-stand">

<link rel="stylesheet" href="images/style.css" type="text/css">
<title>添加内容</title>
<SCRIPT language=JavaScript>
<!--
function CheckForm()
{
	if (document.addNEWS.title.value.length == 0) {
		alert("注意：标题不能为空.");
		document.addNEWS.title.focus();
		return false;
	}
		if (document.addNEWS.sContent.value.length == 0) {
		alert("注意：内容不能为空.");
		document.addNEWS.sContent.focus();
		return false;
	}
	return true;
}
-->
</SCRIPT>
</head>

<body leftmargin="0" topmargin="0">
<table width="98%" height="63" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td>　</td>
  </tr>
</table>
<table width="780" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <form name="addNEWS" method="post" action="about_addinfook.asp" onsubmit="return CheckForm();">
    <tr> 
      <td height="25" colspan="2" align="center" background="images/b1.gif"><font color="#FFFFFF"><strong>添加基本信息</strong></font></td>
    </tr>
    <tr bgcolor="#f9f9f9"> 
      <td width="85" height="25"><font color="#FF0000">*</font>标题：</td>
      <td width="678"> 
        <input name="title" type="text" class="input" size="30">
      </td>
    </tr>
    <tr bgcolor="#f9f9f9">
      <td width="85" height="25"><font color="#FF0000">*</font>大类名称</td>
      <td>
        <%
		set rs=server.createobject("adodb.recordset")
        sql = "select * from BigClass where ClassType='about.asp' "
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
		else
		%>
        <select name="BigClassName" onChange="changelocation(document.addNEWS.BigClassName.options[document.addNEWS.BigClassName.selectedIndex].value)" size="1">
          <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
          <%
		    selectclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
          <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
          <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
        </select>
      </td>
    </tr>
    <tr bgcolor="#f9f9f9"> 
      <td height="25" valign="top"><font color="#FF0000">*</font>内容：</td>
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
    <tr bgcolor="#f9f9f9"> 
      <td height="30" colspan="2" align="center"> <input type="submit" name="Submit" value="提交" class="input">
        　 
        <input type="reset" name="Submit2" value="重置" class="input"> </td>
    </tr>
  </form>
</table>
</body>
</html>