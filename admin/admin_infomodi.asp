<!--#include file="adminconn.asp"-->
<!--#include file="Fckeditor/Fckeditor.Asp" -->
<%
Dim count,id,ors,content,title,BigClassName,SmallClassName,sContent,firstImageName,user,ok,i
if request("modi")="yes" then
id=request("id")
title=request("title")
BigClassName=request("BigClassName")
SmallClassName=request("SmallClassName")	
	sContent = ""
	For i = 1 To Request.Form("d_content").Count
		sContent = sContent & Request.Form("d_content")(i)
	Next	
firstImageName = Request.form("firstImageName")
user=request("user")
ok=request("ok")

'判断输入数据是否完整
if title="" or BigClassName="" or sContent="" or user="" then  
response.write"<div align=center><font color=red size=+3><br><br><strong>※出错了※</strong></font>"
response.write"<br><br><font color=green>你的输入有误,没有添加成功。</font>"
response.write"<br><br>请检查<font color=red>*</font>的地方是否全部填写。"
response.write"<br><br><a href=javascript:history.back()>返回继续填</a></div>"  
response.end 
end if

set rs=server.createobject("adodb.recordset")
sql="select * from content where id="&id
rs.open sql,conn,1,3
rs("title")=title
rs("content")=sContent
rs("user")=user
rs("BigClassName")=BigClassName
rs("SmallClassName")=SmallClassName
if ok<>"" then rs("ok") = ok
if firstImageName<>"" then rs("firstImageName") = firstImageName else rs("firstImageName") = null end if
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('修改成功！');" & Chr(13)
		response.write "window.document.location.href='admin_info.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
end if
%>
<%
set rs=server.createobject("adodb.recordset")
sql = "select * from SmallClass order by SmallClassID asc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("SmallClassName"))%>","<%= trim(rs("BigClassName"))%>","<%= trim(rs("SmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.addNEWS.SmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.addNEWS.SmallClassName.options[document.addNEWS.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/style.css" rel="stylesheet" type="text/css">
<title>添加内容</title>
</head>

<body leftmargin="0" topmargin="0">
<% 
id=request("id")
Set ors=Server.CreateObject("ADODB.RecordSet") 
sql="select * from content where  id="&id
ors.Open sql,conn,1,1
if ors.eof and ors.bof then
response.Write("没有记录")
else
content = ors("content")
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>　</td>
  </tr>
</table>
<table width="780" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <form name="addNEWS" method="post" action="admin_infomodi.asp?modi=yes">
    <tr align="center" bgcolor="#CCCCCC"> 
      <td height="25" colspan="3" background="images/b1.gif"><font color="#000000"><strong>修改文章
</strong></font> </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="92" height="24" align="right"><font color="#FF0000">*</font>标题：</td>
      <td colspan="2" valign="top"> <input name="title" type="text" class="input" value="<%=ors("title")%>" size="30"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="20" align="right"><font color="#FF0000">*</font>类别：</td>
      <td colspan="2" valign="top"> 
        <%
	if session("aleave")="check" then
		response.write ors("BigClassName") & "<input name='BigClassName' type='hidden' value='" & ors("BigClassName") & "'>&gt;&gt;"
	else		
        sql = "select * from BigClass where ClassType='class.asp'"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
		else
		%>
        <select name="BigClassName" onChange="changelocation(document.addNEWS.BigClassName.options[document.addNEWS.BigClassName.selectedIndex].value)" size="1">
          <%
		    do while not rs.eof
			%>
          <option <% if rs("BigClassName")=ors("BigClassName") then response.Write("selected") end if%> value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
          <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
        </select> 
        <%
	end if
	if session("aleave")="check" then
		response.write ors("SmallClassName") & "<input name='SmallClassName' type='hidden' value='" & ors("SmallClassName") & "'>"
	else
	%>
        <select name="SmallClassName">
          <option value="" <%if ors("SmallClassName")="" then response.write "selected"%>>不指定小类</option>
          <%
			sql="select * from SmallClass where BigClassName='" & ors("BigClassName") & "'" 
			rs.open sql,conn,1,1 
			if not(rs.eof and rs.bof) then  
				do while not rs.eof %>
          <option <% if rs("SmallClassName")=ors("SmallClassName") then response.Write("selected") end if%> value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
          <%
			    	rs.movenext
				loop
			end if
	        rs.close
			%>
        </select> 
        <%
	end if
	%>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="20" align="right">首页幻灯图片：</td>
      <td width="188" valign="top"> 
        <input name="firstImageName" type="text" id="imgpath3" value="<%=ors("firstImageName")%>" size="30"> 
      </td>
      <td width="476" valign="top"> 
        <iframe style="top:2px" ID="UploadFiles" src="shangchuan_index.asp?picID=1" frameborder=0 scrolling=no width="320" height="25"></iframe>
		只用于首页项目幻灯展示</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="right" valign="top"><font color="#FF0000">*</font>内容：</td>
      <td colspan="2" valign="top"> 
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
    <tr bgcolor="#FFFFFF"> 
      <td height="24" align="right"><font color="#FF0000">*</font>发布人：</td>
      <td colspan="2" valign="top">　 
        <input name="user" type="text" class="input" size="30" value="<%=ors("user")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="24" align="right">首页幻灯图片显示：</td>
      <td colspan="2">　 
        <input type="radio" value="True" <%if ors("ok")=True then Response.Write "checked"%>  name="ok">
        是 
        <input type="radio" value="False" <%if ors("ok")=False then Response.Write "checked"%> name="ok">
        否 　</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF"> 
      <td height="35" colspan="3"> <input type="submit" name="Submit" value="提交" class="input"> 
        <input type="hidden" name="id" value="<%=id%>">
        　 
        <input type="reset" name="Submit2" value="重置" class="input">
      </td>
    </tr>
  </form>
</table>
<% End If
ors.close
set ors=nothing
 %>
</body>
</html>
