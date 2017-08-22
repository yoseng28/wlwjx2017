<!--#include file="adminconn.asp"-->
<!--#include file="Fckeditor/Fckeditor.Asp" -->
<%
Dim DownName,DownBigClass,DownSmallClass,Downdaxiao,Downlink,DownUser,ok,sContent,i,Count,id,content,ors
if request("modi")="yes" then
id=request("id")
DownName=request("DownName")
DownBigClass=request("BigClassName")
DownSmallClass=request("SmallClassName")
Downdaxiao=request("Downdaxiao")
Downlink=request("firstImageName")
DownUser=request("user")
ok=request("ok")
	sContent = ""
	For i = 1 To Request.Form("d_content").Count
		sContent = sContent & Request.Form("d_content")(i)
	Next

'判断输入数据是否完整
if DownName="" or DownBigClass="" or sContent="" then  
response.write"<div align=center><font color=red size=+3><br><br><strong>※出错了※</strong></font>"
response.write"<br><br><font color=green>你的输入有误,没有添加成功。</font>"
response.write"<br><br>请检查<font color=red>*</font>的地方是否全部填写。"
response.write"<br><br><a href=javascript:history.back()>返回继续填</a></div>"  
response.end 
end if

set rs=server.createobject("adodb.recordset")
sql="select * from download where id="&id
rs.open sql,conn,1,3
rs("DownName")=DownName
rs("DownBigClass")=DownBigClass
rs("DownSmallClass")=DownSmallClass
rs("Downlink")=Downlink
rs("Downdaxiao")=Downdaxiao
rs("DownUser")=DownUser
if ok<>"" then rs("ok") = ok
rs("DownContent")=sContent
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('修改成功！');" & Chr(13)
		response.write "window.document.location.href='Download_info.asp';"&Chr(13)
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
sql="select * from download where id="&id
ors.Open sql,conn,1,1
if ors.eof and ors.bof then
response.Write("没有记录")
else
content = ors("DownContent")
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>　</td>
  </tr>
</table>
<table width="780" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <form name="addNEWS" method="post" action="Download_infomodi.asp?modi=yes">
    <tr> 
      <td height="25" colspan="3" align="center" background="images/b1.gif"><span class="style1"><font color="#FFFFFF"><strong>修 
        改 下 载 内 容</strong></font></span> </td>
    </tr>
    <tr> 
      <td width="86" height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>文件名称：</td>
      <td colspan="2" bgcolor="#f9f9f9"> <input name="DownName" type="text" class="input" id="DownName" value="<%=ors("DownName")%>" size="30"></td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>文件分类：</td>
      <td colspan="2" bgcolor="#f9f9f9"> 
        <%
	if session("aleave")="check" then
		response.write ors("DownBigClass") & "<input name='BigClassName' type='hidden' value='" & ors("DownBigClass") & "'>&gt;&gt;"
	else		
        sql = "select * from BigClass where ClassType='Download.asp'"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
		else
		%>
        <select name="BigClassName" onChange="changelocation(document.addNEWS.BigClassName.options[document.addNEWS.BigClassName.selectedIndex].value)" size="1">
          <%
		    do while not rs.eof
			%>
          <option <% if rs("BigClassName")=ors("DownBigClass") then response.Write("selected") end if%> value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
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
		response.write ors("DownSmallClass") & "<input name='SmallClassName' type='hidden' value='" & ors("DownSmallClass") & "'>"
	else
	%>
        <select name="SmallClassName">
          <option value="" <%if ors("DownSmallClass")="" then response.write "selected"%>>不指定小类</option>
          <%
			sql="select * from SmallClass where BigClassName='" & ors("DownBigClass") & "'" 
			rs.open sql,conn,1,1 
			if not(rs.eof and rs.bof) then  
				do while not rs.eof %>
          <option <% if rs("SmallClassName")=ors("DownSmallClass") then response.Write("selected") end if%> value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
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
    <tr>
      <td height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>文件大小：</td>
      <td colspan="2" bgcolor="#f9f9f9"> <input name="Downdaxiao" type="text" class="input" id="Downdaxiao" value="<%=ors("Downdaxiao")%>" size="30"> 
        <font color="#FF0000"> 如：1M，250M，.....</font></td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>下载地址：</td>
      <td width="188" bgcolor="#f9f9f9"> <input name="firstImageName" type="text" id="firstImageName2" value="<%=ors("Downlink")%>" size="30"> 
      </td>
      <td width="482" bgcolor="#f9f9f9"> <iframe style="top:2px" ID="UploadFiles" src="shangchuan_index.asp?picID=1" frameborder=0 scrolling=no width="320" height="25"></iframe>
		<font color="#CC0000"><b>在此上传文件（小于30M)</b></font></td>
    </tr>
    <tr> 
      <td height="25" valign="top" bgcolor="#f9f9f9"><font color="#FF0000">*</font>使用说明：</td>
      <td colspan="2" valign="top" bgcolor="#f9f9f9"> 
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
    <tr> 
      <td height="20" bgcolor="#f9f9f9">发布人：</td>
      <td colspan="2" bgcolor="#f9f9f9"> <input name="user" type="text" class="input" id="user" value="<%=ors("DownUser")%>" size="30"></td>
    </tr>
    <tr> 
      <td height="30" colspan="3" align="center" bgcolor="#f9f9f9"> <input type="submit" name="Submit" value="提交" class="input"> 
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
