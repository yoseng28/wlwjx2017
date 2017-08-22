<!--#include file="adminconn.asp"-->
<!--#include file="Fckeditor/Fckeditor.Asp" -->
<%
Dim ProductName,ProductBigClass,ProductSmallClass,ProductPrice,ProductPrice1,ProductPrice2,ProductPrice3,ProductPrice4,ProductPrice5,ProductPrice6,ProductType,ProductImages,ProductUser,ok,sContent,i,Count,id,content,ors
if request("modi")="yes" then
id=request("id")
ProductName=request("ProductName")
ProductBigClass=request("BigClassName")
ProductSmallClass=request("SmallClassName")
ProductPrice=request("ProductPrice")
ProductPrice1=request("ProductPrice1")
ProductPrice2=request("ProductPrice2")
ProductPrice3=request("ProductPrice3")
ProductPrice4=request("ProductPrice4")
ProductPrice5=request("ProductPrice5")
ProductPrice6=request("ProductPrice6")
ProductType=request("ProductType")
ProductImages=request("firstImageName")
ProductUser=request("user")
ok=request("ok")
	sContent = ""
	For i = 1 To Request.Form("d_content").Count
		sContent = sContent & Request.Form("d_content")(i)
	Next

'判断输入数据是否完整
if ProductName="" or ProductBigClass="" or ProductPrice1="" or ProductType="" or sContent="" then  
response.write"<div align=center><font color=red size=+3><br><br><strong>※出错了※</strong></font>"
response.write"<br><br><font color=green>你的输入有误,没有添加成功。</font>"
response.write"<br><br>请检查<font color=red>*</font>的地方是否全部填写。"
response.write"<br><br><a href=javascript:history.back()>返回继续填</a></div>"  
response.end 
end if

set rs=server.createobject("adodb.recordset")
sql="select * from Product where id="&id
rs.open sql,conn,1,3
rs("ProductName")=ProductName
rs("ProductPrice")=ProductPrice
rs("ProductPrice1")=ProductPrice1
rs("ProductPrice2")=ProductPrice2
rs("ProductPrice3")=ProductPrice3
rs("ProductPrice4")=ProductPrice4
rs("ProductPrice5")=ProductPrice5
rs("ProductPrice6")=ProductPrice6
rs("ProductBigClass")=ProductBigClass
rs("ProductSmallClass")=ProductSmallClass
rs("ProductType")=ProductType
rs("ProductImages")=ProductImages
rs("ProductUser")=ProductUser
rs("ProductContent")=sContent
if ok<>"" then rs("ok") = ok
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('修改成功！');" & Chr(13)
		response.write "window.document.location.href='Product_info.asp';"&Chr(13)
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
<meta name="renderer" content="webkit"> 
<meta name="renderer" content="ie-comp"> 
<meta name="renderer" content="ie-stand">

<link href="images/style.css" rel="stylesheet" type="text/css">
<title>添加内容</title>
</head>

<body leftmargin="0" topmargin="0">
<% 
id=request("id")
Set ors=Server.CreateObject("ADODB.RecordSet") 
sql="select * from Product where id="&id
ors.Open sql,conn,1,1
if ors.eof and ors.bof then
response.Write("没有记录")
else
content = ors("ProductContent")
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>　</td>
  </tr>
</table>
<table width="780" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <form name="addNEWS" method="post" action="Product_infomodi.asp?modi=yes">
    <tr> 
      <td height="25" colspan="3" align="center" background="images/b1.gif"><span class="style1"><font color="#FFFFFF"><strong>修 
        改 图片 内 容</strong></font></span> </td>
    </tr>
    <tr> 
      <td width="85" height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>标题名称：</td>
      <td colspan="2" bgcolor="#f9f9f9"> <input name="ProductName" type="text" class="input" id="ProductName" value="<%=ors("ProductName")%>" size="30"></td>
    </tr>
    <tr> 
      <td height="-2" bgcolor="#f9f9f9"><font color="#FF0000">*</font>图片类型：</td>
      <td colspan="2" bgcolor="#f9f9f9"> 
        <%
	if session("aleave")="check" then
		response.write ors("ProductBigClass") & "<input name='BigClassName' type='hidden' value='" & ors("ProductBigClass") & "'>&gt;&gt;"
	else		
        sql = "select * from BigClass where ClassType='Product.asp'"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
		else
		%>
        <select name="BigClassName" onChange="changelocation(document.addNEWS.BigClassName.options[document.addNEWS.BigClassName.selectedIndex].value)" size="1">
          <%
		    do while not rs.eof
			%>
          <option <% if rs("BigClassName")=ors("ProductBigClass") then response.Write("selected") end if%> value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
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
		response.write ors("ProductSmallClass") & "<input name='SmallClassName' type='hidden' value='" & ors("ProductSmallClass") & "'>"
	else
	%>
        <select name="SmallClassName">
          <option value="" <%if ors("ProductSmallClass")="" then response.write "selected"%>>不指定小类</option>
          <%
			sql="select * from SmallClass where BigClassName='" & ors("ProductBigClass") & "'" 
			rs.open sql,conn,1,1 
			if not(rs.eof and rs.bof) then  
				do while not rs.eof %>
          <option <% if rs("SmallClassName")=ors("ProductSmallClass") then response.Write("selected") end if%> value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
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
    </tr> <tr> 
      <td height="1" bgcolor="#f9f9f9"><font color="#FF0000">*</font>字段一：</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductPrice5" type="text" id="ProductPrice5" size="30" value="<%=ors("ProductPrice5")%>">
        如：<b>老师简介（字数220）</b></td>
    </tr>

    <tr> 
      <td height="-1" bgcolor="#f9f9f9"><font color="#FF0000">*</font>字段二：</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductType" type="text" id="ProductPrice" size="30" value="<%=ors("ProductType")%>">
        如：学历</td>
    </tr>
    <tr> 
      <td height="0" bgcolor="#f9f9f9"><font color="#FF0000">*</font>字段三：</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductPrice1" type="text" id="ProductType" size="30" value="<%=ors("ProductPrice1")%>"> 
      	如：学位</td>
    </tr>

    <tr> 
      <td bgcolor="#f9f9f9"><font color="#FF0000">*</font>字段四：</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductPrice2" type="text" id="ProductType0" size="30" value="<%=ors("ProductPrice2")%>"> 
		如：职称</td>
    </tr>
     <tr> 
      <td height="0" bgcolor="#f9f9f9"><font color="#FF0000">*</font>字段五：</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductPrice3" type="text" id="ProductType" size="30" value="<%=ors("ProductPrice3")%>"> 
      	如：主讲课程</td>
    </tr>

    <tr> 
      <td bgcolor="#f9f9f9"><font color="#FF0000">*</font>字段六：</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductPrice4" type="text" id="ProductType0" size="30" value="<%=ors("ProductPrice4")%>"> 
		如：视频播放地址</td>
    </tr>
    <tr> 
      <td bgcolor="#f9f9f9">&nbsp;字段七：</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductPrice6" type="text" id="ProductType0" size="30" value="<%=ors("ProductPrice6")%>"> 
		如：排序专用，请填写序号 序号越小排名靠前</td>
    </tr>

    <tr> 
      <td height="20" bgcolor="#f9f9f9">缩略图片地址</td>
      <td width="189" bgcolor="#f9f9f9"> <input name="firstImageName" type="text" id="firstImageName2" value="<%=ors("ProductImages")%>" size="30"> 
      </td>
      <td width="482" bgcolor="#f9f9f9"> <iframe style="top:2px" ID="UploadFiles" src="shangchuan_index.asp?picID=1" frameborder=0 scrolling=no width="320" height="25"></iframe>
		<b>视频上传也在此处（小于30M），上传后剪切至视频播放地址处（字段六）在上传视频缩图</b></td>
    </tr>
    <tr> 
      <td height="25" valign="top" bgcolor="#f9f9f9"><font color="#FF0000">*</font>详细介绍：</td>
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
      <td colspan="2" bgcolor="#f9f9f9"> <input name="user" type="text" class="input" id="user" value="<%=ors("ProductUser")%>" size="30"></td>
    </tr>
    <tr> 
      <td height="25" bgcolor="#f9f9f9">设为首页滚动
：</td>
      <td colspan="2" bgcolor="#f9f9f9"> <input type="radio" value="True" <%if ors("ok")=True then Response.Write "checked"%>  name="ok">
        是 
        <input type="radio" value="False" <%if ors("ok")=False then Response.Write "checked"%> name="ok">
        否 　</td>
    </tr>
    <tr> 
      <td height="30" colspan="3" align="center" bgcolor="#f9f9f9"> <input type="submit" name="Submit" value="提交" class="input"> 
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
