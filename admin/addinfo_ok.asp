<!--#include file="adminconn.asp"-->
<%
dim title,BigClassName,SmallClassName,sContent,firstImageName,user,ok,i
title=request("title")
BigClassName=request("BigClassName")
SmallClassName=request("SmallClassName")
 
	sContent = ""
	For i = 1 To Request.Form("d_content").Count
		sContent = sContent & Request.Form("d_content")(i)
	Next
firstImageName = trim(Request.form("firstImageName"))
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
sql="select * from content where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("content")=sContent
rs("user")=user
rs("BigClassName")=BigClassName
rs("SmallClassName")=SmallClassName
if ok<>"" then rs("ok") = ok
if firstImageName<>"" then rs("firstImageName") = firstImageName
rs.update
rs.close
set rs=nothing
conn.close  
set conn=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('提交成功！');" & Chr(13)
		response.write "window.document.location.href='admin_info.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
%>
