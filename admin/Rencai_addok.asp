<!--#include file="adminconn.asp"-->
<%
Dim zw,zprs,lxr,lxdz,lxdh,zwsm,ok,sContent,i
zw=request("zw")
zprs=request("zprs")
lxr=request("lxr")
lxdz=request("lxdz")
lxdh=request("lxdh")

	zwsm = ""
	For i = 1 To Request.Form("d_content").Count
		zwsm = zwsm & Request.Form("d_content")(i)
	Next

'判断输入数据是否完整
if zw="" or zprs="" or zwsm="" then  
response.write"<div align=center><font color=red size=+3><br><br><strong>※出错了※</strong></font>"
response.write"<br><br><font color=green>你的输入有误,没有添加成功。</font>"
response.write"<br><br>请检查<font color=red>*</font>的地方是否全部填写。"
response.write"<br><br><a href=javascript:history.back()>返回继续填</a></div>"  
response.end 
end if

set rs=server.createobject("adodb.recordset")
sql="select * from rencai where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("zw")=zw
rs("zwsm")=zwsm
rs("zprs")=zprs
rs("lxr")=lxr
rs("lxdz")=lxdz
rs("lxdh")=lxdh
rs.update
rs.close
set rs=nothing
conn.close  
set conn=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('提交成功！');" & Chr(13)
		response.write "window.document.location.href='rencai_info.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
%>