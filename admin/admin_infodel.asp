<!--#include file="adminconn.asp"-->
<%
dim action,id
action=request("action")
if action="del" then
id=request("id")
set rs=server.CreateObject("ADODB.RecordSet")
sql="delete * from content where id in ("& id &")"
conn.execute(sql)
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('成功删除！');" & Chr(13)
		response.write "window.document.location.href='admin_info.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End

else

set rs=server.CreateObject("ADODB.RecordSet")
rs.open "delete * from content where id="&request.QueryString("id"),conn,1
set rs=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('成功删除！');" & Chr(13)
		response.write "window.document.location.href='admin_info.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End

end if
%>
