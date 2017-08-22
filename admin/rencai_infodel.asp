<!--#include file="adminconn.asp"-->
<%
set rs=server.CreateObject("ADODB.RecordSet")
rs.open "delete * from rencai where id="&request.QueryString("id"),conn,1
set rs=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('³É¹¦É¾³ý£¡');" & Chr(13)
		response.write "window.document.location.href='rencai_info.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
%>
