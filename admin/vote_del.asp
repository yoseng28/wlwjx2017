<!--#include file="adminconn.asp"-->
<%
dim id
id=request.querystring("id")
if cint(id)<>"" then
conn.execute("delete from V_title where id="&id&"")
conn.execute("delete from V_vote where lid="&id&"")
response.write "<script>alert('É¾³ý³É¹¦');location.href='vote_admin.asp';</script>"
response.end
end if
%>