<!--#include file="adminconn.asp"-->
<%
dim id,iid
iid=request.querystring("iid")
id=request.querystring("id")
if cint(iid)<>"" then
conn.execute("delete from V_vote where id=cint('"&iid&"')")
response.redirect "vote_item.asp?id="&id
end if
%>