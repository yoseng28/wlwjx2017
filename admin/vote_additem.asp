<!--#include file="adminconn.asp"-->
<%
  dim id,cont
%>
<%
id=request.form("id")
cont=server.htmlencode(request.form("cont"))
if id<>"" and cont<>"" then
conn.execute("insert into V_vote (lid,cont) values ("&id&",'"&cont&"')")
response.redirect "vote_item.asp?id="&id
else
response.write "<script>alert('投票选项不能为空');history.back();</script>"
response.end
end if
%>