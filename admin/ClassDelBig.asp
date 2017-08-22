<!--#include file="adminconn.asp"-->
<%
dim BigClassName
BigClassName=trim(Request("BigClassName"))
if BigClassName<>"" then
	sql="delete from BigClass where BigClassName='" & BigClassName & "'"
	conn.Execute sql
	sql="delete from SmallClass where BigClassName='" & BigClassName & "'"
	conn.Execute sql
	sql="delete from content where BigClassName='" & BigClassName & "'"
	conn.Execute sql
end if
call CloseConn()      
response.redirect "ClassManage.asp"
%>


