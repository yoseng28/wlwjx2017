<!--#include file="adminconn.asp"-->
<%
dim SmallClassID,SmallClassName
SmallClassID=trim(Request("SmallClassID"))
SmallClassName=trim(Request("SmallClassName"))
if SmallClassID<>"" then
	sql="delete from SmallClass where SmallClassID="&Cint(SmallClassID)&""
	conn.Execute sql
	sql="delete from content where SmallClassName='" & SmallClassName & "'"
	conn.Execute sql
end if
call CloseConn()      
response.redirect "ClassManage.asp"
%>


