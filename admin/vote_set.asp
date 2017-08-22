<!--#include file="adminconn.asp"-->
<%
dim id,xindex
id=request.querystring("id")
xindex=request.querystring("setindex")
conn.execute("update V_title set xindex='"&xindex&"' where id=cint('"&id&"')")
response.write"<SCRIPT language=JavaScript>alert('…Ë÷√≥…π¶£°');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end

%>	