<!-- #include file="conn.asp"-->
document.write ("<html><head><title>xinhongh</title>");
document.write ("<link href='css.css' rel='stylesheet' type='text/css'>");
document.write ("<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>");
document.write ("</head><body>");
document.write ("<form method='post' action='votes.asp'><table width='170' border='1' cellpadding='2' cellspacing='0' bordercolorlight='#cccccc' class='inputt'>");
<%
on error resume next
dim id,rs,chk
set rs=server.createobject("adodb.recordset")
sql2="select *  from V_title where xindex='Yes' order by id desc "
Set rs= Server.CreateObject("ADODB.recordset")
rs.Open sql2,conn,1,1

if not rs.eof then
chk=rs("types")
id=rs("id")
%>
document.write ("<tr><td height='30'><font color='#ff0000'><%=rs("title")%></font></td></tr><tr><td>");

<%
dim rs1
set rs1=server.createobject("adodb.recordset")
rs1.open "select * from V_vote where lid=cint('"&id&"')",conn,1,1
if not rs1.eof then
do while not rs1.eof
%>
<%
select case rs("types")
case "1" '单选
%>
document.write ("<input type='radio' name='v' value='<%=rs1("id")%>'><%=rs1("cont")%><br>");
<%
case "2" '多选
%>
document.write ("<input type='checkbox' name='v' value='<%=rs1("id")%>'><%=rs1("cont")%><br>");
<%
end select
%>
<%
rs1.movenext
loop
end if
rs1.close
%>
<%
end if
rs.close
%>
document.write ("</td></tr>");
document.write ("<tr><td height='30' align='center'><input type='submit' name='submit' value='投票' class='inputt'>&nbsp;&nbsp;<a href='view.asp?id=<%=id%>' target='_blank'>查看</a><input type='hidden' name='id' value='<%=id%>'><input type='hidden' name='chk' value='<%=chk%>'></td></tr></table></form></body></html>");
