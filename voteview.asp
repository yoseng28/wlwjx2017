<html>
<head>
<title><% =sitetitle %></title>
<link href="css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<!-- #include file="conn.asp"-->
<table width="90%" border="0" align="center" cellpadding="4" cellspacing="0" bordercolorlight="#cccccc" bordercolordark="#f7f7f7" class="inputt">
  <%
on error resume next
dim id,rs,rs1,rs2
id=request.querystring("id")
id=replace(id,"'","")
id=left(id,8)
set rs=server.createobject("adodb.recordset")
rs.open "select title from V_title where id=cint('"&id&"')",conn,1,1
if not rs.eof then
set rs2=server.createobject("adodb.recordset")
rs2.open "select sum(vcount) as asum from V_vote where lid=cint('"&id&"')",conn,1,1
if not rs2.eof then
%>
  <tr>
<td height="30" colspan="4" align="center"><font color="#0066CC" size="3"><b><%=rs("title")%></b></font><br><br>目前共有<font color="#ff0000"><%=rs2("asum")%></font>人参予了投票</td>
</tr>

<%
set rs1=server.createobject("adodb.recordset")
rs1.open "select * from V_vote where lid=cint('"&id&"')",conn,1,1
if not rs1.eof then
do while not rs1.eof
%>
<tr>
<td>
<%=rs1("cont")%>
</td>
<td>
<img src="vb1.gif" border="0" height="10" width="<%=(rs1("vcount")/rs2("asum"))*50%>">
</td>
<td>
<font color="#A2A2A2"><%=rs1("vcount")%>票</font>
</td>
<td>
<font color="#A2A2A2"><%=formatnumber((rs1("vcount")/rs2("asum"))*100,2)%>%</font>
</td>
</tr>
<%
rs1.movenext
loop
end if
rs1.close
%>
<tr>
    <td bgcolor="#f7f7f7" height="30" colspan="4" align="center"><% =sitetitle %></td>
</tr>
</table>
<%
rs2.close
end if
%>
<%
end if
rs.close
%>
</body>
</html>