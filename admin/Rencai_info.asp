<!--#include file="adminconn.asp"-->
<%
  Dim bigclass,page,pages,j
%>
<% 
bigclass=request("type")
 %>
<script language="JavaScript" type="text/JavaScript">
function ConfirmDelnr()
{
   if(confirm("确定要删除该内容吗？注意：删除后不能恢复！"))
     return true;
   else
     return false;
	 
}

</script>
<html>
<head>
<LINK href="images/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>

<body topmargin="0">
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height=5></td>
  </tr>
</table>
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height=30><div align="center"><strong><font color="#FF0000">招聘管理</font></strong></div></td>
  </tr>
</table>
<table width="95%" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <tr> 
    <td width="5%" height="25" align="center" bgcolor="#CCCCCC" background="images/b1.gif">ID</td>
    <td width="25%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">招聘职位</td>
    <td width="10%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">招聘人数</td>
    <td width="10%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">学历要求</td>
    <td width="9%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">性别要求</td>
    <td width="12%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">年龄要求</td>
    <td width="19%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">发布日期</td>
    <td width="10%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">操作</td>
  </tr>
  <%
page=clng(request("page"))
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from rencai order by id desc"
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
response.Write("没有记录")
else
rs.PageSize=20
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.AbsolutePage=page  
    
for j=1 to rs.PageSize 
%>
  <tr bgcolor="#f9f9f9"> 
    <td height="22" align="center"><%=rs("id")%></td>
    <td bgcolor="#f9f9f9"> &nbsp; <a href="../rencaiview.asp?id=<%=rs("id")%>" target="_blank" title="<%=rs("zw")%>"> 
      <% =rs("zw") %>
      </a></td>
    <td align="center">&nbsp;<%=rs("zprs")%></td>
    <td align="center">&nbsp;<%=rs("lxr")%> </td>
    <td align="center">&nbsp;<%=rs("lxdh")%></td>
    <td align="center">&nbsp;<%=rs("lxdz")%></td>
    <td align="center"><%=rs("infotime")%></td>
    <td align="center"><a href="Rencai_infomodi.asp?id=<%=rs("id")%>">修改</a> <a href="Rencai_infodel.asp?id=<%=rs("id")%>" onClick="return ConfirmDelnr();">删除</a></td>
  </tr>
  <%
rs.movenext
if rs.eof then exit for
next                                                       
%>
</table>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF"> 
      <td height="30" align="center"> 
<%
if bigclass="" then
%>
<%if Page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=Rencai_info.asp?page=1>首页</a>&nbsp;"
    response.write "<a href=Rencai_info.asp?page=" & Page-1 & ">上一页</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=Rencai_info.asp?page=" & (page+1) & ">"
    response.write "下一页</a> <a href=Rencai_info.asp?page="&rs.pagecount&">尾页</a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#FF0000'>"&rs.recordcount&"</font></b>条记录 <b>"&rs.pagesize&"</b>条记录/页"
%>
<%
else
%>
<%if Page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=Rencai_info.asp?type="&bigclass&"&page=1>首页</a>&nbsp;"
    response.write "<a href=Rencai_info.asp?type="&bigclass&"&page=" & Page-1 & ">上一页</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=Rencai_info.asp?type="&bigclass&"&page=" & (page+1) & ">"
    response.write "下一页</a> <a href=Rencai_info.asp?type="&bigclass&"&page="&rs.pagecount&">尾页</a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#FF0000'>"&rs.recordcount&"</font></b>条记录 <b>"&rs.pagesize&"</b>条记录/页"
%>
<%
end if
%>
      </td>
  </tr>
</table>
<% 
end if
rs.close
set rs=nothing
%>

</body>
</html>
