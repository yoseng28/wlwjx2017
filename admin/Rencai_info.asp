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
   if(confirm("ȷ��Ҫɾ����������ע�⣺ɾ�����ָܻ���"))
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
    <td height=30><div align="center"><strong><font color="#FF0000">��Ƹ����</font></strong></div></td>
  </tr>
</table>
<table width="95%" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <tr> 
    <td width="5%" height="25" align="center" bgcolor="#CCCCCC" background="images/b1.gif">ID</td>
    <td width="25%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">��Ƹְλ</td>
    <td width="10%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">��Ƹ����</td>
    <td width="10%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">ѧ��Ҫ��</td>
    <td width="9%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">�Ա�Ҫ��</td>
    <td width="12%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">����Ҫ��</td>
    <td width="19%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">��������</td>
    <td width="10%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">����</td>
  </tr>
  <%
page=clng(request("page"))
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from rencai order by id desc"
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
response.Write("û�м�¼")
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
    <td align="center"><a href="Rencai_infomodi.asp?id=<%=rs("id")%>">�޸�</a> <a href="Rencai_infodel.asp?id=<%=rs("id")%>" onClick="return ConfirmDelnr();">ɾ��</a></td>
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
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=Rencai_info.asp?page=1>��ҳ</a>&nbsp;"
    response.write "<a href=Rencai_info.asp?page=" & Page-1 & ">��һҳ</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=Rencai_info.asp?page=" & (page+1) & ">"
    response.write "��һҳ</a> <a href=Rencai_info.asp?page="&rs.pagecount&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#FF0000'>"&rs.recordcount&"</font></b>����¼ <b>"&rs.pagesize&"</b>����¼/ҳ"
%>
<%
else
%>
<%if Page<2 then      
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=Rencai_info.asp?type="&bigclass&"&page=1>��ҳ</a>&nbsp;"
    response.write "<a href=Rencai_info.asp?type="&bigclass&"&page=" & Page-1 & ">��һҳ</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=Rencai_info.asp?type="&bigclass&"&page=" & (page+1) & ">"
    response.write "��һҳ</a> <a href=Rencai_info.asp?type="&bigclass&"&page="&rs.pagecount&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#FF0000'>"&rs.recordcount&"</font></b>����¼ <b>"&rs.pagesize&"</b>����¼/ҳ"
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
