<!--#include file="adminconn.asp"-->
<%
Dim page,pages,j
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
<style type="text/css">
<!--
.style1 {
	color: #FF0000;
	font-weight: bold;
}
-->
</style>
</head>

<body >
<table width="80%" height="56" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="25"><div align="center" class="style1">������Ϣ����</div></td>
  </tr>
</table>
<table width="95%" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <tr> 
    <td width="6%" height="25" align="center" bgcolor="#CCCCCC" background="images/b1.gif">ID</td>
    <td width="62%" align="center" background="images/b1.gif" bgcolor="#CCCCCC">����</td>
    <td width="16%" align="center" background="images/b1.gif" bgcolor="#CCCCCC">���</td>
    <td width="16%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">����</td>
  </tr>
  <%
	page=clng(request("page"))
	Set rs=Server.CreateObject("ADODB.RecordSet") 
	sql="select * from about order by id desc"
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
  <tr bgcolor="#FFFFFF"> 
    <td height="22" align="center"><%=rs("id")%></td>
    <td>�� <a href="../about.asp?type=<%=rs("id")%>" target="_blank" title="<%=rs("title")%>"><%=left(rs("title"),12)%></a></td>
    <td>&nbsp;<%=rs("bigclassname")%></td>
    <td align="center" bgcolor="#FFFFFF"><a href="about_infomodi.asp?id=<%=rs("id")%>">�޸�</a> 
      <a href="about_infodel.asp?id=<%=rs("id")%>" onClick="return ConfirmDelnr();">ɾ��</a></td>
  </tr>
  <%
		rs.movenext
		if rs.eof then exit for
		next                                                       
	%>
</table>
  
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>��</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
<form method=Post action="about_info.asp">  
      <td height="30" align="center"> 
<%
	if Page<2 then      
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=about_info.asp?page=1>��ҳ</a>&nbsp;"
    response.write "<a href=about_info.asp?page=" & Page-1 & ">��һҳ</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=about_info.asp?page=" & (page+1) & ">"
    response.write "��һҳ</a> <a href=about_info.asp?page="&rs.pagecount&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#FF0000'>"&rs.recordcount&"</font></b>����¼ <b>"&rs.pagesize&"</b>����¼/ҳ"
   response.write " ת����<input type='text' name='page' size=4 maxlength=10 class=input value="&page&">"
   response.write " <input class=input type='submit'  value=' Goto '  name='cndok'></span></p>"     
%>
      </td></form>
  </tr>
</table>
<% 
end if
rs.close
set rs=nothing
%>
</body>
</html>
