<!--#include file="adminconn.asp"-->
<%
Dim page,judge,judge2,judge3,ipage
%>
<%
if request("sh")="yes" then
sql="select * from guestbook where id="& request.QueryString("id") &" "
set rs=Server.CreateObject("ADODB.recordset") 
rs.open sql,conn,1,3
rs("reply")="-1"
rs.update
rs.close


Response.Write"<script language=JavaScript>"
Response.Write"alert(""��ϲ.������˳ɹ�"");"
Response.Write"window.location='product_guest_admin.asp'"
Response.Write"</script>"
end if
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="images/style.css" type="text/css">
<title>�������</title>
</head>

<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="40" bgcolor="#990000">
<div align="center"><strong><font color="#FFFFFF" size="+2">��ƷԤ������
</font></strong></div></td>
  </tr>
</table>
<TABLE width=95% border=0 align="center" cellPadding=0 
            cellSpacing=0>
  <TBODY>
    <TR> 
      <TD>��</TD>
    </TR>
    <TR> 
      <TD height=496 align="center" vAlign=top> 
        <% 
	page = CLng(request("page"))							'����CLng������pageֵת��ΪLong��
	judge=request("judge")
	judge2=request("judge2")
	judge3=0
sql="select *  from guestbook where classname='product' order by id desc "
Set rs= Server.CreateObject("ADODB.recordset")
rs.Open sql,conn,1,3
%>
        <%
			if not (rs.EOF or rs.BOF) then
			rs.MoveFirst
			rs.PageSize = 8
			If page < 1 Then page = 1 
			If page > rs.PageCount Then page = rs.PageCount  
			rs.AbsolutePage = page
			For ipage = 1 To rs.PageSize%>
        <table width="100%" border="1" cellpadding="2" cellspacing="1" bordercolor="#376ed0" class="itable" height="161">
          <tr> 
            <td width="21%" height="35" class="distance" bgcolor="#EEEEEE"> 
              <strong>������</strong><font color="#000000"><strong>�⣺</strong></font> 
              <strong></span><font color="#663300"><%=rs("name") %></font></strong></td>
            <td width="22%" class="distance" bgcolor="#EEEEEE"><strong>�ύʱ�䣺</strong><b><font color="#663300"><span class="style14"><%=year(rs("time01"))%>-<%=month(rs("time01"))%>-<%=day(rs("time01"))%></span></font></b></td>
            <td width="29%" class="distance" bgcolor="#EEEEEE"><span class="style14"> 
              <% if rs("reply")="0" then%>
              [ <a href="product_guest_admin.asp?id=<%=rs("id")%>&sh=yes">
			<font color="#000000"><strong>δ����</strong></font></a> 
              ] 
              <% end if %>
              [ <a href="guest_delete.asp?id=<%=rs("id")%>" class="blue"  onClick="return ConfirmDelnr();">
			<font color="#000000"><strong>ɾ������Ԥ����Ϣ</strong></font></a> 
              ]</span></td>
            <td width="26%" class="distance" bgcolor="#EEEEEE">��</td>
          </tr>
          <tr> 
            <td height="35" bgcolor="#EAEAEA" class="distance"><strong>
			<font color="#000000">�ա�������</font><font color="#663300"><%=rs("zhifu")%></font></strong></td>
            <td height="21" bgcolor="#EAEAEA" class="distance"><strong><font color="#000000">��ϵ�绰��</font><font color="#663300"><%=rs("oicq")%></font></strong></td>
            <td height="21" bgcolor="#EAEAEA" class="distance" width="29%">
			<strong><font color="#000000">�ʡ����ࣺ</font><font color="#663300"><%=rs("email")%></font></strong></td>
            <td height="21" bgcolor="#EAEAEA" class="distance"><b>������ַ��<font color="#663300"><%=rs("lifang")%></font></b></td>
          </tr>
          <tr> 
            <td height="35" bgcolor="#EAEAEA" class="distance"><strong><font color="#000000">֧����ʽ��</font></strong><b><font color="#663300"><span class="style14"><%=rs("fangshi")%></span></font></b>��</td>
            <td height="27" bgcolor="#EAEAEA" class="distance"><b>Ԥ��������<font color="#663300"><%=rs("jianshu")%></font></b></td>
            <td height="27" bgcolor="#EAEAEA" class="distance" colspan="2"><strong><font color="#000000"><a href="../productview.asp?id=<%=rs("biaotiid")%>" target="_blank">�鿴Ԥ����</a></font><a href="../productview.asp?id=<%=rs("biaotiid")%>">��Ʒ</a></strong></td>
          </tr>
          <tr> 
            <td colspan="4" bgcolor="#f9f9f9" class="distance">
			<span class="style14"><font color="#000000"><strong>�������ݣ�</strong></font><%=rs("guestcontent")%></span></td>
          </tr>
        </table>
        <div align="center">
          <table width="98%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td>��</td>
            </tr>
          </table>
          <%   
				rs.MoveNext      
				If rs.EOF Then Exit For   
				Next   
			%>
        </div>
        <table width="100%" border="0" align="center" cellspacing="0">
          <tr> 
            <td>&nbsp;&nbsp;&nbsp; <div align="center">����<font 
                     color=red><%=rs.recordCount%> </font>������ ҳ����<font 
                        color=red><%=page%> </font>/&nbsp; <%=rs.PageCount%>&nbsp;&nbsp; 
                <%
				If page = 1 Then 
				   Response.Write "<font class='white'>��һҳ����һҳ</font>" 
				End If
				If page <> 1 Then 
				   Response.Write "<a href=product_guest_admin.asp?page=1 class='yellow'>��һҳ</a>" 
				   Response.Write "��<a href=product_guest_admin.asp?page=" & (page - 1) & " class='yellow'>��һҳ</a>" 
				End If
				If page <> RS.PageCount Then 
				   Response.Write "��<a href=product_guest_admin.asp?page=" & (page + 1) & " class='yellow'>��һҳ</a>" 
				   Response.Write "��<a href=product_guest_admin.asp?page=" & RS.PageCount & " class='yellow'>���һҳ</a>" 
				End If 
				If page = RS.PageCount Then 
				   Response.Write "��<font class='white'>��һҳ�����һҳ</font>" 
				End If
			  %>
                &nbsp; </div></td>
          </tr>
        </table>
        <%else%>
        <table width="75%" border="0" align="center" cellspacing="0">
          <tr> 
            <td>Ŀǰû���û����ԣ�</td>
          </tr>
        </table>
        <%end if%>
      </TD>
    </TR>
    <TR> 
      <TD>��</TD>
    </TR>
  </TBODY>
</TABLE>
</body>
</html>
