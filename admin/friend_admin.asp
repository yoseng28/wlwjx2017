<!--#include file="adminconn.asp"-->
<%
Dim page,judge,judge2,judge3,ipage
if request("sh")="yes" then
sql="select * from frendlink where id="& request.QueryString("id") &" "
set rs=Server.CreateObject("ADODB.recordset") 
rs.open sql,conn,1,3
rs("reply")="-1"
rs.update
rs.close


Response.Write"<script language=JavaScript>"
Response.Write"alert(""��ϲ.������˳ɹ�"");"
Response.Write"window.location='friend_admin.asp'"
Response.Write"</script>"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="images/style.css" type="text/css">
<title>�������</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFff">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="20"> 
      <div align="center"><strong></strong></div></td>
  </tr>
</table>
<TABLE width=90% border=1 align="center" cellPadding=2 
            cellSpacing=1 bordercolor="#376ed0">
  <TBODY>
    <TR> 
      <TD height="25" background="images/b1.gif"><div align="center"><strong><font color="#0000FF">�������ӹ���</font></strong>��������������<a href="../friendlink.asp" target="_blank"><font color="#FF0000"><strong><font color="#0000FF"></font>�����������</strong></font></a><strong></strong></div></TD>
    </TR>
    <TR> 
      <TD height=496 align="center" vAlign=top>
        <% 
	page = CLng(request("page"))							'����CLng������pageֵת��ΪLong��
	judge=request("judge")
	judge2=request("judge2")
	judge3=0
sql="select *  from frendlink order by id desc "
Set rs= Server.CreateObject("ADODB.recordset")
rs.Open sql,conn,1,3
%>
        <table width="98%" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
          <tr bgcolor="#EAEAEA"> 
            <td width="23%"><strong>�� վ �� ��</strong></td>
            <td width="59%" height="25"><span class="style14"><font color="#000000"><strong>����ַ</strong></font></span></td>
            <td width="18%"><strong>���� ��</strong></td>
          </tr>
		  <%
			if not (rs.EOF or rs.BOF) then
			rs.MoveFirst
			rs.PageSize = 15
			If page < 1 Then page = 1 
			If page > rs.PageCount Then page = rs.PageCount  
			rs.AbsolutePage = page
			For ipage = 1 To rs.PageSize%>
          <tr>
            <td height="30"><font color="#0000FF">&nbsp;</font><strong><%=rs("webname") %></strong><span class="style14"> 
              </span></td>
            <td height="20"><span class="style14">&nbsp;<%=rs("weburl")%>��</span></td>
            <td><span class="style14"> 
              <% if rs("reply")<>"-1" then%>
              [<a href="friend_admin.asp?id=<%=rs("id")%>&sh=yes"><font color="#FF0000">���</font></a>] 
              <% else %>
              <a href="friend_admin.asp?id=<%=rs("id")%>&sh=yes">�����</a> 
              <% end if %>
              [<a href="friend_delete.asp?id=<%=rs("id")%>" class="blue" onClick="return Delete();"><font color="#FF0000">ɾ��</font></a>]</span></td>
          </tr>
		  <%   
				rs.MoveNext      
				If rs.EOF Then Exit For   
				Next   
			%>
        </table> 
        
        <div align="center"> 
          
        </div>
        <table width="100%" border="0" align="center" cellspacing="0">
          <tr> 
            <td>&nbsp;&nbsp;&nbsp; <div align="center">����<font 
                     color=red><%=rs.recordCount%> </font>���������� ҳ����<font 
                        color=red><%=page%> </font>/&nbsp; <%=rs.PageCount%>&nbsp;&nbsp; 
                <%
				If page = 1 Then 
				   Response.Write "<font class='white'>��һҳ����һҳ</font>" 
				End If
				If page <> 1 Then 
				   Response.Write "<a href=friend_admin.asp?page=1 class='yellow'>��һҳ</a>" 
				   Response.Write "��<a href=friend_admin.asp?page=" & (page - 1) & " class='yellow'>��һҳ</a>" 
				End If
				If page <> RS.PageCount Then 
				   Response.Write "��<a href=friend_admin.asp?page=" & (page + 1) & " class='yellow'>��һҳ</a>" 
				   Response.Write "��<a href=friend_admin.asp?page=" & RS.PageCount & " class='yellow'>���һҳ</a>" 
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
            <td>Ŀǰû���������ӣ�</td>
          </tr>
        </table>
        <%end if%>
      </TD>
    </TR>
  </TBODY>
</TABLE>
</body>
</html>
