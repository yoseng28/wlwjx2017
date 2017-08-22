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
Response.Write"alert(""恭喜.留言审核成功"");"
Response.Write"window.location='rencai_guest_admin.asp'"
Response.Write"</script>"
end if
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="images/style.css" type="text/css">
<title>添加内容</title>
</head>

<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="40">
<div align="center"><strong><font color="#0000FF" size="+2">应聘留言管理</font></strong></div></td>
  </tr>
</table>
<TABLE width=95% border=0 align="center" cellPadding=0 
            cellSpacing=0>
  <TBODY>
    <TR> 
      <TD>&nbsp;</TD>
    </TR>
    <TR> 
      <TD height=496 align="center" vAlign=top> 
        <% 
	page = CLng(request("page"))							'利用CLng函数把page值转换为Long型
	judge=request("judge")
	judge2=request("judge2")
	judge3=0
sql="select *  from guestbook where classname='rencai' order by id desc "
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
        <table width="100%" border="1" cellpadding="2" cellspacing="1" bordercolor="#376ed0" class="itable">
          <tr> 
            <td width="46%" height="12" background="images/b1.gif" class="distance"> 
              <font color="#000000"><strong><font color="#FFFFFF">应聘职位：</font></strong></font> 
              <strong></span><a href="../rencaiview.asp?id=<%=rs("biaotiid")%>" target="_blank"><font color="#FFFFFF"><%=rs("name") %></font></a></strong></td>
            <td width="24%" background="images/b1.gif" class="distance"><strong><font color="#FFFFFF">留言时间：</font></strong><font color="#FFFFFF"><span class="style14"><%=year(rs("time01"))%> 
              - <%=month(rs("time01"))%> - <%=day(rs("time01"))%></span></font></td>
            <td width="30%" background="images/b1.gif" class="distance"><span class="style14"> 
              <% if rs("reply")="0" then%>
              [ <a href="rencai_guest_admin.asp?id=<%=rs("id")%>&sh=yes"><font color="#ffffff"><strong>未处理</strong></font></a> 
              ] 
              <% end if %>
              [ <a href="guest_delete.asp?id=<%=rs("id")%>" class="blue"  onClick="return ConfirmDelnr();"><font color="#FFFFFF"><strong>删除该条留言</strong></font></a> 
              ]</span></td>
          </tr>
          <tr> 
            <td height="12" bgcolor="#EAEAEA" class="distance"><strong><font color="#000000">联 
              系 人：</font><font color="#FF00FF"><%=rs("homepage")%></font></strong></td>
            <td height="12" bgcolor="#EAEAEA" class="distance"><strong><font color="#000000">联系电话：</font><font color="#FF00FF"><%=rs("oicq")%></font></strong></td>
            <td height="12" bgcolor="#EAEAEA" class="distance"><strong><font color="#000000">联系人E_mail：</font><font color="#FF00FF"><%=rs("email")%></font></strong></td>
          </tr>
          <tr> 
            <td height="25" colspan="3" bgcolor="#EAEAEA" class="distance"><strong><font color="#000000">联系地址：</font></strong><span class="style14"><%=rs("from01")%></span>　</td>
          </tr>
          <tr> 
            <td colspan="3" bgcolor="#f9f9f9" class="distance"><span class="style14"><font color="#000000"><strong>留言内容：</strong></font><%=rs("guestcontent")%></span></td>
          </tr>
        </table>
        <div align="center">
          <table width="98%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td>&nbsp;</td>
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
            <td>&nbsp;&nbsp;&nbsp; <div align="center">共有<font 
                     color=red><%=rs.recordCount%> </font>个留言 页数：<font 
                        color=red><%=page%> </font>/&nbsp; <%=rs.PageCount%>&nbsp;&nbsp; 
                <%
				If page = 1 Then 
				   Response.Write "<font class='white'>第一页　上一页</font>" 
				End If
				If page <> 1 Then 
				   Response.Write "<a href=rencai_guest_admin.asp?page=1 class='yellow'>第一页</a>" 
				   Response.Write "　<a href=rencai_guest_admin.asp?page=" & (page - 1) & " class='yellow'>上一页</a>" 
				End If
				If page <> RS.PageCount Then 
				   Response.Write "　<a href=rencai_guest_admin.asp?page=" & (page + 1) & " class='yellow'>下一页</a>" 
				   Response.Write "　<a href=rencai_guest_admin.asp?page=" & RS.PageCount & " class='yellow'>最后一页</a>" 
				End If 
				If page = RS.PageCount Then 
				   Response.Write "　<font class='white'>下一页　最后一页</font>" 
				End If
			  %>
                &nbsp; </div></td>
          </tr>
        </table>
        <%else%>
        <table width="75%" border="0" align="center" cellspacing="0">
          <tr> 
            <td>目前没有用户留言！</td>
          </tr>
        </table>
        <%end if%>
      </TD>
    </TR>
    <TR> 
      <TD>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
</body>
</html>
