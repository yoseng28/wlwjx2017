<!--#include file="adminconn.asp"-->
<% 
Dim bigclass,page,pages,j,keyword
bigclass=request("type")
keyword=request("keyword")
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
<LINK href="../../../admin/images/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>

<body topmargin="0">
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height=5></td>
  </tr>
</table>
<table width="95%" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <tr bgcolor="#f9f9f9"> 
    <td width="14%" height="50"> 
      <div align="right"><font color="#FF0000"><strong>栏目内容管理：</strong></font></div></td>
    <td width="86%" height="40"> 
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td> 
            <% sql="select * from bigclass where ClassType='class.asp' order by xuehao asc"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
do while not rs.eof
%>
            &nbsp;[<a href="admin_info.asp?type=<%=rs("BigClassName")%>" class=nav14><font color="#0000FF"><%=rs("BigClassName")%></font></a>] 
            <%
rs.movenext
loop
rs.Close
set rs=nothing
%>
          </td>
        </tr>
      </table></td>
  </tr>
</table>

  
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height=5></td>
  </tr>
</table>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height=5><table width="602" border="1" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
        <form name="form2" method="post" action="admin_info.asp">
          <tr> 
            <td width="90" height=30 align="center"><strong>标题搜索</strong></td>
            <td width="213"> 标题：
<input name="keyword" type="text" id="keyword"></td>
            <td width="123" align="center"> <input type="submit" name="Submit" value=" 搜 索 "> 
            </td>
            <td width="145" align="center"><a href="admin_info.asp"><font color="#FF0000">查看所有内容</font></a></td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height=5></td>
  </tr>
</table>
<SCRIPT language=javascript>
<!--
 function check1()//表单判断
  {
   var str="";
   var j=0;
   var el=document.form1.getElementsByTagName("input")
   for(i=0;i<el.length;i++){
    if(el[i].type=="checkbox"){
     if((el[i].checked)&&(el[i].name.indexOf("id")!='-1')){
      j++;
     }
    }
   }
   if(j > 100 || j<1)
   {
    alert("最少必须选择一项，最多不能超过100项!\r\t你现在选择了"+j+"项");
    return false;
   }
   
     if(!confirm("您确定要删除所选文章吗，并且不能恢复？"))
   {
    return false
   } 
   return true;
}

function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll"&&e.disabled==false)
       e.checked = form.chkAll.checked;
    }
  }
//-->
</SCRIPT>
  <table width="95%" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <form name="form1" method="post" action="admin_infodel.asp">
    <tr> 
      <td width="5%" height="25" align="center" bgcolor="#CCCCCC" background="../../../bijie520_qyjz_blue/admin/images/b1.gif">选择</td>
      <td width="25%" align="center" bgcolor="#CCCCCC" background="../../../bijie520_qyjz_blue/admin/images/b1.gif">标题</td>
      <td width="10%" align="center" bgcolor="#CCCCCC" background="../../../bijie520_qyjz_blue/admin/images/b1.gif">发布者</td>
      <td width="15%" align="center" bgcolor="#CCCCCC" background="../../../bijie520_qyjz_blue/admin/images/b1.gif">所属一级分类</td>
      <td width="15%" align="center" bgcolor="#CCCCCC" background="../../../bijie520_qyjz_blue/admin/images/b1.gif">所属二级分类</td>
      <td width="20%" align="center" bgcolor="#CCCCCC" background="../../../bijie520_qyjz_blue/admin/images/b1.gif">发布日期</td>
      <td width="10%" align="center" bgcolor="#CCCCCC" background="../../../bijie520_qyjz_blue/admin/images/b1.gif">操作</td>
    </tr>
    <%
page=clng(request("page"))
Set rs=Server.CreateObject("ADODB.RecordSet") 
if bigclass<>"" then
sql="select * from content Where BigClassName='" & bigclass & "' order by id desc"
elseif keyword<>"" then
sql="select * from content Where title Like '%"& keyword &"%' order by id desc"
else
sql="select * from content order by id desc"
end if
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
response.Write("没有记录")
else
rs.PageSize=30
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.AbsolutePage=page  
    
for j=1 to rs.PageSize 
%>
    <tr bgcolor="#f9f9f9"> 
      <td height="22" align="center"><input name="id" type="checkbox" id="id" value="<%=rs("id")%>"> 
      </td>
      <td bgcolor="#f9f9f9">　 
        <% if rs("firstImageName")<>"" then response.write "<img src='images/news.gif' border=0 alt='图片内容'>" end if %>
        <a href="admin_infomodi.asp?id=<%=rs("id")%>" title="<%=rs("title")%>"><%=left(rs("title"),12)%></a></td>
      <td align="center">&nbsp;<%=left(rs("user"),5)%></td>
      <td align="center">&nbsp;<%=rs("BigClassName")%> </td>
      <td align="center" bgcolor="#f9f9f9">&nbsp;<%=rs("SmallClassName")%></td>
      <td align="center">&nbsp;<%=rs("infotime")%></td>
      <td align="center"><a href="admin_infomodi.asp?id=<%=rs("id")%>">修改</a> 
        <a href="admin_infodel.asp?id=<%=rs("id")%>" onClick="return ConfirmDelnr();">删除</a></td>
    </tr>
    <%
rs.movenext
if rs.eof then exit for
next                                                       
%>
    <tr bgcolor="#f9f9f9"> 
      <td height="35" colspan="7" align="center" bgcolor="#f9f9f9"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="1%" height="30" align="center">　</td>
            <td width="40%">
<input name="chkAll" type="checkbox" id="chkAll5" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的所有信息</td>
            <td width="59%"> 
              <input name="submit" type='submit' value='删除选定的信息'> 
              <input name="action" type="hidden" id="action3" value="del"> </td>
          </tr>
        </table>
      </td>
    </tr>
  </form>
  </table>
  
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF"> 
      <td height="30" align="center"> 
<%
if bigclass="" and keyword="" then
%>
<%if Page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=admin_info.asp?page=1>首页</a>&nbsp;"
    response.write "<a href=admin_info.asp?page=" & Page-1 & ">上一页</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=admin_info.asp?page=" & (page+1) & ">"
    response.write "下一页</a> <a href=admin_info.asp?page="&rs.pagecount&">尾页</a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#FF0000'>"&rs.recordcount&"</font></b>条记录 <b>"&rs.pagesize&"</b>条记录/页"
%>
<%
elseif keyword="" then
%>
<%if Page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=admin_info.asp?type="&bigclass&"&page=1>首页</a>&nbsp;"
    response.write "<a href=admin_info.asp?type="&bigclass&"&page=" & Page-1 & ">上一页</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=admin_info.asp?type="&bigclass&"&page=" & (page+1) & ">"
    response.write "下一页</a> <a href=admin_info.asp?type="&bigclass&"&page="&rs.pagecount&">尾页</a>"
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
    response.write "<a href=admin_info.asp?keyword="&keyword&"&page=1>首页</a>&nbsp;"
    response.write "<a href=admin_info.asp?keyword="&keyword&"&page=" & Page-1 & ">上一页</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=admin_info.asp?keyword="&keyword&"&page=" & (page+1) & ">"
    response.write "下一页</a> <a href=admin_info.asp?keyword="&keyword&"&page="&rs.pagecount&">尾页</a>"
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
