<!--#include file="adminconn.asp"-->
<%
  Dim bigclass,page,pages,j,keyword
%>
<% 
keyword=request("keyword")
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
    <td height=30><div align="center"><strong><font color="#FF0000">ͼƬ����
</font></strong></div></td>
  </tr>
</table>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height=5><table width="602" border="1" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
        <form name="form2" method="post" action="Product_info.asp">
          <tr> 
            <td width="90" height=30 align="center"><strong>ͼƬ
����</strong></td>
            <td width="217"> ���⣺ 
              <input name="keyword" type="text" id="keyword2"></td>
            <td width="119" align="center"> <input type="submit" name="Submit" value=" �� �� "> 
            </td>
            <td width="145" align="center"><a href="Product_info.asp"><font color="#FF0000">�鿴��������</font></a></td>
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
 function check1()//���ж�
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
    alert("���ٱ���ѡ��һ���಻�ܳ���100��!\r\t������ѡ����"+j+"��");
    return false;
   }
   
     if(!confirm("��ȷ��Ҫɾ����ѡ�����𣬲��Ҳ��ָܻ���"))
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
 <form name="form1" method="post" action="Product_infodel.asp"> <tr> 
    <td width="5%" height="25" align="center" bgcolor="#CCCCCC" background="images/b1.gif">ѡ��</td>
    <td width="25%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">����</td>
    <td width="10%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">������</td>
    <td width="15%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">����һ������</td>
    <td width="15%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">������������</td>
    <td width="20%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">��������</td>
    <td width="10%" align="center" bgcolor="#CCCCCC" background="images/b1.gif">����</td>
  </tr>
  <%
page=clng(request("page"))
Set rs=Server.CreateObject("ADODB.RecordSet") 

if keyword<>"" then
sql="select * from Product Where ProductName Like '%"& keyword &"%' order by id desc"
else
sql="select * from Product order by id desc"
end if


rs.Open sql,conn,1,1
if rs.eof and rs.bof then
response.Write("û�м�¼")
else
rs.PageSize=30
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.AbsolutePage=page  
    
for j=1 to rs.PageSize 
%>
  <tr bgcolor="#f9f9f9"> 
    <td height="22" align="center"><input name="id" type="checkbox" id="id" value="<%=rs("id")%>"></td>
    <td bgcolor="#f9f9f9"> &nbsp; 
      <% if rs("ProductImages")<>"" then response.write "<img src='images/news.gif' border=0 alt='ͼƬ����'>" end if %>
      <a href="Product_infomodi.asp?id=<%=rs("id")%>" title="<%=rs("ProductName")%>"> 
      <% =rs("ProductName") %>
      </a></td>
    <td align="center">&nbsp;<%=rs("ProductUser")%></td>
    <td align="center">&nbsp;<%=rs("ProductBigClass")%> </td>
    <td align="center">&nbsp;<%=rs("ProductSmallClass")%></td>
    <td align="center">&nbsp;<%=rs("ProductTime")%></td>
    <td align="center"><a href="Product_infomodi.asp?id=<%=rs("id")%>">�޸�</a> 
      <a href="Product_infodel.asp?id=<%=rs("id")%>" onClick="return ConfirmDelnr();">ɾ��</a></td>
  </tr>
  <%
rs.movenext
if rs.eof then exit for
next                                                       
%>
  <tr bgcolor="#f9f9f9"> 
    <td height="22" colspan="7" align="center"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="1%" height="30" align="center">��</td>
          <td width="40%"> <input name="chkAll" type="checkbox" id="chkAll5" onclick=CheckAll(this.form) value="checkbox">
            ѡ�б�ҳ��ʾ��������Ϣ</td>
          <td width="59%"> <input name="submit" type='submit' value='ɾ��ѡ������Ϣ'> 
            <input name="action" type="hidden" id="action3" value="del"> </td>
        </tr>
      </table></td>
  </tr>
  </form>
</table>
  
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF"> 
      <td height="30" align="center"> 
<%
if keyword="" then
%>
<%if Page<2 then      
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=Product_info.asp?page=1>��ҳ</a>&nbsp;"
    response.write "<a href=Product_info.asp?page=" & Page-1 & ">��һҳ</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=Product_info.asp?page=" & (page+1) & ">"
    response.write "��һҳ</a> <a href=Product_info.asp?page="&rs.pagecount&">βҳ</a>"
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
    response.write "<a href=Product_info.asp?keyword="&keyword&"&page=1>��ҳ</a>&nbsp;"
    response.write "<a href=Product_info.asp?keyword="&keyword&"&page=" & Page-1 & ">��һҳ</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=Product_info.asp?keyword="&keyword&"&page=" & (page+1) & ">"
    response.write "��һҳ</a> <a href=Product_info.asp?keyword="&keyword&"&page="&rs.pagecount&">βҳ</a>"
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
