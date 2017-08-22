<!--#include file="adminconn.asp"-->
<!--#include file="Fckeditor/Fckeditor.Asp" -->
<%
Dim count,ranNum,iddata,Content
set rs=server.createobject("adodb.recordset")
sql = "select * from SmallClass order by SmallClassID asc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("SmallClassName"))%>","<%= trim(rs("BigClassName"))%>","<%= trim(rs("SmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.addNEWS.SmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.addNEWS.SmallClassName.options[document.addNEWS.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
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
    <td>　</td>
  </tr>
</table>
<table width="780" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <form name="addNEWS" method="post" action="Rencai_addok.asp">
    <tr> 
      <td height="25" colspan="2" align="center" background="images/b1.gif"><span class="style1"><font color="#FFFFFF">发布招聘信息</font></span> 
      </td>
    </tr>
    <tr> 
      <td width="84" height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>招聘职位：</td>
      <td width="679" bgcolor="#f9f9f9"> 
        <input name="zw" type="text" class="input" id="zw" size="30">
        如<font color="#FF0000">：义工经理</font></td>
    </tr>
    <tr> 
      <td height="-2" bgcolor="#f9f9f9"><font color="#FF0000">*</font>招聘人数：</td>
      <td bgcolor="#f9f9f9"><input name="zprs" type="text" class="input" id="zprs" size="30">
        如<font color="#FF0000">：5 人</font> </td>
    </tr>
    <tr> 
      <td height="-1" bgcolor="#f9f9f9"><font color="#FF0000">*</font>学历要求：</td>
      <td bgcolor="#f9f9f9"><input name="lxr" type="text" id="lxr" size="30">
        如<font color="#FF0000">：本科、大专等等</font> </td>
    </tr>
    <tr> 
      <td height="0" bgcolor="#f9f9f9"><font color="#FF0000">*</font>性别要求：</td>
      <td bgcolor="#f9f9f9"> 
        <select name="lxdh" id="lxdh">
          <option value="不限" selected>不限</option>
          <option value="男">男</option>
          <option value="女">女</option>
        </select>
        如<font color="#FF0000">：男、女或不限</font> </td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>年龄要求：</td>
      <td bgcolor="#f9f9f9"> <input name="lxdz" type="text" id="imgpath3" size="30">
        如<font color="#FF0000">：20-45岁等等</font></td>
    </tr>
    <tr> 
      <td height="25" valign="top" bgcolor="#f9f9f9"><font color="#FF0000">*</font>招聘说明：</td>
      <td valign="top" bgcolor="#f9f9f9">
        <%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "FCKeditor/"
oFCKeditor.ToolbarSet = "bijie520"
oFCKeditor.Width = "100%"
oFCKeditor.Height = "400"
oFCKeditor.Value = Content
oFCKeditor.Create "d_content"
%>
      </td>
    </tr>
    <tr> 
      <td height="30" colspan="2" align="center" bgcolor="#f9f9f9"> <input type="submit" name="Submit" value="提交" class="input">
        　 
        <input type="reset" name="Submit2" value="重置" class="input"> </td>
    </tr>
  </form>
</table>
</body>
</html>