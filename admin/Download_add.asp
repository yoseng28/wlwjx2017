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
<table width="799" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <form name="addNEWS" method="post" action="Download_addok.asp">
    <tr> 
      <td height="25" colspan="3" align="center" background="images/b1.gif"><strong><span class="style1"><font color="#FFFFFF">添 
        加 下 载 内 容</font></span> </strong></td>
    </tr>
    <tr> 
      <td width="88" height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>文件名称：</td>
      <td colspan="2" bgcolor="#f9f9f9"> <input name="DownName" type="text" class="input" id="DownName" size="30"></td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>文件分类：</td>
      <td colspan="2" bgcolor="#f9f9f9"> 
        <%
        sql = "select * from BigClass where ClassType='download.asp' "
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
		else
		%>
        <select name="BigClassName" onChange="changelocation(document.addNEWS.BigClassName.options[document.addNEWS.BigClassName.selectedIndex].value)" size="1">
          <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
          <%
			dim selclass
		    selclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
          <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
          <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
        </select> <select name="SmallClassName">
          <option value="" selected>选择小类</option>
          <%
			sql="select * from SmallClass where BigClassName='" & selclass & "'"
			rs.open sql,conn,1,1
			if not(rs.eof and rs.bof) then
			%>
          <option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
          <% rs.movenext
				do while not rs.eof%>
          <option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
          <%
			    	rs.movenext
				loop
			end if
	        rs.close
			%>
          <%
			ranNum=int(9*rnd)+10
			iddata=month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
			%>
        </select> </td>
    </tr>
    <tr>
      <td height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>文件大小：</td>
      <td colspan="2" bgcolor="#f9f9f9"> <input name="Downdaxiao" type="text" class="input" id="Downdaxiao" size="30">
        <font color="#FF0000"> 如：1M，250M，.....</font></td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>下载地址：</td>
      <td width="188" bgcolor="#f9f9f9"> <input name="firstImageName" type="text" id="imgpath3" size="30"> 
      </td>
      <td width="499" bgcolor="#f9f9f9"> <iframe style="top:2px" id="UploadFiles" src="shangchuan_index.asp?picID=1" frameborder=0 scrolling=no width="320" height="25"></iframe>
		<font color="#CC0000"><b>在此上传文件（小于30M)</b></font></td>
    </tr>
    <tr> 
      <td height="20" valign="top" bgcolor="#f9f9f9"><font color="#FF0000">*</font>使用说明：</td>
      <td colspan="2" valign="top" bgcolor="#f9f9f9"> 
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
      <td height="20" bgcolor="#f9f9f9">发布人：</td>
      <td colspan="2" bgcolor="#f9f9f9"> <input name="user" type="text" class="input" value="管理员" size="30"></td>
    </tr>
    <tr> 
      <td height="30" colspan="3" align="center" bgcolor="#f9f9f9"> <input type="submit" name="Submit" value="提交" class="input">
        　 
        <input type="reset" name="Submit2" value="重置" class="input"> </td>
    </tr>
  </form>
</table>
</body>
</html>