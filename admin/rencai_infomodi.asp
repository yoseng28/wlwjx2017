<!--#include file="adminconn.asp"-->
<!--#include file="Fckeditor/Fckeditor.Asp" -->
<%
Dim zw,zprs,lxr,lxdz,lxdh,zwsm,ok,sContent,i,Count,id,content,ors
if request("modi")="yes" then
id=request("id")
zw=request("zw")
zprs=request("zprs")
lxr=request("lxr")
lxdz=request("lxdz")
lxdh=request("lxdh")

	zwsm = ""
	For i = 1 To Request.Form("d_content").Count
		zwsm = zwsm & Request.Form("d_content")(i)
	Next

'判断输入数据是否完整
if zw="" or zprs="" or zwsm="" then  
response.write"<div align=center><font color=red size=+3><br><br><strong>※出错了※</strong></font>"
response.write"<br><br><font color=green>你的输入有误,没有添加成功。</font>"
response.write"<br><br>请检查<font color=red>*</font>的地方是否全部填写。"
response.write"<br><br><a href=javascript:history.back()>返回继续填</a></div>"  
response.end 
end if

set rs=server.createobject("adodb.recordset")
sql="select * from rencai where id="&id
rs.open sql,conn,1,3
rs("zw")=zw
rs("zwsm")=zwsm
rs("zprs")=zprs
rs("lxr")=lxr
rs("lxdz")=lxdz
rs("lxdh")=lxdh
rs.update
rs.close
set rs=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('修改成功！');" & Chr(13)
		response.write "window.document.location.href='rencai_info.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
end if
%>
<%
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
<link href="images/style.css" rel="stylesheet" type="text/css">
<title>添加内容</title>
</head>

<body leftmargin="0" topmargin="0">
<% 
id=request("id")
Set ors=Server.CreateObject("ADODB.RecordSet") 
sql="select * from rencai where id="&id
ors.Open sql,conn,1,1
if ors.eof and ors.bof then
response.Write("没有记录")
else
content = ors("zwsm")
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="780" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <form name="addNEWS" method="post" action="rencai_infomodi.asp?modi=yes">
    <tr> 
      <td height="25" colspan="2" align="center" background="images/b1.gif"><span class="style1"><font color="#FFFFFF"><strong>修 
        改 产 品 内 容</strong></font></span> </td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>招聘职位：</td>
      <td width="681" bgcolor="#f9f9f9"> 
        <input name="zw" type="text" class="input" id="zw" value="<%=ors("zw")%>" size="30">
        如<font color="#FF0000">：产品经理</font></td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>招聘人数：</td>
      <td bgcolor="#f9f9f9"><input name="zprs" type="text" class="input" id="zprs" value="<%=ors("zprs")%>" size="30">
        如<font color="#FF0000">：5 人</font> </td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>学历要求：</td>
      <td bgcolor="#f9f9f9"><input name="lxr" type="text" id="lxr" value="<%=ors("lxr")%>" size="30">
        如<font color="#FF0000">：本科、大专等等</font> </td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>性别要求：</td>
      <td bgcolor="#f9f9f9"><input name="lxdh" type="text" id="lxdh" value="<%=ors("lxdh")%>" size="30">
        如<font color="#FF0000">：男、女或不限</font> </td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>年龄要求：</td>
      <td bgcolor="#f9f9f9"> <input name="lxdz" type="text" id="imgpath3" value="<%=ors("lxdz")%>" size="30">
        如<font color="#FF0000">：20-45岁等等</font></td>
    </tr>
    <tr> 
      <td width="82" height="25" valign="top" bgcolor="#f9f9f9"><font color="#FF0000">*</font>招聘说明：</td>
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
        <input type="hidden" name="id" value="<%=id%>">
        　 
        <input type="reset" name="Submit2" value="重置" class="input"> </td>
    </tr>
  </form>
</table>
<% End If
ors.close
set ors=nothing
 %>
</body>
</html>
