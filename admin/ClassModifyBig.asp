<!--#include file="adminconn.asp"-->
<%
dim BigClassID,Action,NewBigClassName,WriteErrMsg,OldBigClassName,FoundErr,ErrMsg,xianshi,xuehao,ClassType,aboutname,xsfs,linkfs,wblj
BigClassID=trim(Request("BigClassID"))
Action=trim(Request("Action"))
if Action="Modify" then
ClassType=trim(Request("ClassType"))
xianshi=trim(Request("xianshi"))
xuehao=trim(Request("xuehao"))
xsfs=trim(Request("xsfs"))
linkfs=trim(Request("linkfs"))
wblj=trim(Request("wblj"))
NewBigClassName=trim(Request("NewBigClassName"))
OldBigClassName=trim(Request("OldBigClassName"))
if BigClassID="" then
  response.Redirect("ClassManage.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from BigClass where BigClassID=" & CLng(BigClassID),conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>此产品大类不存在！</li>"
else
            rs("BigClassName")=NewBigClassName
			rs("ClassType")=ClassType
			rs("xianshi")=xianshi
			rs("xuehao")=xuehao
			rs("xsfs")=xsfs
			rs("linkfs")=linkfs
			if wblj="是" then rs("wblj")=true
    	 	rs.update
			rs.Close
	     	set rs=Nothing
			if NewBigClassName<>OldBigClassName then
				conn.execute "Update SmallClass set BigClassName='" & NewBigClassName & "' where BigClassName='" & OldBigClassName & "'"
				conn.execute "update content set BigClassName='" & NewBigClassName & "' where BigClassName='" & OldBigClassName & "'"
				conn.execute "update Product set ProductBigClass='" & NewBigClassName & "' where ProductBigClass='" & OldBigClassName & "'"
			end if
     		Response.Redirect "ClassManage.asp" 
end if
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from BigClass where BigClassID=" & CLng(BigClassID),conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>此产品大类不存在！</li>"
else
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改大类名称</title>
<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("大类名称不能为空！");
    document.form1.BigClassName.focus();
    return false;
  }
}
</script>
<LINK href="images/style.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>

<body>
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"><br>
      <a href="ClassManage.asp"><strong> 类 别 修 改</strong></a> <br> <br> <form name="form1" method="post" action="ClassModifyBig.asp">
        <table width="786" border="1" align="center" cellpadding="4" cellspacing="1" bordercolor="#376ed0" class="border">
          <tr bgcolor="#999999" class="title"> 
            <td height="25" colspan="3" align="center" background="images/b1.gif"><strong>添加大类</strong></td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td width="158" height="9"> <div align="right"><strong>大 类 名 称 ：</strong></div></td>
            <td width="150"> <input name="NewBigClassName" type="text" id="NewBigClassName" value="<%=rs("BigClassName")%>" size="20" maxlength="30"> <input name="BigClassID" type="hidden" id="BigClassID" value="<%=rs("BigClassID")%>"> 
              <input name="OldBigClassName" type="hidden" id="OldBigClassName" value="<%=rs("BigClassName")%>">
            </td>
            <td width="442">　</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td width="158" height="4"> <div align="right"><strong>栏 目 类 型 ：</strong></div></td>
            <td bgcolor="#f9f9f9"> <select name="ClassType" size="1" id="ClassType" onchange="dis(document.form1.ClassType.value)">
                <option  <%if rs("ClassType")="class.asp" then%> selected <%end if%> value="class.asp">普通页面</option>
                <option  <%if rs("ClassType")="about.asp" then%> selected <%end if%> value="about.asp">基本信息介绍</option>
                <option  <%if rs("ClassType")="rencai.asp" then%> selected <%end if%> value="rencai.asp">招聘页面</option>
                <option  <%if rs("ClassType")="product.asp" then%> selected <%end if%>  value="product.asp">图片介绍</option>
				<option  <%if rs("ClassType")="download.asp" then%> selected <%end if%>  value="download.asp">软件下载</option>
                <option  <%if rs("ClassType")="guestbook.asp" then%> selected <%end if%> value="guestbook.asp">意见反馈</option>
              </select> </td>
            <td bgcolor="#f9f9f9">　</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg" id="xsfs" style='display:none'> 
            <td width="158" height="1"> <div align="right"><strong>显 示 方 式 ：</strong></div></td>
            <td><select name="xsfs" size="1" id="select2">
                <option value="图片方式" <%if rs("xsfs")="图片方式" then%> selected <%end if%>>图片方式</option>
                <option value="图文列表" <%if rs("xsfs")="图文列表" then%> selected <%end if%>>图文列表</option>
              </select></td>
            <td><font color="#FF0000">只有图片才有效</font></td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td width="158" height="2"> <div align="right"><strong>首页栏目显示：</strong></div></td>
            <td> <select name="xianshi" size="1" id="select5">
                <option  <%if rs("xianshi")="ok" then%> selected <%end if%> value="ok">显示</option>
                <option  <%if rs("xianshi")="no" then%> selected <%end if%> value="no">不显示</option>
              </select></td>
            <td>首页导航是否显示</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td height="22" align="center"><div align="right"><strong>显 示 顺 序：</strong></div></td>
            <td height="22" align="center"> <div align="left"> 
                <input name="xuehao" type="text" id="xuehao2" value="<%=rs("xuehao")%>" size="10" maxlength="5">
                <span class="style1">只能是数字</span></div></td>
            <td bgcolor="#f9f9f9">数字越小越靠前显示</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td height="22" align="center"><div align="right"><strong>是否外部链接：</strong></div></td>
            <td height="22" bgcolor="#f9f9f9"><select name="wblj" size="1" id="wblj"  onchange="wdis(document.form1.wblj.value)">
                <option value="否" <%if rs("wblj")=false then%> selected <%end if%>>否</option>
                <option value="是" <%if rs("wblj")=true then%> selected <%end if%>>是</option>
              </select> </td>
            <td bgcolor="#f9f9f9">如果选择是，请务必输入外部链接地址。如外部链接地址未显示，</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"  id="wbljdz" style='display:none'> 
            <td height="22" align="right" bgcolor="#f9f9f9"><strong>外部链接地址：</strong></td>
            <td height="22" bgcolor="#f9f9f9"> <input name="linkfs" type="text" id="linkfs2" value="<%=rs("linkfs")%>" size="25" maxlength="100"> 
            </td>
            <td bgcolor="#f9f9f9">注意地址中必须含有 <font color="#FF0000">http://</font>,如：<font color="#FF0000">http://www.lzkeyuan.cn</font> 
              。</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td height="22" align="center">　</td>
            <td height="22" align="center"> <input name="Action" type="hidden" id="Action" value="Modify"> 
              <input name="Save" type="submit" id="Save" value=" 修 改 "> </td>
            <td align="center">　</td>
          </tr>
        </table>
        </form></td>
  </tr>
</table>
<script language="JavaScript">
<!--
function dis(vid)
{
   if(vid=="product.asp")
   {
   eval("xsfs.style.display=\"\";");
   form1.class_cell.value=1; 
   }
   else
   {
    eval("xsfs.style.display=\"none\";");
   form1.class_cell.value=0; 
   }
}
function wdis(vid)
{
   if(vid=="是")
   {
   eval("wbljdz.style.display=\"\";");
   form1.class_cell.value=1; 
   }
   else
   {
    eval("wbljdz.style.display=\"none\";");
   form1.class_cell.value=0; 
   }
}
//-->
</SCRIPT>
</body>
</html>
<%
end if
rs.close
set rs=nothing
%>