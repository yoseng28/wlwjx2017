<!--#include file="adminconn.asp"-->

<%
dim Action,BigClassName,FoundErr,ErrMsg,ClassType,xianshi,xuehao,rsabout,xsfs,linkfs,wblj
Action=trim(Request("Action"))
BigClassName=trim(request("BigClassName"))
ClassType=trim(request("ClassType"))
xianshi=trim(request("xianshi"))
xuehao=request("xuehao")
xsfs=request("xsfs")
linkfs=trim(request("linkfs"))
wblj=trim(request("wblj"))
if wblj="是" and linkfs="" then
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('您选择的是外部链接，但未输入外部链接地址，请重新输入外部地址！');" & Chr(13)
		response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
end if
if Action="Add" then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From BigClass Where BigClassName='" & BigClassName & "'",conn,1,3
		if not (rs.bof and rs.EOF) then
		%>
<script language="JavaScript" type="text/JavaScript">

    alert("大类名称已经存在，请重新输入！");

</script>
<%
			
		else
    	 	rs.addnew
     		rs("BigClassName")=BigClassName
			rs("ClassType")=ClassType	
			rs("xianshi")=xianshi		
			rs("xuehao")=xuehao
			rs("xsfs")=xsfs
			rs("linkfs")=linkfs
			if wblj="是" then rs("wblj")=true
    	 	rs.update
     		rs.Close
	     	set rs=Nothing
			
		if ClassType="about.asp" then
		Set rsabout=Server.CreateObject("Adodb.RecordSet")
		rsabout.open "Select * From about Where title='" & BigClassName & "'",conn,1,3
		if not (rsabout.bof and rsabout.EOF) then
		%>
<script language="JavaScript" type="text/JavaScript">

    alert("大类名称已经存在，请重新输入！");

</script>
<%
			
		else
    	 	rsabout.addnew
     		rsabout("title")=BigClassName
			rsabout("content")=BigClassName
			rsabout("abouttype")=1
    	 	rsabout.update
     		rsabout.Close
	     	set rsabout=Nothing
    	  
	end if
	end if
	call CloseConn()
	Response.Redirect "ClassManage.asp" 					
	end if
	end if
%>

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
<table width="794" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="794" align="center" valign="top"> <br>
      <a href="ClassManage.asp"><strong>类 别 设 置</strong></a> <br>
      <form name="form1" method="post" action="ClassAddBig.asp" onsubmit="return checkBig()">
        <table width="786" border="1" align="center" cellpadding="4" cellspacing="1" bordercolor="#376ed0" class="border">
          <tr bgcolor="#999999" class="title"> 
            <td height="25" colspan="3" align="center" background="images/b1.gif"><strong>添加大类</strong></td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td width="158" height="9"> 
              <div align="right"><strong>大 类 名 称 ：</strong></div></td>
            <td width="150"> 
              <input name="BigClassName" type="text" size="20" maxlength="30"> 
            </td>
            <td width="442">　</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td width="158" height="4"> 
              <div align="right"><strong>栏 目 类 型 ：</strong></div></td>
            <td bgcolor="#f9f9f9"> <select name="ClassType" id="ClassType" onchange="dis(document.form1.ClassType.value)">
                <option value="class.asp" selected>普通页面</option>
                <option value="about.asp">基本信息介绍</option>
                <option value="product.asp">图片介绍</option>
				<option value="download.asp">软件下载</option>
                <option value="guestbook.asp">意见反馈</option>
                <option value="rencai.asp">招聘页面</option>
              </select> </td>
            <td bgcolor="#f9f9f9">　</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg" id="xsfs" style='display:none'> 
            <td width="158" height="1"> 
              <div align="right"><strong>显 示 方式 ：</strong></div></td>
            <td><select name="xsfs" size="1" id="select2">
                <option value="图片方式" selected>图片方式</option>
                <option value="图文列表">图文列表</option>
              </select></td>
            <td><font color="#FF0000">只有图片模板 才有效</font></td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td width="158" height="2"> 
              <div align="right"><strong>首页栏目显示：</strong></div></td>
            <td> <select name="xianshi" size="1" id="select">
                <option selected value="ok">显示</option>
                <option value="no">不显示</option>
              </select></td>
            <td>首页导航是否显示</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td height="22" align="center"><div align="right"><strong>显 示 顺 序：</strong></div></td>
            <td height="22" align="center"> <div align="left"> 
                <input name="xuehao" type="text" id="xuehao3" size="10" maxlength="5">
                <span class="style1">只能是数字</span></div></td>
            <td bgcolor="#f9f9f9">数字越小越靠前显示</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td height="22" align="center"><div align="right"><strong>是否外部链接：</strong></div></td>
            <td height="22" bgcolor="#f9f9f9"><select name="wblj" size="1" id="wblj"  onchange="wdis(document.form1.wblj.value)">
                <option value="否" selected>否</option>
                <option value="是">是</option>
              </select> </td>
            <td bgcolor="#f9f9f9">如果选择是，请务必输入外部链接地址。</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"  id="wbljdz" style='display:none'> 
            <td height="22" align="right" bgcolor="#f9f9f9"><strong>外部链接地址：</strong></td>
            <td height="22" bgcolor="#f9f9f9"> 
              <input name="linkfs" type="text" id="linkfs2" size="25" maxlength="100"> 
            </td>
            <td bgcolor="#f9f9f9">注意地址中必须含有 <font color="#FF0000">http://</font>,如：<font color="#FF0000">http://www.lzkeyuan.cn</font> 
              。</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg">
            <td height="22" align="center">　</td>
            <td height="22" align="center"> <input name="Action" type="hidden" id="Action2" value="Add"> 
              <input name="Add" type="submit" value=" 添 加 "></td>
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
