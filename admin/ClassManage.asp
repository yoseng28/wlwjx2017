<!--#include file="adminconn.asp"-->
<%
dim rsBigClass,rsSmallClass,ErrMsg
set rsBigClass=server.CreateObject("adodb.recordset")
rsBigClass.open "Select * From BigClass",conn,1,3
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
function checkSmall()
{
  if (document.form2.BigClassName.value=="")
  {
    alert("请先添加大类名称！");
	document.form1.BigClassName.focus();
	return false;
  }

  if (document.form2.SmallClassName.value=="")
  {
    alert("小类名称不能为空！");
	document.form2.SmallClassName.focus();
	return false;
  }
}
function ConfirmDelBig()
{
   if(confirm("确定要删除此大类吗？删除此大类同时将删除所包含的小类和该类下的所有新闻，并且不能恢复！"))
     return true;
   else
     return false;
	 
}

function ConfirmDelSmall()
{
   if(confirm("确定要删除此小类吗？删除此小类同时将删除该类下的所有新闻，并且不能恢复！"))
     return true;
   else
     return false;
	 
}
</script>
<LINK href="images/style.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {
	color: #FF0000;
	font-weight: bold;
}
-->
</style>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top">
      <a href="ClassAddBig.asp"><strong><font color="#FF0000"><u>添加一级分类</u></font></strong></a><br>
      <br> 
      <table width="747" border="1" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
        <tr bgcolor="#999999"> 
          <td width="271" height="25" align="center" background="images/b1.gif" bgcolor="#999999"><strong>栏目名称</strong></td>
          <td width="128" height="25" align="center" background="images/b1.gif"><strong>首页栏目显示</strong></td>
          <td width="128" height="25" align="center" background="images/b1.gif"><strong>显示顺序</strong></td>
          <td width="215" height="25" align="center" background="images/b1.gif"><strong>操作选项</strong></td>
        </tr>
        <%
	do while not rsBigClass.eof
%>
        <tr bgcolor="#EAEAEA" class="tdbg"> 
          <td width="271" height="22" bgcolor="#EAEAEA"><img src="images/%2B.gif" width="15" height="15"><%=rsBigClass("BigClassName")%></td>
          <td align="right" bgcolor="#EAEAEA" style="padding-right:10"> 
            <div align="center">
              <%if rsBigClass("xianshi")="ok" then%>
              <strong>显　示</strong> 
              <%else%>
              <span class="style1">不显示</span> 
              <%end if%>
            </div></td>
          <td align="right" style="padding-right:10"> 
            <div align="center"><%=rsBigClass("xuehao")%></div></td>
          <td align="right" style="padding-right:10"><a href="ClassAddSmall.asp?BigClassName=<%=rsBigClass("BigClassName")%>"><font color="#FF0000">添加二级分类</font></a> 
            | <a href="ClassModifyBig.asp?BigClassID=<%=rsBigClass("BigClassID")%>">修改</a> 
            | <a href="ClassDelBig.asp?BigClassName=<%=rsBigClass("BigClassName")%>" onClick="return ConfirmDelBig();">删除</a></td>
        </tr>
        <%
	  set rsSmallClass=server.CreateObject("adodb.recordset")
	  rsSmallClass.open "Select * From SmallClass Where BigClassName='" & rsBigClass("BigClassName") & "'",conn,1,1
	  if not(rsSmallClass.bof and rsSmallClass.eof) then
		do while not rsSmallClass.eof
	%>
        <tr bgcolor="#f9f9f9" class="tdbg"> 
          <td width="271" height="22" bgcolor="#f9f9f9">&nbsp;&nbsp;<img src="images/-.gif" width="15" height="15"><%=rsSmallClass("SmallClassName")%>&nbsp;</td>
          <td colspan="2" align="right" bgcolor="#f9f9f9" style="padding-right:10"> 
            <div align="center">&nbsp;</div></td>
          <td align="right" style="padding-right:10"><a href="ClassModifySmall.asp?SmallClassID=<%=rsSmallClass("SmallClassID")%>">修改</a> 
            | <a href="ClassDelSmall.asp?SmallClassID=<%=rsSmallClass("SmallClassID")%>&SmallClassName=<%=rsSmallClass("SmallClassName")%>" onClick="return ConfirmDelSmall();">删除</a></td>
        </tr>
        <%
			rsSmallClass.movenext
		loop
	  end if
	  rsSmallClass.close
	  set rsSmallClass=nothing	
	  rsBigClass.movenext
	loop
%>
      </table>
      
    </td>
  </tr>
</table>
<%
rsBigClass.close
set rsBigClass=nothing
%>
