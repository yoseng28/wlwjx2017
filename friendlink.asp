<!--#include file="conn.asp"-->
<% 
bigclass=request("type")
smallclass=request("smallclass")
 %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="images/style.css" type=text/css rel=stylesheet>
<title><% =sitetitle %></title>
<script language="javascript">
<!--
function check(){
if(document.form1.name.value.length<1){
  	    alert("请留下网站名称！");
		  document.form1.name.focus();
		return false;
		}else if(document.form1.homepage.value.length<1){
		alert("请填写网站网址！？");
		  document.form1.homepage.focus();
		return false;
		} else if(document.form1.TxbContent.value.length<1){
		alert("请填写网站介绍！？");
		  document.form1.TxbContent.focus();
		return false;
	} else{
     return true;
   }
}
-->
</script>  
</head>
<BODY  topMargin=0 leftmargin="0" bgcolor="#FFFFFF">
<!--#include file="top.asp"-->
<TABLE cellSpacing=0 cellPadding=0 width=1100 align=center bgColor=#ffffff 
border=0>
  <TBODY>
    <TR> 
      <TD width="5" bgcolor="#FFFFFF"></TD>
      <TD width="1080">
		<TABLE width=1099 border=0 align=center cellPadding=0 cellSpacing=0>
          <TBODY>
            <TR> 
              <TD width="235" vAlign=top background="images/class_bg4.gif" style="background-repeat: repeat-x"
 bgcolor="#FFFFFF"> <TABLE style="MARGIN: 10px 0px 5px" cellSpacing=0 
                  cellPadding=0 width=100% border=0>
                  <TBODY>
                    <TR> 
                      <TD width="6%" height="25">　</TD>
                      <TD width="94%"><IMG height=16 
                        src="images/class_left0.gif" width=11 
                        align=absMiddle><STRONG><font color="FF9900">关于我们</font></STRONG> 
                        <STRONG>About Us</STRONG></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE width=90% border=0 align="center" cellPadding=5 cellSpacing=1 
                  bgColor=#dddddd>
                  <TBODY>
                    <TR> 
                      <TD width=186 bgColor=#ffffff> <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
                          <TBODY>
                            <TR> 
                              <TD height=8></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <%
sql2="select *  from about order by id asc "
Set oRS= Server.CreateObject("ADODB.recordset")
oRS.Open sql2,conn,1,3
 if NOT oRS.EOF then 
for i=1 to 10
  if NOT oRS.EOF then
   %>
                        <TABLE width=100% 
              border=0 align=center cellPadding=0 cellSpacing=0>
                          <TBODY>
                            <TR> 
                              <TD width="12%" height="25"> <div align="center"><font color="#0066CC">·</font></div></TD>
                              <TD width="88%"><a href="about.asp?type=<%=oRS("id")%>"><SPAN style="FONT-SIZE: 14px"><font color="#0066CC"><%=oRS("title")%></font></SPAN></a></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <%
oRS.MoveNEXT
end if
next
else Response.Write ("暂无内容.") end if
oRS.close 
Set oRS = Nothing

%>
                        <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
                          <TBODY>
                            <TR> 
                              <TD height=8></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE style="MARGIN: 10px 0px 5px" cellSpacing=0 
                  cellPadding=0 width=100% border=0>
                  <TBODY>
                    <TR> 
                      <TD width="12" height="25">　</TD>
                      <TD width="186"><IMG height=16 
                        src="images/class_left0.gif" width=11 
                        align=absMiddle><STRONG><font color="FF9900">站内搜索</font></STRONG> 
                        <STRONG>Search</STRONG></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE width=90% border=0 align="center" cellPadding=5 cellSpacing=1 
                  bgColor=#dddddd>
                  <TBODY>
                    <TR> 
                      <TD width=186 bgColor=#ffffff> <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
                          <TBODY>
                            <TR> 
                              <TD height=8></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE width=100% border=0 align="center" cellPadding=2 cellSpacing=0>
                          <SCRIPT language=JavaScript>
<!--
function postSearch(theform) {
	if(theform.key.value=="") {
		alert("请输入搜索关键字！");
		theform.key.focus();
		return false;
	}
	if(theform.leixing.value=="1") {
		alert("请选择搜索类型！");
		theform.leixing.focus();
		return false;
	}
}
-->
</SCRIPT>
                          <form name="form2" method="post" onsubmit="javascript:return postSearch(this)" action="search.asp">
                            <TBODY>
                              <TR> 
                                <TD width="60" rowspan="2" align=center><img src="images/class_search.gif" width="42" height="40"></TD>
                                <TD width="155" height="25" align=left><font color="#666666">&nbsp;请输入关键字：</font> 
                                </TD>
                              </TR>
                              <TR> 
                                <TD height="25" align=middle><INPUT class=box03 id=key size=15 
                        name=key> <input name="leixing" type="hidden" id="leixing" value="ProductName"></TD>
                              </TR>
                              <TR> 
                                <TD height="25" align=middle>　</TD>
                                <TD align=middle><INPUT id=ImageButton1 type=image alt="" 
                        src="images/go.gif" border=0 
                      name=ImageButton12></TD>
                              </TR>
                            </TBODY>
                          </form>
                        </TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
                          <TBODY>
                            <TR> 
                              <TD height=8></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
                  <TBODY>
                    <TR> 
                      <TD height=8></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
              <TD width="1">&nbsp;</TD>
              <TD width="830" vAlign=top style="BORDER-TOP: #C5CED7 1px solid;BORDER-LEFT: #C5CED7 1px solid;BORDER-right: #C5CED7 1px solid;BORDER-bottom: #C5CED7 1px solid;"> 
                <TABLE cellSpacing=0 cellPadding=0 width="724" 
border=0>
                  <TBODY>
                    <TR> 
                      <TD width="10" height="29" align="center">　</TD>
                      <TD width="296"> <font color="FF9900">申请友情链接</font> 
                      </TD>
                      <TD width="404" align="right"><font color="#FF0000">⊙ </font><a href="index.asp">首页</a> 
                        &gt;&gt; 申请友情链接 </TD>
                      <TD width="14">　</TD>
                    </TR>
                    <TR> 
                      <TD height="5" colspan="4" align="center" background="images/lanmufeng.gif"></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=100 align=center bgColor=#ffffff 
border=0>
                  <TBODY>
                    <TR> 
                      <TD height=8></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td><form action="friengadd.asp" method="POST" name=form1 target="_parent">
                        <table align=center border=0 cellspacing=0 
                          width="74%">
                          <tbody>
                            <tr> 
                              <td height="30"> <div align="right">　</div></td>
                            </tr>
                            <tr> 
                              <td height="25">网站名称: 
                                <input 
                              name=name class=bg id=name2 size=20 maxlength="108"> 
                                <font color="#FF0000">*</font> </td>
                            </tr>
                            <tr> 
                              <td height="25">网站网址: 
                                <input class=bg name="homepage" type="text" id="homepage2" size="20" maxlength="150"> 
                                <font color="#FF0000"> * <font color="#FFFFFF">注意：必须含有 
                                http://</font></font></td>
                            </tr>
                            <tr> 
                              <td height="25">QQ 号 码: 
                                <input 
                              name=oicq class=bg id=oicq size=20 maxlength="15"> 
                              </td>
                            </tr>
                            <tr> 
                              <td height="25"> E_ mail : 
                                <input class=bg name="email" type="text" id="email" size="20" maxlength="50"></td>
                            </tr>
                            <tr> 
                              <td height="25">网站介绍:</td>
                            </tr>
                            <tr> 
                              <td valign="top"> <table class=p11 cellspacing=0 cellpadding=1 
                        width="98%" border=0>
                                  <tbody>
                                    <tr> 
                                      <td valign=top><textarea id=textarea style="WIDTH: 350px; HEIGHT: 176px" name=TxbContent></textarea> 
                                        <font color="#FF0000">*</font> </td>
                                    </tr>
                                  </tbody>
                                </table></td>
                            </tr>
                            <tr> 
                              <td height="30"> <div align="left"> 
                                  <input type="submit" name="Submit" value="提 交 链 接 " onClick="return check();">
                                  　　 
                                  <input type="reset" name="Submit2" value="重 写">
                                </div></td>
                            </tr>
                          </tbody>
                        </table>
                      </FORM></td>
                  </tr>
                </table>
                <TABLE cellSpacing=0 cellPadding=0 width=100 align=center bgColor=#ffffff 
border=0>
                  <TBODY>
                    <TR> 
                      <TD height=8></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
                  <TBODY>
                    <TR> 
                      <TD height=8></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
      <TD width="5"></TD>
    </TR>
  </TBODY>
</TABLE>
<!--#include file="foot.asp"-->
</body>
</html>
