<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">

<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"--> 
<%
id=trim(request("id"))
id=replace(id,"'","")
id=left(id,8)
call view_content(title,infotime,hits,content,user)
%>
<SCRIPT language=JavaScript>
var currentpos,timer;

function initialize()
{
timer=setInterval("scrollwindow()",50);
}
function sc(){
clearInterval(timer);
}
function scrollwindow()
{
currentpos=document.body.scrollTop;
window.scroll(0,++currentpos);
if (currentpos != document.body.scrollTop)
sc();
}
document.onmousedown=sc
document.ondblclick=initialize
</SCRIPT>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%= title %>--<% =sitetitle %></title>
<meta name="keywords" content="<%= title %>,<% =webKeywords %>">
<meta name="description" content="<%= title %>,<% =webDescription %>">
<meta name="robots" content="index, follow">
<meta name="googlebot" content="index, follow">
<style fprolloverstyle>A:hover {color: #205081; font-weight: bold}
</style>
<style>
<!--
.xuesheng    { font-size: 12pt; font-weight: bold }
.title        { font-size: 14pt }
.shuye        { font-size: 11pt }

.as          { font-size: 12pt; font-family: 微软雅黑; font-weight: bold }
.anli        { font-family: 微软雅黑; font-size: 14pt; font-weight: bold }

td{font-family:arial}td{font-size:10pt;line-height:22px}.text        { font-family: 微软雅黑; font-size: 12pt }
-->
</style>
</head>
<BODY topMargin=0 leftmargin="0" link="#666666" vlink="#666666" alink="#666666">
<!--#include file="top.asp"-->
<div align="center">
<TABLE cellSpacing=0 cellPadding=0 width=1102 bgColor=#ffffff  vAlign=top style="BORDER-TOP: #808080 1px solid;BORDER-LEFT: #808080 1px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 1px solid;"
border=0>
  <TBODY>
    <TR> 
      <TD width="1100"><TABLE cellSpacing=0 cellPadding=0 width=1100 align=center border=0>
          <TBODY>
            <TR> 
              <TD width="1100" vAlign=top> 
				<TABLE cellSpacing=0 cellPadding=0 width="1100" 
border=0 height="38">
                  <TBODY>
                    <TR> 
                      <TD width="1100" height="56" align="center" style="BORDER-BOTTOM: #CFCFCF 1px dashed" background="fl/cp1.gif">
						<p class="title" align="left"><b><font color="#205081">
						<img border="0" src="fl/jiantou.jpg" width="17" height="14"></font></b><font face="微软雅黑"><a href="index.asp" style="text-decoration: none"><SPAN style="FONT-WEIGHT: bold;"><font color="#004993">首页</font></SPAN></a><b><font color="#004993"> 
                        &gt;&gt;<%=bigclass%> 
 						</font></b></font>
						<span style="letter-spacing: normal; background-color: #FFFFFF">
						<b><font face="微软雅黑" color="#004993">图</font></b></span><b style="color: rgb(0, 0, 0); font-family: Arial, 微软雅黑; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: 37.3333px; orphans: auto; text-align: -webkit-left; text-indent: 0px; text-transform: none; white-space: normal; widows: 1; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255);"><font color="#004993" face="微软雅黑">文动态&gt;&gt;</font></b><font face="微软雅黑"><b><font color="#004993">正文内容 </font></b></font></TD>
                    </TR>
					</TBODY>
                </TABLE>
                <TABLE width=1100 
border=0 align="center" cellPadding=0 cellSpacing=0 id=table4>
                  <TBODY>
                    <TR> 
                      <TD height=45 colspan="4"> <div align="center"><B><SPAN style="FONT-SIZE: 12px; LINE-HEIGHT: 35px; LETTER-SPACING: 2px">
						<FONT size=3 color="#5E5E5E" face="微软雅黑"> 
                          <%=title %> </FONT></SPAN></B></div></TD>
                    </TR>
                    <TR> 
                      <TD height=9 colspan="4"> <HR width="100%" color=#c0c0c0 SIZE=1></TD>
                    </TR>
                    <TR> 
                      <TD height=30 colspan="4"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td> <div align="right">
								<font color="#333333" face="微软雅黑">发布时间：<%=infotime%>&nbsp;&nbsp;　　浏览：<%=hits%></font></div></td>
                          </tr>
                        </table></TD>
                    </TR>
                    <TR> 
                      <TD height=15 colspan="4"> <HR width="100%" color=#c0c0c0 SIZE=1></TD>
                    </TR>
                    <TR> 
                      <TD colspan="4" valign="top"> <DIV align=center> 
                          <table width="98%" height="60" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td align="left"><SPAN style="FONT-SIZE: 14px; LINE-HEIGHT: 25px;">
								<font color="#333333" face="微软雅黑"><%=content %> 
                                </font> 
                                </SPAN></td>
                            </tr>
                          </table>
                        </DIV></TD>
                    </TR>
                    <TR valign="middle"> 
                      <TD height="45" align="center" width="535">　</TD>
                      <TD align="right" valign="bottom" width="81"><A href="javascript:window.print();">
						<font color="#666666">
						<IMG 
            height=23 src="images/r_print.gif" width=71 
            border=0></font></A></TD>
                      <TD align="center" width="34">　</TD>
                      <TD valign="bottom" width="193"><A href="javascript:window.top.close();">
						<font color="#666666">
						<IMG 
            height=23 src="images/r_closew.gif" width=71 
            border=0></font></A></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
                  <TBODY>
                    <TR> 
                      <TD height=20></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
     
    </TR>
  </TBODY>
</TABLE>
</div>
<!--#include file="foot.asp"-->