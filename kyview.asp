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
<style fprolloverstyle>A:hover {color: red; font-weight: bold}
<!--
.shuye       { font-family: 微软雅黑;  font-size:16pt  font-weight: bold}
.cp          { font-family: 微软雅黑; font-size: 14pt; font-weight: bold }
.shuye1       { font-family: 微软雅黑;  font-size:12pt }

.index1      { font-family: 宋体; font-size: 14px }
.index2      { font-size: 10pt; font-family: 微软雅黑 }
-->

</style>
<style>
<!--
.xuesheng    { font-size: 12pt; font-weight: bold }
.title        { font-size: 14pt }
-->
</style><style>
<!--
.as          { font-size: 12pt; font-family: 微软雅黑; font-weight: bold }
-->
</style>
</head>
<BODY topMargin=0 leftmargin="0" link="#666666" vlink="#666666" alink="#666666">
<!--#include file="top1.asp"-->
<TABLE cellSpacing=0 cellPadding=0 width=1250 align=center bgColor=#ffffff  vAlign=top style="BORDER-TOP: #808080 1px solid;BORDER-LEFT: #808080 1px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 1px solid;"
border=0>
  <TBODY>
    <TR> 
      <TD width="1250"><TABLE cellSpacing=0 cellPadding=0 width=1250 align=center border=0>
          <TBODY>
            <TR> 
              <TD width="262" vAlign=top> 
				<TABLE width=213 border=0 cellPadding=0 cellSpacing=0>
                  <TBODY>
                    <TR> 
                      <TD height="28" align="center" bgcolor="#FFFFFF">
                <table border="0" width="240" cellspacing="0" height="435">
					<tr>
                      <TD height="49" align="center" bgcolor="#004993">
						<p align="left" class="title">
										<font color="#006699" face="微软雅黑">　</font><span style="font-weight: 700"><font color="#FFFFFF" face="微软雅黑">科研研究生</font></span></TD>
                    </tr>
					<tr>
                      <TD height="99" bgcolor="#FFFFFF">
                                <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                              border=0 height="87" >
                                  <TBODY>
                                    <TR> 
                                      <TD align="left" height="48" background="bj.jpg">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 7px">
										<font color="#004993" face="微软雅黑">　</font><font color="#004993"><img border="0" src="images/about1.jpg" width="19" height="27"><a style="text-decoration: none" href="kylist.asp?type=64&type2=74"><font face="微软雅黑" color="#004993">科研工作</font></a></font></TD>
                                      
                                    </TR>
                                    <TR> 
                                      <TD align="left" height="39" background="bj.jpg">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 7px">
										<font color="#004993" face="微软雅黑">　</font><font color="#004993"><img border="0" src="images/about1.jpg" width="19" height="27"><a style="text-decoration: none" href="kylist.asp?type=64&type2=75"><font face="微软雅黑" color="#004993">研究生工作</font></a></font></TD>
                                      
                                    </TR>
                                  </TBODY>
                                </TABLE>
								</TD>
                    </tr>
					<tr>
						<td bgcolor="#FFFFFF"> 
						<font color="#666666">
						<table border="0" width="250" cellspacing="0" cellpadding="0" height="40" bgcolor="#004993">
							<tr>
								<td height="49">
								<p class="title" align="left">
										<font color="#006699" face="微软雅黑">　</font><span style="font-weight: 700"><font color="#FFFFFF" face="微软雅黑">校 
								内 连 接</font></span></td>
							</tr>
				</table>
						<table border="0" width="101%" cellspacing="0">
							<tr>
								<td> 
                <table border="0" width="247" cellspacing="0" height="253">
					<tr>
						<td bgcolor="#FFFFFF"> 
                <table border="0" width="99%" cellspacing="0" cellpadding="0">
					<tr style="font-size: 15px; border: 0px none; margin: 0px; padding: 0px">
													<td style="font-size: 15px; border: 0px none; margin: 0px; padding: 0px" height="35"> 
													<font face="微软雅黑"> <%
sql2="select *  from frendlink where reply='-1' order by id desc "
Set oRS= Server.CreateObject("ADODB.recordset")
oRS.Open sql2,conn,1,3
 if NOT oRS.EOF then 
%>
                        </font>
                        <TABLE border="0" align="left" cellpadding="0" cellSpacing=0>
                          <%
 do while not oRS.eof
  if NOT oRS.EOF then 
   %>
                          <tr> 
                            <% for m=1 to 1
  if NOT oRS.EOF then
%>
                            <td valign="middle"> <div align="left"> 
                                <TABLE border=0 align="left" cellPadding=0 cellSpacing=0 >
                                  <TBODY>
                                    <TR> 
                                      <TD  height=25 bgcolor="#FFFFFF">
										<p style="margin-top: 0; margin-bottom: 0">
										<font color="#004993" face="微软雅黑">　</font><font color="#004993"><img border="0" src="images/about1.jpg" width="19" height="27"></font><font face="微软雅黑" color="#008080">&nbsp; </font>
										<font face="微软雅黑">
										<a href="<% =ors("weburl")  %>" target="_blank" style="text-decoration: none"><%=left(oRS("webname"),16)%></a> 
                                      	</font> 
                                      </TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                              </div></td>
                            <% oRS.MoveNEXT
end if
next %>
                          </tr>
                          <%
end if
loop
%>
                        </table>
                        							<font face="微软雅黑">
                        <%
else Response.Write ("暂无内容.") end if
oRS.close 
Set oRS = Nothing
%>

													　</font></td>
												</tr>					
					</table>
						</td>
					</tr>
					<tr>
                      <TD height="49" bgcolor="#FFFFFF">
                                　</TD>
                    </tr>
					<tr>
						<td bgcolor="#FFFFFF"> 
              	　</td>
					</tr>

					<tr>
						<td bgcolor="#FFFFFF" width="245" height="130" align="left"> 
              	　</td>
					</tr>

					</table>
              					</td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
              		</TD>
                    </TR>
                  </TBODY>
                </TABLE>
				</TD>
              <TD width="4" bgcolor="#C0C0C0"></TD>
              <TD width="992" vAlign=top> 
				<div align="center">
				<TABLE cellSpacing=0 cellPadding=0 width="992" 
border=0 height="38">
                  <TBODY>
                    <TR> 
                      <TD width="4" height="34" align="center" style="BORDER-BOTTOM: #CFCFCF 1px dashed">　</TD>
                      <TD width="992"  style="BORDER-BOTTOM: #CFCFCF 1px dashed" height="34" align="left">
						<p class="title"><b><font color="#205081">
						<img border="0" src="fl/jiantou.jpg" width="17" height="14"></font></b><font face="微软雅黑"><SPAN style="FONT-WEIGHT: bold;"><font color="#004993">首页</font></SPAN><b><font color="#004993"> 
                        &gt;&gt;<%=bigclass%>科研研究生&gt;&gt;正文内容 </font></b></font></TD>
                      <TD width="197" style="BORDER-BOTTOM: #CFCFCF 1px dashed">　</TD>
                    </TR>
                    <TR> 
                      <TD height="4" align="center"></TD>
                      <TD height="4" colspan="2" align="center" background="images/lanmufeng.gif"></TD>
                    </TR>                  </TBODY>
                </TABLE>
                </div>
                <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
                  <TBODY>
                    <TR> 
                      <TD height=8></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE width=992 
border=0 align="center" cellPadding=0 cellSpacing=0 class=xzr id=table4>
                  <TBODY>
                    <TR> 
                      <TD height=45 colspan="4"> <div align="center"><B><SPAN style="FONT-SIZE: 12px; LINE-HEIGHT: 35px; LETTER-SPACING: 2px">
						<FONT size=3 color="#5E5E5E" face="微软雅黑"> 
                          <%=title %> </FONT></SPAN></B></div></TD>
                    </TR>
                    <TR> 
                      <TD height=9 colspan="4"> <HR width="100%" color=#c0c0c0 SIZE=1></TD>
                    </TR>
                    <tr>
                      <TD height=30 colspan="2"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td height="50"> <div align="right">
								<p style="margin-top: 7px; margin-bottom: 0">
								<font color="#333333" face="微软雅黑">发布时间：<%=infotime%>&nbsp;&nbsp;　　浏览：<%=hits%></font></div></td>
                          </tr>
                        </table></TD>
                      <TD height=30 colspan="2">        <div class="share" style="margin-top: 20px;"><!-- JiaThis Button BEGIN -->
			<div class="jiathis_style_24x24">
			<a class="jiathis_button_tsina"></a>
			<a class="jiathis_button_weixin"></a>
			<a class="jiathis_button_qzone"></a>
			<a class="jiathis_button_tqq"></a>
			<a class="jiathis_button_renren"></a>
			<a class="jiathis_button_feixin"></a>
			<a class="jiathis_button_cqq"></a>
			<a href="http://www.jiathis.com/share?uid=1981363" class="jiathis jiathis_txt jiathis_separator jtico jtico_jiathis" target="_blank"></a>
			<a class="jiathis_counter_style"></a>
			</div>
<script type="text/javascript">
var a = document.getElementById("fx_article_title").innerHTML;
console.log(a);
			var jiathis_config={
				data_track_clickback:true,
				summary:"",
				shortUrl:false,
				hideMore:false,
                                 title: document.getElementById("fx_article_title").innerHTML
 
		}
			</script>
			                         <script type="text/javascript" ignoreapd="1" src="http://v3.jiathis.com/code/jia.js?uid=1981363" charset="utf-8"></script>
			<!-- JiaThis Button END -->
</div>   </TD>
                    </tr>
                    <TR> 
                      <TD height=30 colspan="4"> 　</TD>
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
                      <TD valign="bottom" width="342"><A href="javascript:window.top.close();">
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
      <TD width="10">　</TD>
    </TR>
  </TBODY>
</TABLE>
<!--#include file="foot.asp"-->