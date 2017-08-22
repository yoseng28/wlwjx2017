<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">

<!--#include file="conn.asp"--> 
<!--#include file="inc/function.asp"--> 
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% =sitetitle %></title>
<meta name="keywords" content="<% =webKeywords %>">
<meta name="description" content="<% =webDescription %>">
<meta name="robots" content="index, follow">
<meta name="googlebot" content="index, follow">
<style media="screen" type=text/css >
<!--
html,body{margin:0;text-align:left;over-flow:hidden;height:100%;width:100%;}
.anli        { font-family: 微软雅黑; font-size: 16pt; font-weight: bold }
.text        { font-family: 微软雅黑; font-size: 11pt }
UL{list-style-type:none; margin:0px;}
/* 标准盒模型 */
.ttl{height:43px;}
.ctt{height:auto;padding:0px;clear:both;border:0px solid #064ca1;border-top:0;text-align:left;}
.w936{margin:2px 0;clear:both;width:1000px;/*这里调整整个滑动门的宽度*/}
/* TAB 切换效果 */
.tb_{background-image: url('../fl/1001271.gif'); background-repeat: color:#FFFFFF; repeat-x;padding-left:0px}
.tb_ ul{height:43px;}
.tb_ li{float:left;height: 43px;line-height:1.9;width:112px;cursor:pointer;}
/* 用于控制显示与隐藏的css类 */
.normaltab   { background-image:url('../fl/1001272.gif'); background-repeat: no-repeat; color:#FFFFFF ;font-weight:bold padding-left:1px}
.hovertab    { background-image: url('../fl/1001273.gif'); background-repeat: no-repeat; color:#FFFFFF; font-weight:bold  padding-left:0px}
.dis{display:block;}
.undis{display:none;}
-->

</style>

<style fprolloverstyle>A:hover {color: #FF0000; font-weight: bold}
</style>
<style>
<!--
.sg1          { font-size: 12pt; font-family: 微软雅黑; font-weight: bold }
.sg2          { font-size: 12pt; font-family: 微软雅黑; }
.sg3          { font-size: 12pt; font-family: 微软雅黑; }

-->
</style>
</head>
<BODY bgcolor="#FFFFFF" topMargin=0 leftmargin="0" link="#4D4D4D" vlink="#4D4D4D" alink="#4D4D4D" style="text-align: left">
<!--#include file="top.asp"-->
<div align="center">
<table border="0" width="1107" cellspacing="0" vAlign=top style="BORDER-TOP: #BBCFF0 0px solid;BORDER-LEFT: #BBCFF0 0px solid;BORDER-right: #BBCFF0 0px solid;BORDER-bottom: #BBCFF0 0px solid;">
	<tr>
		<td width="156" height="42" bgcolor="#0099FF">
		<p align="left" class="anli" style="line-height: 200%; margin-top: 0; margin-bottom: 0">
		<font face="微软雅黑" color="#D2E9FF">&nbsp;<img border="0" src="fl/home.png" width="18" height="18"><b> 
		团队动态</b></font></td>
		<td width="146" height="42" bgcolor="#0099FF">
		<p align="right" style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<a href="news.asp?type=31">
										<img border="0" src="images/cw_more.gif" width="30" height="11"></a></td>
		<td width="3" rowspan="2">
		<p style="line-height: 200%; margin-top: 0; margin-bottom: 0"></td>
		<td width="463" bgcolor="#0099FF">
		<p class="anli" style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="left">
		<font face="微软雅黑" color="#D2E9FF">&nbsp;<img border="0" src="fl/home.png" width="18" height="18"></font><b><font face="微软雅黑" color="#FFFFFF"> </font>
		<font face="微软雅黑" color="#D2E9FF">团队简介</font></b></td>
		<td width="2" rowspan="2">
		<p style="margin-top: 0; margin-bottom: 0; line-height:200%"></td>
		<td width="163" height="22" bgcolor="#0099FF">
		<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0" class="anli">
										&nbsp;<font face="微软雅黑" color="#D2E9FF"><img border="0" src="fl/home.png" width="18" height="18"> 
										</font><font color="#D2E9FF">图文动态</font></td>
		<td width="162" height="22" bgcolor="#0099FF">
		<p align="right">
										　</td>
	</tr>
	<tr>
		<td width="302" style="BORDER-TOP: #1649A2 1px solid;BORDER-LEFT: #1649A2 1px solid;BORDER-right: #1649A2 1px solid;BORDER-bottom: #1649A2 1px solid;" colspan="2">
													<p align="left" style="margin-top: 0; margin-bottom: 0; line-height:200%" class="text">
													<div align="left">
				<table border="0" width="100%" id="table75" cellspacing="0" height="248" bgcolor="#FFFFFF">
					<tr>
						<td>
						<TABLE cellSpacing=0 cellPadding=0 width="290" border=0 bordercolor="#D01E0E" >
                                  <TBODY>
                                    <TR> 
                                      <TD><font color="#333300">
										<MARQUEE direction=up WIDTH=295 HEIGHT=246 scrollamount=2 id=info onMouseOver=info.stop() onMouseOut=info.start() ><%
sql2="select *  from content where bigclassname='"&index1&"' order by id desc "
Set oRS= Server.CreateObject("ADODB.recordset")
oRS.Open sql2,conn,1,3
 if NOT oRS.EOF then 
for i=1 to 7
  if NOT oRS.EOF then
   %>
                                              </font>
                                              <div align="left">
                                             <TABLE class=sg3 id=table79 cellSpacing=0 
                        cellPadding=0 width="100%" border=0 height="30">
																<TBODY>
																	<TR>
																		<TD width="100%" height="30">&nbsp;<IMG height=4 
                              src="images/icon03.gif" width=4>&nbsp;<A style="text-decoration: none"
                              title=<%=oRS("title")%> 
                              href="view.asp?id=<%=oRS("id")%>" 
                              target=_self><%=left(oRS("title"),16)%></A> </TD>
																		
																	</TR>
																</TBODY>
															</TABLE>                                              </div>
                                              <font color="#333300">
                                              <%
oRS.MoveNEXT
end if
next
else Response.Write ("暂无内容.") end if
oRS.close 
Set oRS = Nothing
%>
                                            </MARQUEE>  </font>
                                              </TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
						</td>
					</tr>
				</table>
			</div>
</td>
		<td width="463">
		<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0" style="color: rgb(0, 0, 0); font-family: Arial, 微软雅黑; font-size: 12px; font-style: normal; font-variant: normal; font-weight: normal; letter-spacing: normal; line-height: 15.6px; orphans: auto; text-indent: 0px; text-transform: none; white-space: normal; widows: 1; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255);">
			<tr>
				<td>
				<font color="#333333">
				<span style="font-size: 14px; line-height: 25px;">
				<p class="sg2" style="text-indent: 0; margin-left:0cm; margin-right:0cm; margin-top:0cm; margin-bottom:0pt" align="left">
				<font color="#000000"> 
				<span style="font-size: 12pt"> <% =companycontent %>
</span><span style="font-size: 12pt; line-height:25px"><font color="#003399" face="微软雅黑">【</font><font color="#006994" face="微软雅黑"><a style="text-decoration: none" href="about.asp?type=6"><font color="#FF0000">详细介绍</font></a><font color="#003399" face="微软雅黑">】</font></font></span></font></span></td>
			</tr>
		</table>
		</td>
		<td width="323" height="21" style="BORDER-TOP: #1649A2 1px solid;BORDER-LEFT: #1649A2 1px solid;BORDER-right: #1649A2 1px solid;BORDER-bottom: #1649A2 1px solid;" colspan="2">
				<div align="center" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
								<table border="1" bgcolor="#FFFFFF" cellspacing="0" bordercolor="#CFCFCF" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" width="291">
									<tbody style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
									<tr style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
										<td style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
							<body>
						
<div >
<p style="margin-top: 0; margin-bottom: 0">
<IFRAME name=I1 
                              marginWidth=1 marginHeight=0 
                              src="gundong.asp" 
                              frameBorder=no width=320 scrolling=no height=295
                              DESIGNTIMESP="1840"></IFRAME>
</td>
									</tr>
								</table>
							</div>
                </td>
	</tr>
	</table>
</div>

<div align="center" >
	<div align="center">
<table border="0" width="1105" cellspacing="0" cellpadding="0">
	<tr>
		<td height="60" colspan="2"></td>
		<td height="60" width="952">
		<p class="anli" style="line-height: 200%; margin-left: 10px; margin-top: 0; margin-bottom: 0" align="left">
		<font color="#0099FF">
		<img border="0" src="fl/lanmqiantb.gif" width="8" height="15"> 
		<a href="zxktsp.asp?type=85" style="text-decoration: none">最新视频</a></font></td>
		<td height="60" width="148">
		<p align="right">
		<a href="zxktsp.asp?type=85" style="text-decoration: none">更多</a></td>
		<td height="60" width="1"></td>
	</tr>
	<tr>
		<td width="4" >
		　</td>
		<td   style="BORDER-TOP: #CDCED0 1px solid;BORDER-LEFT: #CDCED0 1px solid;BORDER-right: #CDCED0 1px solid;BORDER-bottom: #CDCED0 1px><font color="#006994" face="微软雅黑" width="545" colspan="3">
		<div align="center">
                <TABLE cellSpacing=0 cellPadding=0 width=1099 border=0 height="23%" style="border: 0 none">
                    <TR> 
                      <TD align="center" style="font-size: 10pt; line-height: 22px; font-family: arial" >
						<div id=demo style="OVERFLOW: hidden; WIDTH: 1093px; align: center; height:220px">
  <div align="center">
  <table cellspacing="0" cellpadding="0" 
border="0" width="279" style="border: 0 none">
    <tbody>
      <tr>
        <td id="marquePic1" valign="top" width="274" style="font-size: 10pt; line-height: 22px; font-family: arial">
		
  <table width="100%" border="0" style="border: 0 none" >
																		<tr>
																			<td style="font-size: 10pt; line-height: 22px; font-family: arial" width="71%">
                                										<TABLE border="0" align="left" cellpadding="0" cellSpacing=0 style="border: 0 none"><%
sql2="select *  from Product where ok=true order by id desc "
Set oRS= Server.CreateObject("ADODB.recordset")
oRS.Open sql2,conn,1,3
 if NOT oRS.EOF then 
%>
										<% for i=1 to 1 %>	<tr> 
                                    <% for m=1 to 10
  if NOT oRS.EOF then
%>
                                    												<td valign="middle" style="font-size: 10pt; line-height: 22px; font-family: arial">
																					<div align="center">
																						<div align="center"> 
                                        
                                      <TABLE border=0 cellPadding=0 cellSpacing=0 width="220">
                                        <TBODY>
                                            <TR> 
                                            <TD width="220" height=25> 
                                              <font face="微软雅黑"> 
                                              <% if oRS("ProductImages")="" then %>
                                              <div align="center">
                                              <TABLE width="204" height="120" border=0 
                                cellPadding=0 cellSpacing=0>
                                                  <TBODY>
                                                    <TR> 
                                                      <TD width="204">
														<p style="margin-top: 0; margin-bottom: 0; line-height:200%" align="center">
														<a href="spview.asp?id=<%= oRS("id") %>">
														<img src="../images/wupic.jpg" width="195" height="161" border="0" style="border: 0px solid #808080; padding: 1px"></a></font></TD>
                                                    </TR>
                                                    <TR align="center"> 
                                                      <TD width="204" height="20" align="center">
														<p  class="sg1" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><a href="spview.asp?id=<%= oRS("id") %>" style="text-decoration: none"> 
                         																					 
                                  <% =left(oRS("Productname"),18)%>
														　</TD>
                                                    </TR>
                                                  </TBODY>
                                                </TABLE>
                                                
                                                
                                              </div>

                                                                                              
                                                
                                              <font face="微软雅黑">

                                                                                              
                                                
                                              <% else %>
                                              														</font>
                                              														<div align="center">
																										<TABLE width="204" height="100%" border=0 cellSpacing=0 style="border-style: none; border-color: inherit">
																											<TBODY>
																												<TR>
																														<TD width="202" style="font-size: 10pt; line-height: 22px; font-family: arial" height="16%">
																													<p style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="center">
																													<font face="微软雅黑">
																													<font color="#663300"><a href="spview.asp?id=<%= oRS("id") %>"><img src="<% =oRS("ProductImages")%>" width="195" height="163" border="0" style="padding:3px; border:0px ridge #808080; "></a></font></b></font></TD>
																												</TR>
																												<TR>
																													<TD width="202" style="font-size: 10pt; line-height: 22px; font-family: arial" align="center" height="50"><p style="line-height: 150%; margin-left: 16px; margin-top: 0; margin-bottom: 0"><a href="spview.asp?id=<%= oRS("id") %>" style="text-decoration: none"> 
                                
                                  																					 
                                  																					<font face="微软雅黑"> 
                                
                                  																					 
                                  <% =left(oRS("Productname"),18)%></font></TD>																								</TR>
																											</TBODY>
																										</TABLE></div>
                                                <font face="微软雅黑">
                                                <% end if %>
                                                													</font>
                                                </font>
                                                </TD>
																								</TR>
																							</TBODY>
																						</TABLE></div></td>
                                   <% oRS.MoveNEXT
end if
next %>
                  </tr><%next %>
                </table>
                <font color="#BFDFFF" face="微软雅黑">
                <%
else Response.Write ("暂无内容.") end if
oRS.close 
Set oRS = Nothing
%></td>
	</tr>
</table>	</td>
        <td id="marquePic2" valign="top" style="font-size: 10pt; line-height: 22px; font-family: arial">
		<p style="margin-top: 0; margin-bottom: 0"></td>
      </tr>
    </tbody>
  </table>
	</div>
</div>
<script type=text/javascript> 
var speed=30 
marquePic2.innerHTML=marquePic1.innerHTML 
function Marquee(){ 
if(demo.scrollLeft>=marquePic1.scrollWidth){ 
demo.scrollLeft=0 
}else{ 
demo.scrollLeft++ 
}} 
var MyMar=setInterval(Marquee,speed) 
demo.onmouseover=function() {clearInterval(MyMar)} 
demo.onmouseout=function() {MyMar=setInterval(Marquee,speed)} 
</script>
					</td>
				</tr>
			</table>
                                                                      
                              </div></td>
	</tr>
</table>
	</div>
	</div><!--#include file="foot.asp"-->