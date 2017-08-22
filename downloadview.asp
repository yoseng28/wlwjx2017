<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">

<!--#include file="conn.asp"-->
<%
id=trim(request("id"))
id=replace(id,"'","")
id=left(id,8)
	if id="" then '
  		response.write "<br><p align=center><font color='red'>对不起，你访问的文件不存在。</font></p>"
		response.end
	end if	
	if  not isNumeric(id) then
  		response.write "<br><p align=center><font color='red'>对不起，你访问的文件不存在。</font></p>"
		response.end
	end if
    if len(id)>8 then 
		response.write "<br><p align=center><font color='red'>对不起，你访问的文件不存在。</font></p>"
		response.end
	end if
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
<% 
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="update download set hits=hits+1 where id="&cstr(request("id"))
conn.execute sql
sql="select * from download where id="&id
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
response.Write("数据库出错")
else
title=rs("downName")
bigclass=rs("downBigClass")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% =sitetitle %>--<% =title %></title>
<meta name="keywords" content="<% =title %>,<% =webKeywords %>">
<meta name="description" content="<% =webDescription %>,<% =title %>">
<meta name="robots" content="index, follow">
<meta name="googlebot" content="index, follow">
<style fprolloverstyle>A:hover {color: #004993; font-weight: bold}
</style>
<style>
<!--
.xuesheng    { font-size: 12pt; font-weight: bold }
.title        { font-size: 14pt }
.shuye        { font-size: 11pt }

-->
</style>

</head>
<BODY topMargin=0 leftmargin="0" link="#004993" vlink="#0099FF" alink="#0099FF">
<!--#include file="top.asp"-->
<TABLE cellSpacing=0 cellPadding=0 width=1100 align=center bgColor=#ffffff 
border=0>
  <TBODY>
    <TR> 
      <TD width="5" bgcolor="#FFFFFF"></TD>
      <TD width="1080">
		<TABLE width=1234 border=0 align=center cellPadding=0 cellSpacing=0>
          <TBODY>
            <TR> 
              <TD width="149" vAlign=top background="images/class_bg4.gif" style="background-repeat: repeat-x"
 bgcolor="#FFFFFF"> 
             </TD>
              <TD width="4" rowspan="2"></TD>
              <TD width="924" vAlign=top style="BORDER-TOP: #C5CED7 1px solid;BORDER-LEFT: #C5CED7 1px solid;BORDER-right: #C5CED7 1px solid;BORDER-bottom: #C5CED7 1px solid;" rowspan="2"> 
                <TABLE cellSpacing=0 cellPadding=0 width="924" 
border=0>
                  <TBODY>
                    <TR> 
                      <TD width="10" height="29" align="center">　</TD>
                      <TD width="296" height="45"> <font color="#60616C"><b>
						<font face="微软雅黑"><%= rs("downName") %></font></b></font><b><font face="微软雅黑" color="#60616C">
						</font></b> 
                      </TD>
                      <TD width="404" align="right"><b><font face="微软雅黑">
						<font color="#0066CC">⊙ </font>
						<a href="http://www.ectgemlju.com" style="text-decoration: none">
						<font color="#0066CC">首页</font></a><font color="#0066CC"> 
                        &gt;&gt; <%= rs("downName") %> </font></font></b> </TD>
                      <TD width="14">　</TD>
                    </TR>
                    <TR> 
                      <TD height="5" colspan="4" align="center" background="images/lanmufeng.gif"></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=100 
                        align=center border=0>
                  <TBODY>
                    <TR> 
                      <TD>　</TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width="98%" 
                              align=center border=0>
                  <TBODY>
                    <TR> 
                      <TD bgColor=#dfff80> <TABLE width="100%" 
                                border=0 cellPadding=0 cellSpacing=1 bgcolor="f1f1f1">
                          <TBODY>
                            <TR bgcolor="#FFFFFF"> 
                              <TD width="18%" height=42> <DIV align=center>
								<font face="微软雅黑" color="#60616C"><strong>文件名称：</strong></font></DIV></TD>
                              <TD width="82%" bgcolor="#FFFFFF">
								<font face="微软雅黑" color="#60616C">&nbsp;<STRONG> 
                                <% =RS("downName") %>
                                </STRONG></font></TD>
                            </TR>
                            <TR bgcolor="#FFFFFF"> 
                              <TD height=42> <DIV align=center>
								<font face="微软雅黑" color="#60616C"><strong>文件大小：</strong></font></DIV></TD>
                              <TD><font face="微软雅黑" color="#60616C">&nbsp; 
                                <% =RS("Downdaxiao") %>
                              </font>
                              </TD>
                            </TR>
                            <TR bgcolor="#FFFFFF"> 
                              <TD height=42> <DIV align=center>
								<font face="微软雅黑" color="#60616C"><strong>关注指数：</strong></font></DIV></TD>
                              <TD><font face="微软雅黑" color="#60616C">&nbsp; 浏览：<%= rs("hits") %> 次</font></TD>
                            </TR>
                            <TR bgcolor="#FFFFFF"> 
                              <TD height=42> <DIV align=center>
								<font face="微软雅黑" color="#60616C"><strong>发布时间：</strong></font></DIV></TD>
                              <TD><font face="微软雅黑" color="#60616C">&nbsp; 
                                <% =RS("downTime") %>
                              </font>
                              </TD>
                            </TR>
                            <TR bgcolor="#FFFFFF"> 
                              <TD height=55> <div align="center">
								<font face="微软雅黑" color="#60616C"><strong>文件下载：</strong></font></div></TD>
                              <TD valign="middle"><font face="微软雅黑">
								<font color="#60616C">&nbsp; </font> 
								<a href="<% =RS("downlink") %>" style="text-decoration: none">
								<font color="#60616C">点击下载</font></a></font></TD>
                            </TR>
                          </TBODY>
                        </TABLE></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=100 
                        align=center border=0>
                  <TBODY>
                    <TR> 
                      <TD>　</TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE width=910 height=159 border=0 
                        align=center cellPadding=0 cellSpacing=1 bgcolor="#FFFFFF">
                  <TBODY>
                    <TR> 
                      <TD height=35 bgcolor="#ebebeb"><STRONG>&nbsp;&nbsp;&nbsp;<font color="#373737">&nbsp;详细说明：</font></STRONG></TD>
                    </TR>
                   
                    <TR vAlign=top> 
                      <TD 
                              height=62><%=rs("downContent") %></TD>
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
                <TABLE width=910 height="25" 
border=0 align=center cellPadding=0 cellSpacing=0 bgColor=#ffffff>
                  <TBODY>
                    <TR> 
                      <TD height="35" align="center" valign="bottom"> <font color="#0066CC"><strong>上一
						条：</strong></font> 
                        <% 
						Set srs=Server.CreateObject("ADODB.RecordSet") 
						sqls="select * from download where id<"&id&" order by id desc"
srs.Open sqls,conn,1,1
if srs.eof and srs.bof then
response.Write("这是第一条信息了！")
else
					  %>
                        <a href="downloadview.asp?id=<%= srs("id") %>" style="text-decoration: none"> 
                        <% =srs("downName") %>
                        </a> 
                        <%
						end if
						srs.close
		                set srs=nothing
						%>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font color="#0066CC"><strong>下一
						条：</strong></font> 
                        <% 
						Set xrs=Server.CreateObject("ADODB.RecordSet") 
						sqlx="select * from download where id>"&id&" order by id asc"
xrs.Open sqlx,conn,1,1
if xrs.eof and xrs.bof then
response.Write("这是最后一条信息了！")
else
					  %>
                        <a href="downloadview.asp?id=<%= xrs("id") %>" style="text-decoration: none"> 
                        <% =xrs("downName") %>
                        </a> 
                        <%
						end if
						xrs.close
		                set xrs=nothing
						%>
                      </TD>
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
                <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
                  <TBODY>
                    <TR> 
                      <TD height=8></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
            <TR> 
              <TD width="149" vAlign=top background="images/class_bg4.gif" style="background-repeat: repeat-x"
 bgcolor="#FFFFFF" height="481"> 
               
                <table border="0" width="306" cellspacing="0" height="436"   style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 0px solid;BORDER-right: #808080 0px solid;BORDER-bottom: #808080 0px solid;">
					<tr>
                      <TD height="49" align="center" bgcolor="#0099FF">
						<p align="left" class="title" style="margin-top: 0; margin-bottom: 0">
										<font color="#006699" face="微软雅黑">　</font><span style="font-weight: 700"><font color="#FFFFFF" face="微软雅黑">下载中心</font></span></TD>
                    </tr>
					<tr>
                      <TD height="99" bgcolor="#FFFFFF">
                                <TABLE cellSpacing=0 cellPadding=0 width="91%" 
                              border=0 height="87" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px" >
                                  <TBODY style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
                                    <TR style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"> 
                                      <TD align="left" height="48" background="bj.jpg" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 7px">
										<font color="#004993" face="微软雅黑" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">　</font><font face="微软雅黑" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"><img border="0" src="images/about1.jpg" width="19" height="27" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"><span class="Apple-converted-space">&nbsp;</span></font><font face="微软雅黑"><a style="text-decoration: none" href="download.asp?type=79&type2=40">课件资料</a></font></TD>
                                      
                                    </TR>
                                    <TR style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"> 
                                      <TD align="left" height="13" background="bj.jpg" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 7px">
										<font color="#004993" face="微软雅黑" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">　</font><font face="微软雅黑" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"><img border="0" src="images/about1.jpg" width="19" height="27" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"><span class="Apple-converted-space">&nbsp;</span></font><font face="微软雅黑" style="font-size: 16px; "><a style="text-decoration: none" href="download.asp?type=79&type2=41">审报资料</a></font></TD>
                                      
                                    </TR>
                                    <TR style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"> 
                                      <TD align="left" height="40" background="bj.jpg" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
										<font color="#004993" face="微软雅黑" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">　</font><font face="微软雅黑" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"><img border="0" src="images/about1.jpg" width="19" height="27" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"><span class="Apple-converted-space">&nbsp;</span></font><font face="微软雅黑" style="font-size: 16px; "><a href="download.asp?type=79&type2=42" style="text-decoration: none; font-weight: 700">其它资料</a></font></TD>
                                      
                                    </TR>
                                    <TR style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"> 
                                      <TD align="left" height="40" background="bj.jpg" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
										　</TD>
                                      
                                    </TR>
                                  </TBODY>
                                </TABLE>
								</TD>
                    </tr>
					<tr>
						<td bgcolor="#FFFFFF" height="260"> 
						<font color="#666666">
						<table border="0" width="302" cellspacing="0" cellpadding="0" height="40" bgcolor="#004993">
							<tr>
								<td height="50" bgcolor="#0099FF">
								<p class="xuesheng" style="margin-top: 0; margin-bottom: 0" align="left">
										<font color="#006699" face="微软雅黑">　</font><span style="font-weight: 700"><font color="#FFFFFF" face="微软雅黑">团队动态</font></span></td>
							</tr>
							<tr>
								<td height="10" bgcolor="#FFFFFF">
								　</td>
							</tr>
				</table>
						<table border="0" width="101%" cellspacing="0">
							<tr>
								<td> 
                <table border="0" width="252" cellspacing="0" height="253">
					<tr>
						<td bgcolor="#FFFFFF"> 
				<div align="center" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
								<table border="1" bgcolor="#FFFFFF" cellspacing="0" bordercolor="#CFCFCF" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" width="255">
									<tbody style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
									<tr style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
										<td style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
							<body link="#004993" vlink="#0099FF" alink="#0099FF">
<div id="demo" style="border:0px none; margin:0px; padding:0px; overflow:hidden;height:209px;width:298px; font-size:15px">
<ul id="demo1" style="height: auto; text-align: left; border: 0px none; margin: 0px; padding: 0px; list-style-type:none; font-size:15px">  
<font color="#000F1A" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" face="微软雅黑">
                <b style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
                <%
sql2="select *  from content where bigclassname='"&index1&"' order by id desc "
Set oRS= Server.CreateObject("ADODB.recordset")
oRS.Open sql2,conn,1,3
 if NOT oRS.EOF then 
for i=1 to 20
  if NOT oRS.EOF then
   %>
                </b>
                </font>
                <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0 height="24" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
                  <TBODY style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
                    <TR style="border: 0px none; margin: 0px; padding: 0px; font-size:15px"> 
                      <TD class=sd13 height=24>
						<p style="margin-top: 0; margin-bottom: 0">
						<font face="微软雅黑" style="font-size: 15px; border: 0px none; margin: 0px; padding: 0px">
						<font color="#000F1A" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">·  
						</font>  
						<A 
                              title=<%=oRS("title")%> 
                              href="newsview.asp?id=<%=oRS("id")%>" 
                               style="border:0px none; margin:0px; padding:0px; text-decoration: none; font-size:15px">
						<font color="#000F1A" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
						<%=left(oRS("title"),15)%></font></A></font></TD>
						<TD class=sd13 height=27 width="56">
						<p align="center" style="margin-top: 0; margin-bottom: 0">
						<font color="#000F1A" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" face="微软雅黑">(<% =month(ors("infotime")) %>.<% =day(ors("infotime")) %>)</font></TD>
                    </TR>

                  </TBODY>
                </TABLE>
                <font color="#FFFF00" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" face="微软雅黑">
                <b style="border: 0px none; margin: 0px; padding: 0px; font-size:15px"><%
oRS.MoveNEXT
end if
next
else Response.Write ("暂无内容.") end if
oRS.close 
Set oRS = Nothing
%>
              </b>
              </font></ul> 
<div id="demo2" style="height: auto; text-align: left; border: 0px none; margin: 0px; padding: 0px; font-size:15px">
	<p style="margin-top: 0; margin-bottom: 0"></div>
</div> 
							<font color="#FFFF00" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" face="微软雅黑"> 
							<b style="border: 0px none; margin: 0px; padding: 0px; font-size:15px"> 
<script> 
var speed=40 
var demo=document.getElementById("demo"); 
var demo2=document.getElementById("demo2"); 
var demo1=document.getElementById("demo1"); 
demo2.innerHTML=demo1.innerHTML 
function Marquee(){ 
if(demo2.offsetTop-demo.scrollTop<=0) 
  demo.scrollTop-=demo1.offsetHeight 
else{ 
  demo.scrollTop++ 
} 
} 
var MyMar=setInterval(Marquee,speed) 
demo.onmouseover=function() {clearInterval(MyMar)} 
demo.onmouseout=function() {MyMar=setInterval(Marquee,speed)} 
</script> 
</body>
							</b>
							</font>
</td>
									</tr>
								</table>
							</div>
                　</td>
					</tr>

					<tr>
						<td bgcolor="#FFFFFF" width="248" height="130" align="left"> 
              	　</td>
					</tr>

					</table>
              					</td>
							</tr>
						</table></td>
					</tr>
					<tr>
						<td bgcolor="#FFFFFF"> 
										　</td>
					</tr>
				</table>
              　</TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    
    </TR>
  </TBODY>
</TABLE>
<% 
		end if
		rs.close
		set rs=nothing
		 %><!--#include file="foot.asp"-->