<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">

<!--#include file="conn.asp"-->
<% 
'获取收入参数
bigclassid=trim(request("type"))
smallclassid=trim(request("type2"))
bigclassid=replace(bigclassid,"'","")
smallclassid=replace(smallclassid,"'","")
bigclassid=left(bigclassid,8)
smallclassid=left(smallclassid,8)
	if bigclassid="" then '
  		response.write "<br><p align=center><font color='red'>对不起，你访问的文件不存在1。</font></p>"
		response.end
	elseif  not isNumeric(bigclassid) then
  		response.write "<br><p align=center><font color='red'>对不起，你访问的文件不存在2。</font></p>"
		response.end
	else
        sqlb="select BigClassName from BigClass where bigclassid="&bigclassid&""
        set rsb=server.createobject("ADODB.Recordset")
        rsb.open sqlb,conn,1,1
        if not rsb.eof then
            bigclass=rsb("BigClassName")
        end if
        rsb.Close
        set rsb=nothing
	end if
	if smallclassid<>"" then
	if  not isNumeric(smallclassid) then
	    response.write "<br><p align=center><font color='red'>对不起，你访问的文件不存在3。</font></p>"
		response.end
	else
	    sqls="select SmallClassName from SmallClass where SmallClassID="&SmallClassID&""
        set rss=server.createobject("ADODB.Recordset")
        rss.open sqls,conn,1,1
        if not rss.eof then
          smallclass=rss("SmallClassName")
        else
        response.write "<br><p align=center><font color='red'>对不起，你访问的文件不存在4。</font></p>"
        response.end
        end if
        rss.Close
        set rss=nothing
    end if
    end if
'获取产品显示方式
sqlxs="select *  from BigClass where ClassType='product.asp' and BigClassName='"&bigclass&"' order by BigClassID desc "
Set oRSxs= Server.CreateObject("ADODB.recordset")
oRSxs.Open sqlxs,conn,1,3
 if NOT oRSxs.EOF then 
xsfs=oRSxs("xsfs")
else
		response.write "<br><p align=center><font color='red'>对不起，你访问的文件不存在。</font></p>"
		response.end
end if
        oRSxs.Close
        set oRSxs=nothing
   %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% =sitetitle %>--<% =bigclass %></title>
<meta name="keywords" content="<% =bigclass %>,<% =webKeywords %>">
<meta name="description" content="<% =webDescription %>,<% =bigclass %>">
<meta name="robots" content="index, follow">
<meta name="googlebot" content="index, follow">
<style fprolloverstyle>A:hover {color: #006699; font-weight: bold}
</style>
<style>
<!--
.xuesheng    { font-size: 12pt; font-weight: bold }
.title        { font-size: 14pt }
.shuye        { font-size: 11pt }

-->
</style>

</head>
<BODY  topMargin=0 leftmargin="0" link="#666666" vlink="#666666" alink="#666666" class="shuye" bgcolor="#FFFFFF">
<!--#include file="top.asp"-->
<div align="center">
<TABLE cellSpacing=0 cellPadding=0 width=1080 bgColor=#ffffff align="center"
border=0>
  <TBODY>
    <TR> 
      <TD width="1080">
		<div align="center">
			<TABLE cellSpacing=0 cellPadding=0 width=1108 border=0 bgcolor="#FFFFFF">
          <TBODY>
            <TR> 
              <TD width="155" vAlign=top style="BORDER-TOP: #C5CED7 0px solid;BORDER-LEFT: #E3E3E3 1px solid;BORDER-right: #C5CED7 1px solid;BORDER-bottom: #C5CED7 0px solid;" rowspan="3"> 
               
                <table border="0" width="260" cellspacing="0" height="436"   style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 0px solid;BORDER-right: #808080 0px solid;BORDER-bottom: #808080 0px solid;">
					<tr>
                      <TD height="49" align="center" bgcolor="#0099FF">
						<p align="left" class="title" style="margin-top: 0; margin-bottom: 0">
										<font color="#006699" face="微软雅黑">　</font><span style="font-weight: 700"><font color="#FFFFFF" face="微软雅黑">团队成员</font></span></TD>
                    </tr>
					<tr>
                      <TD height="99" bgcolor="#FFFFFF">
                                <TABLE cellSpacing=0 cellPadding=0 width="91%" 
                              border=0 height="43" >
                                  <TBODY>
                                    <TR> 
                                      <TD align="left" height="43" background="bj.jpg">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 0">
										<font color="#004993" face="微软雅黑">　</font><font face="微软雅黑"><img border="0" src="images/about1.jpg" width="19" height="27"><span class="Apple-converted-space">&nbsp;</span><a href="shizi.asp?type=48" style="text-decoration: none">成员简介</a></font></TD>
                                      
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
								<p class="title" style="margin-top: 0; margin-bottom: 0" align="left">
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
							<body>
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
            
              <TD width="835"  vAlign=top style="BORDER-TOP: #C5CED7 1px solid;BORDER-LEFT: #FFFFF 0px solid;BORDER-right: #C5CED7 1px solid;BORDER-bottom: #C5CED7 0px solid;"> 
				<TABLE cellSpacing=0 cellPadding=0 width="832" 
border=0>
                  <TBODY>
                    <TR> 
                      <TD width="832" height="50" align="center" style="BORDER-BOTTOM: #CFCFCF 0px dashed"> 
                        
						<p class="title" align="left">
						<SPAN style="FONT-WEIGHT: bold;"> <img border="0" src="fl/jiantou.jpg" width="17" height="14">
						</font></b> 
                        
                      	<font face="微软雅黑"> 
                        
                      <a href="index.asp" style="text-decoration: none">
						<font color="#004993">首页</font></a><font color="#004993"> 
						                        
                      <b>&gt;&gt; 
						<%=bigclass%><%if smallclass<>"" then%><span style="text-decoration: none"> </span>
						</b>
						</font>
						</font>
						</span> 
						                        
                      	<font color="#004993" face="微软雅黑"> 
						                        
                      	<b><span style="text-decoration: none; font-weight:bold">&gt;&gt; <%=smallclass%><%end if%></span></b></font></TD>
                    </TR>

                  	<tr>
                      <TD height="4" align="center" background="images/lanmufeng.gif" width=832></TD>
                    </tr>
                    <TR> 
                      <TD width="832" align="center" style="BORDER-BOTTOM: #CFCFCF 0px dashed"> 
                        
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
<%if xsfs="图文列表" then%>
                <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <% 
page=clng(request("page"))		 
Set rs=Server.CreateObject("ADODB.RecordSet") 
if bigclass<>"" and smallclass<>"" then
sql="select * from product where ProductBigClass='"&bigclass&"' and ProductSmallClass='"&smallclass&"' order by ProductPrice6 desc"
rs.Open sql,conn,1,1
elseif smallclass="" then
sql="select * from product where ProductBigClass='"&bigclass&"' order by ProductPrice6 desc"
rs.Open sql,conn,1,1
end if
if rs.eof and rs.bof then
response.Write("暂时没有内容")
else 
%>
                  <% 
rs.PageSize=5
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.AbsolutePage=page  
for j=1 to rs.PageSize
%>
                  <tr> 
                    <td  class=lst01 height="20" colspan="2"><table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td height="5"></td>
                        </tr>
                      </table>
                      <TABLE cellSpacing=0 cellPadding=0 width=95% 
                        align=center>
                        <TBODY>
                          <TR> 
                            <TD width=25% align=middle> <TABLE width="136" border=0 align="center" 
                              cellPadding=2 cellSpacing=0>
                                <TBODY>
                                  <TR> 
                                    <TD width="250"> <span class="bh"> 
                                      <% if RS("ProductImages")="" then %>
                                      </span> <TABLE width="130" height="110" border=1 
                                cellPadding=3 cellSpacing=0 borderColor=#b2b2b2>
                                        <TBODY>
                                          <TR> 
                                            <TD width="120" height="108" bgColor=#ffffff>
											<a href="shiziview.asp?id=<%= RS("id") %>" target="_blank"><img src="images/wupic.jpg" width="120" height="100" border="0"></a></TD>
                                          </TR>
                                        </TBODY>
                                      </TABLE>
                                      <span class="bh"> 
                                      <% else %>
                                      </span> <TABLE width="130" height="110" border=1 
                                cellPadding=3 cellSpacing=0 borderColor=#b2b2b2>
                                        <TBODY>
                                          <TR> 
                                            <TD width="120" height="108" bgColor=#ffffff>
											<a href="shiziview.asp?id=<%= RS("id") %>" target="_blank"><img src="<% Response.Write (""&RS("ProductImages")&"")%>" width="120" height="100" border="0"></a></TD>
                                          </TR>
                                        </TBODY>
                                      </TABLE>
                                      <span class="bh"> 
                                      <% end if %>
                                      </span></TD>
                                  </TR>
                                </TBODY>
                              </TABLE>
                              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td>　</td>
                                </tr>
                              </table></TD>
                            <TD width=75% valign="top" style="PADDING-LEFT: 4px"> 
                              <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td height="5"></td>
                                </tr>
                              </table>
                              <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                              align=center border=0>
                                <TBODY>
                                  <TR> 
                                    <TD bgColor=#dfff80> <TABLE width="100%" 
                                border=0 cellPadding=0 cellSpacing=1 bgcolor="f1f1f1">
                                        <TBODY>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD width="17%" height=21> 　</TD>
                                            <TD width="83%">　</TD>
                                          </TR>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD height=21> 　</TD>
                                            <TD>　</TD>
                                          </TR>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD height=21> 　</TD>
                                            <TD>　</TD>
                                          </TR>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD height=21> 　</TD>
                                            <TD>　</TD>
                                          </TR>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD height=21> 　</TD>
                                            <TD bgcolor="#FFFFFF">　</TD>
                                          </TR>
                                        </TBODY>
                                      </TABLE></TD>
                                  </TR>
                                </TBODY>
                              </TABLE></TD>
                          </TR>
                        </TBODY>
                      </TABLE></td>
                  </tr>
                  <%
RS.MoveNEXT
if rs.eof then exit for
next
%>
                  <tr valign="bottom"> 
                    <td height="50" colspan="3" align="center" > 
                      <%if Page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=shizi.asp?type="&bigclassid&"&type2="&smallclassid&"&page=1><font color=red>首页</font></a>&nbsp;"
    response.write "<a href=shizi.asp?type="&bigclassid&"&type2="&smallclass&"&page=" & Page-1 & "><font color=red>上一页</font></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=shizi.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & (page+1) & ">"
    response.write "<font color=red>下一页</font></a> <a href=shizi.asp?type="&bigclassid&"&type2="&smallclassid&"&page="&rs.pagecount&"><font color=red>尾页</font></a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#FF0000'>"&rs.recordcount&"</font></b>条 <b>"&rs.pagesize&"</b>条/页"
%>
                    </td>
                  </tr>
                  <% 
end if
rs.close
set rs=nothing
%>
                </table>
				<% else %>
                <table width="92%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <% 
page=clng(request("page"))		 
Set rs=Server.CreateObject("ADODB.RecordSet") 
if bigclass<>"" and smallclass<>"" then
sql="select * from product where ProductBigClass='"&bigclass&"' and ProductSmallClass='"&smallclass&"' order by ProductPrice6 asc"
rs.Open sql,conn,1,1
elseif smallclass="" then
sql="select * from product where ProductBigClass='"&bigclass&"' order by ProductPrice6 asc"
rs.Open sql,conn,1,1
end if
if rs.eof and rs.bof then
response.Write("暂时没有内容")
else 
%>
                  <% 
rs.PageSize=10
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.AbsolutePage=page  
%>
                  <tr> 
                    <td  class=lst01 height="20" colspan="2"><table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td height="5"></td>
                        </tr>
                      </table>
                      <div align="left">
                      <TABLE cellSpacing=0 cellPadding=0 width=823>
                        <TBODY>
                          <TR> 
                            <TD width=823 align=middle>
<table border="0" width="828">
                                <% 
							   for j=1 to 10 
							   if NOT RS.EOF then
							   %>
                                <tr>
                                  <% 
							   for i=1 to 1
							   if NOT RS.EOF then
							   %> 
                                  <td width="822">
<div align="center">
<TABLE width="750" border=0 
                              cellPadding=2 cellSpacing=0>
                                      <TBODY>
                                        <TR> 
                                          <TD width="752"> <span class="bh"> 
                                            <% if trim(RS("ProductImages"))="" then %>
                                            </span> 
											<div align="center">
											<TABLE width="815" height="177" border=0 
                                cellPadding=3 cellSpacing=0 borderColor=#b2b2b2>
                                              <TBODY>
                                                <TR> 
                                                  <TD width="235" height="88" bgColor=#ffffff>
													<a href="shiziview.asp?id=<%= RS("id") %>">
													<img src="images/wupic.jpg" width="359" height="282" border="0" style="border-style: solid; border-width: 0px; "></a></TD>
                                                  <TD width="568" height="54" bgColor=#ffffff>
														<STRONG>
														<p class="title" style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="left">
														</STRONG>	
														<a style="text-decoration: none" href="shiziview.asp?id=<%= RS("id") %>">												
														<font color="#2B2B2B">
														<STRONG><% =left(RS("ProductName"),20) %>
                                            			</STRONG>														
                                            			</font>
                                            			</p>														
														<p class="shuye2" style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="left">
														<font color="#2B2B2B">
														<b></b><% =left(RS("ProductPrice5"),250) %></font></p>
														<p class="shuye2" style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="left">
														<font color="#2B2B2B">
														<b></font></a>
</p>

														</TD>
													</TR>

                                                <TR> 
                                                  <TD width="803" height="20" bgColor=#ffffff colspan="2">
													<STRONG style="font-weight: 400">
														<p class="shuye2" style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="left">
														<a style="text-decoration: none" href="shiziview.asp?id=<%= RS("id") %>">
														<font color="#2B2B2B">&nbsp;                                            			</font>
														<font color="#0066CC">
														【查看详细介绍】                                            			</font>
                                            			</a>
                                            			</p></STRONG></TD>
													</TR>

                                                <TR> 
                                                  <TD width="803" height="1" bgColor=#C5CED7 colspan="2" bordercolor="#C5CED7"></TD>
													</TR>

                                                </TR>
                                              </TBODY>
                                            </TABLE>
                                            </div>
                                            <span class="bh"> 
                                            <% else %>
                                            </span> 
											<TABLE width="807" height="110" border=0 
                                cellPadding=3 cellSpacing=0 borderColor=#b2b2b2>
												<TBODY>
													<TR>
														<TD width="235" height="54" bgColor=#ffffff>
														<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
														<a href="shiziview.asp?id=<%= RS("id") %>">
														<img src="<% Response.Write (""&RS("ProductImages")&"")%>" width="359" height="282" border="0" style="border: 0px solid #006699; "></a></TD>
														<TD width="560" height="54" bgColor=#ffffff>
														<STRONG>
														<p class="title" style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="left">
														</STRONG>														
														<a style="text-decoration: none" href="shiziview.asp?id=<%= RS("id") %>">
														<font color="#2B2B2B">
														<STRONG><% =left(RS("ProductName"),20) %>
                                            			</STRONG>														
                                            			</font>
                                            			</a>
                                            			</p>														
														<p class="shuye2" style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="left">
														<font color="#2B2B2B">
														<b></b><% =left(RS("ProductPrice5"),250) %></font></p>

														</TD>
													</TR>
													<TR>
														<TD width="795" height="40" bgColor=#ffffff colspan="2">
														<p align="left">
														<STRONG style="font-weight: 400">
														<font color="#2B2B2B">&nbsp;&nbsp;
														</font>
														<a style="text-decoration: none" href="shiziview.asp?id=<%= RS("id") %>">
														<font color="#0066CC">【查看详细介绍】                                            			</font>
                                            			</a></STRONG></TD>
													</TR>
													
                                                <TR> 
                                                  <TD width="803" height="1" bgColor=#C5CED7 colspan="2" bordercolor="#C5CED7"></TD>
													</TR>

												</TBODY>
											</TABLE>
                                            <span class="bh"> 
                                            <% end if %>
                                            </span></TD>
                                        </TR>
                                      </TBODY>
                                    </TABLE>
                                  </div>
                                  </td>
 <% RS.MoveNEXT
end if
next %>
                                </tr>
								 <% 
end if
next %>
                              </table> 
                              
                            </TD>
                          </TR>
                        </TBODY>
                      </TABLE></div>
					</td>
                  </tr>
                  <tr valign="bottom"> 
                    <td height="50" colspan="3" align="center" > 
                      <font color="#676767"> 
                      <%if Page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=shizi.asp?type="&bigclassid&"&type2="&smallclassid&"&page=1><font color=red>首页</font></a>&nbsp;"
    response.write "<a href=shizi.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & Page-1 & "><font color=red>上一页</font></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=shizi.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & (page+1) & ">"
    response.write "<font color=red>下一页</font></a> <a href=shizi.asp?type="&bigclassid&"&type2="&smallclassid&"&page="&rs.pagecount&"><font color=red>尾页</font></a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#FF0000'>"&rs.recordcount&"</font></b>条 <b>"&rs.pagesize&"</b>条/页"
%>
                    </font>
                    </td>
                  </tr>
                  <% 
end if
rs.close
set rs=nothing
%>
                </table>
                <TABLE cellSpacing=0 cellPadding=0 width=100 align=center bgColor=#ffffff 
border=0>
                  <TBODY>
                    <TR> 
                      <TD height=8></TD>
                    </TR>
                  </TBODY>
                </TABLE><% end if %></TD>
            </TR>
          </TBODY>
        </TABLE></div>
		</TD>
     
    </TR>
  </TBODY>
</TABLE>
</div>
<!--#include file="foot.asp"-->
</body>
</html>