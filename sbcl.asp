<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">

	<!--#include file="conn.asp"-->
	<!--#include file="inc/function.asp"-->
	<% 
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
 %>
	<html>

	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<link rel="shortcut icon" href="images/new/wlw.ico" >
		<title>
			<% =bigclass %>--
			<% =sitetitle %>
		</title>
		<meta name="keywords" content="<% =bigclass %>,<% =webKeywords %>">
		<meta name="description" content="<% =webDescription %>,<% =bigclass %>">
		<meta name="robots" content="index, follow">
		<meta name="googlebot" content="index, follow">
		<style fprolloverstyle>
			A:hover {
				color: #004993;
				font-weight: bold
			}
		</style>
		<style>
			<!-- .xuesheng {
				font-size: 12pt;
				font-weight: bold
			}
			
			.title {
				font-size: 14pt
			}
			
			.shuye {
				font-size: 11pt
			}
			
			-->
		</style>
	</head>

	<BODY topMargin=0 leftmargin="0" link="#5E5E5E" vlink="#5E5E5E" alink="#5E5E5E">
		<!--#include file="top.asp"-->
		<div align="center">
			<TABLE cellSpacing=0 cellPadding=0 width=1100 bgColor=#ffffff align="center" border=0>
				<TBODY>
					<TR>
						<TD width="1100">
							<TABLE width=1100 border=0 align=center cellPadding=0 cellSpacing=0 style="BORDER-TOP: #808080 1px solid;BORDER-LEFT: #808080 1px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 1px solid;">
								<TBODY>
									<TR>
										<TD width="260" vAlign=top style="background-repeat: repeat-x" bgcolor="#FFFFFF">
											<table border="0" width="260" cellspacing="0" height="436" style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 0px solid;BORDER-right: #808080 0px solid;BORDER-bottom: #808080 0px solid;">
												<tr>
													<TD height="49" align="center" bgcolor="#0099FF">
														<p align="left" class="title" style="margin-top: 0; margin-bottom: 0">
															<font color="#006699" face="微软雅黑">　</font><span style="font-weight: 700"><font color="#FFFFFF" face="微软雅黑">任务材料</font></span></TD>
												</tr>
												<tr>
													<TD height="99" bgcolor="#FFFFFF">
														<TABLE cellSpacing=0 cellPadding=0 width="85%" border=0 height="42">
															<TBODY>
																<TR>
																	<TD align="left" height="42" background="bj.jpg">
																		<p class="xuesheng" style="margin-top: 0; margin-bottom: 0; line-height:200%">
																			<font color="#004993" face="微软雅黑">　</font>
																			<font face="微软雅黑"><img border="0" src="images/about1.jpg" width="19" height="27"><span class="Apple-converted-space">&nbsp;&nbsp;</span>
																				<a style="text-decoration: none" href="sbcl.asp?type=72">任务书</a>
																			</font>
																	</TD>
																</TR>
															</TBODY>
														</TABLE>
													</TD>
												</tr>
												<tr>
													<td bgcolor="#FFFFFF" height="260">
														<font color="#666666">
													</td>
												</tr>
												<tr>
													<td bgcolor="#FFFFFF">
														</td>
												</tr>
											</table>
											</TD>
										<TD width="1"></TD>
										<TD width="850" vAlign=top style="BORDER-TOP: #D9DDD7 1px solid;BORDER-LEFT: #D9DDD7 1px solid;BORDER-right: #D9DDD7 1px solid;BORDER-bottom: #D9DDD7 1px solid;">
											<TABLE cellSpacing=0 cellPadding=0 width="850" border=0 height="50">
												<TBODY>
													<TR>
														<TD width="10" height="55" align="center" background="fl/cp1.gif">　</TD>
														<TD width="404" align="right" height="55" background="fl/cp1.gif">
															<p align="left" style="margin-top: 0; margin-bottom: 0" class="shuye">
																<b>
						<font face="微软雅黑" color="#5E5E5E">
						<p class="title" style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="left"></font><font face="微软雅黑"><font color="#004993">
							<img border="0" src="fl/jiantou.jpg" width="17" height="14"><a href="index.asp"style="color: #004993;text-decoration: none">首页</a> &gt;&gt; <%=bigclass%> 
						</font> </p></font> 
                      	</b>
														</TD>
														<TD width="512" background="fl/cp1.gif">　</TD>
													</TR>
													<TR>
														<TD height="1" align="center"></TD>
														<TD height="1" colspan="2" align="center" background="images/lanmufeng.gif"></TD>
													</TR>
												</TBODY>
											</TABLE>
											<TABLE cellSpacing=0 cellPadding=0 width=100 align=center border=0>
												<TBODY>
													<TR>
														<TD height=8></TD>
													</TR>
												</TBODY>
											</TABLE>
											<table width="850" border="0" align="center" cellpadding="0" cellspacing="0">
												<tr>
													<td width="850">
														<%
set rsSmallClass=server.CreateObject("adodb.recordset")
rsSmallClass.open "Select * From SmallClass Where BigClassName='" & bigclass & "'",conn,1,1
if not(rsSmallClass.bof and rsSmallClass.eof) then
do while not rsSmallClass.eof
Small=rsSmallClass("SmallClassName")
%>
														<div align="left"></div>
														<TABLE width="847" border=0 cellPadding=0 cellSpacing=0 style="BORDER-bottom: #FF9900 1px solid;">
															<TBODY>
																<TR>
																	<TD width="4%" height="30">
																		<div align="center">
																			<img src="images/smallclasstb.gif" width="20" height="20"></div>
																	</TD>
																	<TD width="38%">
																		<B><a  href="list.asp?type=<%=bigclass%>&type2=<%=rsSmallClass("SmallClassName")%>"><SPAN style="FONT-SIZE: 14px"><font color="205081"><%=rsSmallClass("SmallClassName")%></font></SPAN></a></B>
																	</TD>
																	<TD width="51%" align=middle>　</TD>
																	<TD width="7%">
																		<a href="list.asp?type=<%=bigclassid%>&type2=<%=rsSmallClass(" SmallClassid ")%>">
																			<IMG height=11 src="images/cw_more.gif" width=30 border=0>
																		</a>
																	</TD>
																</TR>
															</TBODY>
														</TABLE>
														<TABLE style="WORD-BREAK: break-all" cellSpacing=0 cellPadding=0 width="810" border=0>
															<TBODY>
																<TR>
																	<TD height=25 rowSpan=2 align=middle vAlign=top> <span class="bh"> 
                              <%
sql2="select *  from content where BigClassName='"&bigclass&"' and SmallClassName='"&Small&"' order by id desc "
Set oRS= Server.CreateObject("ADODB.recordset")
oRS.Open sql2,conn,1,3
 if NOT oRS.EOF then 
for i=1 to 8
  if NOT oRS.EOF then
   %>
                              </span>
			<TABLE width="810" border="0" align="center" cellPadding=0 cellSpacing=0>
				<TBODY>
					<TR>
						<TD width=20 height="9" vAlign=middle class=listbg align="left">
						<p align="center">
						<img border="0" src="fl/icon03.gif" width="4" height="4"></TD>
						<TD width="679" height="30" class=listbg align="left">
							<a href="newsview.asp?id=<% =ors("id") %>">
						<font color="#404A53">
							<% =ors("title") %>
						</font></a>
					    </TD>
						<TD width="109" height="30" class=listbg align="left">
							<font color="#404A53">
								(<% =year(ors("infotime")) %>-
								<% =month(ors("infotime")) %>-
								<% =day(ors("infotime")) %>)
							</font>
						</TD>
					</TR>
					<TR>
						<TD height="2" colspan="3" vAlign=middle class=listb background="images/newx_fenge.gif" g></TD>
					</TR>
				</TBODY>
			</TABLE>
			<span class="bh">  
<%
oRS.MoveNEXT
end if
next
else Response.Write ("暂无内容") end if
oRS.close 
Set oRS = Nothing
%>
                              </span>
																		<table width="200" height=5 border=0 align="center" cellpadding=0 cellspacing=0>
																			<tbody>
																				<tr height=5>
																					<td></td>
																				</tr>
																			</tbody>
																		</table>
																		<span class="bh"> </span></TD>
																</TR>
															</TBODY>
														</TABLE>
														<%
			rsSmallClass.movenext
		loop
		else %>
														<DIV align=center>
															<% 
page=clng(request("page"))		 
Set rs=Server.CreateObject("ADODB.RecordSet") 
if bigclass<>"" and smallclass <>"" then
sql="select * from content where BigClassName='"&bigclass&"' and SmallClassName='"&smallclass&"' order by id desc"
rs.Open sql,conn,1,1
elseif bigclass<>"" then
sql="select * from content where BigClassName='"&bigclass&"' order by id desc"
rs.Open sql,conn,1,1
end if
if rs.eof and rs.bof then
response.Write("暂无内容")
else 
%>
															<div align="left">
																<TABLE width="810" border=0 cellpadding="0" cellSpacing=0 id=table5>
																	<TBODY>
																		<% 
rs.PageSize=35
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.AbsolutePage=page  
for j=1 to rs.PageSize 
%>
		<TR height=24>
		<TD width="20" height="25" align=left class="shuye">
			<div align="center"><b><font color="#404A53">
			<img border="0" src="fl/icon03.gif" width="4" height="4"> 
                                </font> 
                                </b>
			</div>
		</TD>
		<TD height="30" align=left width="682" class="shuye">
		<font face="微软雅黑" color="#404A53">
		<% if rs("firstImageName")<>"" then response.write "<img src='images/news.gif' border=0 alt='图片'>" end if %>
		</font>
		<a href="newsview.asp?id=<% =rs("id") %>" style="text-decoration: none">
		<font face="微软雅黑" color="#404A53">
		<%=left(rs("title"),38)%>
		</font>
																				</a>
																			</TD>
		<TD height="30" align=left width="94" class="shuye">
		<font color="#404A53" face="微软雅黑">
			(<% =year(rs("infotime")) %>.
			<% =month(rs("infotime")) %>.
			<% =day(rs("infotime")) %> )
		</font>
		</TD>
	</TR>
	<TR>
		<TD height="2" colspan="3" vAlign=middle class=listb background="../images/newx_fenge.gif" g></TD>
	</TR>
<%
rs.movenext
if rs.eof then exit for
next
%>
																	</TBODY>
																</TABLE>
															</div>
															<CENTER>
																<table width="710" border="0">
																	<tr>
																		<td height="40" class="shuye">
																			<div align="center">
																				<font color="#515753" face="微软雅黑">
																					<%if Page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=sbcl.asp?type="&bigclassid&"&type2="&smallclassid&"&page=1><FONT color=#ff0000>首页</FONT></a>&nbsp;"
    response.write "<a href=sbcl.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & Page-1 & "><FONT color=#ff0000>上一页</FONT></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=sbcl.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & (page+1) & ">"
    response.write "<FONT color=#ff0000>下一页</FONT></a> <a href=sbcl.asp?type="&bigclassid&"&type2="&smallclassid&"&page="&rs.pagecount&"><FONT color=#ff0000>尾页</FONT></a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#FF0000'>"&rs.recordcount&"</font></b>条信息 <b>"&rs.pagesize&"</b>条信息/页"
%> </font>
																			</div>
																		</td>
																	</tr>
																</table>
															</CENTER>
														</DIV>
														<% 
end if
rs.close
set rs=nothing
%>
														<%
	  end if
	  rsSmallClass.close
	  set rsSmallClass=nothing	
%>
													</td>
												</tr>
											</table>
											<TABLE cellSpacing=0 cellPadding=0 width=100 align=center bgColor=#ffffff border=0>
												<TBODY>
													<TR>
														<TD height=8></TD>
													</TR>
												</TBODY>
											</TABLE>
											<TABLE cellSpacing=0 cellPadding=0 width=100 align=center border=0>
												<TBODY>
													<TR>
														<TD height=8></TD>
													</TR>
												</TBODY>
											</TABLE>
										</TD>
									</TR>
									</TBODY>
							</TABLE>
						</TD>
						<TD width="45"></TD>
					</TR>
					</TBODY>
			</TABLE>
		</div>
		<!--#include file="foot.asp"-->
		</body>

	</html>