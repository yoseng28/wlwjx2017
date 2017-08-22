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
<title><% =sitetitle %>--<% =bigclass %></title>
<meta name="keywords" content="<% =bigclass %>,<% =smallclass %>,<% =webKeywords %>">
<meta name="description" content="<% =webDescription %>,<% =bigclass %>,<% =smallclass %>">
<meta name="robots" content="index, follow">
<meta name="googlebot" content="index, follow">
<style fprolloverstyle>A:hover {color: red; font-weight: bold}
</style>
</head>
<BODY topMargin=0 leftmargin="0" link="#515753" vlink="#515753" alink="#515753">
<!--#include file="top.asp"-->
<TABLE cellSpacing=0 cellPadding=0 width=1100 align=center bgColor=#ffffff 
border=0>
  <TBODY>
    <TR> 
      <TD width="5" bgcolor="#FFFFFF"></TD>
      <TD width="1100"><TABLE width=1100 border=0 align=center cellPadding=0 cellSpacing=0  vAlign=top style="BORDER-TOP: #808080 1px solid;BORDER-LEFT: #808080 2px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 1px solid;">
          <TBODY>
            <TR> 
              <TD width="265" vAlign=top background="../../../../../../../../images/class_bg4.gif" style="background-repeat: repeat-x"
 bgcolor="#FFFFFF"> 
                
                <table border="0" width="240" cellspacing="0" height="436">
					<tr>
                                      <TD align="left" height="20" background="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/bj.jpg">
										<p style="margin-top: 0; margin-bottom: 0">
										<font color="#006699" face="微软雅黑">&nbsp;</font><img border="0" src="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../fl/jituan.gif" width="258" height="71"></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="20" background="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/bj.jpg">
										<font color="#006699" face="微软雅黑">
										<img border="0" src="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/about1.jpg" width="19" height="27"><font color="#CDCED0">
										<span style="font-weight: 700">
										<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=31">
										<font color="#4D4D4D">电子产品</font></a></span></font></font></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="60">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑" color="#4D4D4D">　　　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=31&type2=30">电脑整机</a></font><font face="微软雅黑">
										　　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=31&type2=31">电脑配件</a></font></p>
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑">　　　</font><font face="微软雅黑" color="#4D4D4D"><a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=31&type2=39">家用电器</a></font><font face="微软雅黑">
										　　</font><font face="微软雅黑" color="#4D4D4D"><a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=31&type2=45">手机数码</a></font></p>
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑">　　　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=31&type2=46">网络产品</a></font></p></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="40" background="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/bj.jpg">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font color="#006699" face="微软雅黑">&nbsp;<img border="0" src="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/about1.jpg" width="19" height="27"></font><font color="#CDCED0" face="微软雅黑">
										<b>
										<font color="#CDCED0">
										<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=60">
										<font color="#4D4D4D">体育用品</font></a></font></b></font></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="60">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										　<font color="#4D4D4D" face="微软雅黑">　　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=60&type2=47">篮球排球</a></font><font face="微软雅黑">
										　　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=60&type2=63">乒乓球排</a></font></p>
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑">　　　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=60&type2=64">羽毛球排</a>
										　　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=60&type2=48">足球网球</a></font><p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑">　　　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=60&type2=65">旅游用品</a>
										　　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=60&type2=66">其　　它</a></font></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="40" background="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/bj.jpg">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font color="#006699" face="微软雅黑">&nbsp;<img border="0" src="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/about1.jpg" width="19" height="27"></font><font color="#CDCED0" face="微软雅黑">
										</font>
										<b>
										<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=61"><font face="微软雅黑" color="#4D4D4D">
										健身用品</font></a></b></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="60">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑">　<font color="#4D4D4D">　</font><font color="#669900">　</font><font color="#4D4D4D"><a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=61&type2=50">跑步机</a>　
										　　 
										 
										<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=61&type2=50">踏步机</a>
										<a href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=61&type2=51" style="text-decoration: none">&nbsp;　</a></font></font></p>
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑" color="#4D4D4D">　　　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=61&type2=51">综合训练器</a> 
										　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=61&type2=52">运动护具</a></font><p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑" color="#4D4D4D">　　　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=61&type2=53">武术搏击</a> 
										　&nbsp; </font><font face="微软雅黑">&nbsp;<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=61&type2=67">中小型器材</a></font></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="40" background="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/bj.jpg">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font color="#006699" face="微软雅黑">&nbsp;<img border="0" src="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/about1.jpg" width="19" height="27"></font><font color="#CDCED0" face="微软雅黑">
										</font>
										<b>
										<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=62"><font face="微软雅黑" color="#4D4D4D">
										生活用品</font></a></b></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="40">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										　<font color="#4D4D4D" face="微软雅黑">　　<a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=62&type2=54">书桌课桌</a> 
										　　</font><p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑" color="#4D4D4D">　　　</font><font face="微软雅黑"><a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=62&type2=55">其　　它</a></font></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="40" background="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/bj.jpg">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑" color="#006699">&nbsp;<img border="0" src="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/about1.jpg" width="19" height="27"></font><font color="#CDCED0" face="微软雅黑"><b>
										</b></font>
										<a href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=63" style="text-decoration: none">
										<b><font face="微软雅黑" color="#4D4D4D">
										招聘求职</font></b></a></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="30">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑" color="#4D4D4D">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 　<a href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=63&type2=57" style="text-decoration: none">专职招聘</a>　　</font><a href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=63&type2=60" style="text-decoration: none"><font face="微软雅黑">兼职招聘</font></a></p>
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑" color="#4D4D4D">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 　<a href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=63&type2=59" style="text-decoration: none">专职求职</a>　　</font><a href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=63&type2=58" style="text-decoration: none"><font face="微软雅黑">兼职求职</font></a></p>
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑" color="#4D4D4D">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
										　<a href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=63&type2=61" style="text-decoration: none">寒假暑假工</a>　<a href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=63&type2=61" style="text-decoration: none">其　它</a></font></p></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="40" background="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/bj.jpg">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑" color="#006699">&nbsp;<img border="0" src="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/about1.jpg" width="19" height="27"></font><font color="#CDCED0" face="微软雅黑"><b>
										<a href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=70" style="text-decoration: none">
										<font color="#4D4D4D">旅游信息</font></a></b></font></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="15">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑" color="#4D4D4D">&nbsp;&nbsp;&nbsp;&nbsp; 
										　 
										<a style="text-decoration: none" href="lynews.asp?type=68&type2=69">国内旅游</a></font></p>
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑" color="#4D4D4D">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 　<a href="lynews.asp?type=68&type2=70" style="text-decoration: none">国外旅游</a></font></p>
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑" color="#4D4D4D">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 　<a style="text-decoration: none" href="lynews.asp?type=68&type2=71">周边当地旅游</a><a style="text-decoration: none" href="product.asp?type=70&type2=68"> 
										</a>&nbsp;&nbsp;&nbsp;</font></p></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="40" background="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/bj.jpg">
										<font face="微软雅黑" color="#006699">
										<img border="0" src="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../images/about1.jpg" width="19" height="27"></font><b><font color="#CDCED0" face="微软雅黑"> </font>
										<a href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../product.asp?type=70" style="text-decoration: none">
										<font color="#4D4D4D" face="微软雅黑">培训班</font></a></b></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="50">
										<font face="微软雅黑" color="#4D4D4D">&nbsp;&nbsp;&nbsp;&nbsp; 
										　</font><font face="微软雅黑"><a style="text-decoration: none" href="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../xxnews.asp?type=66">培训班信息</a></font></TD>
                                      
                                    </tr>
					<tr>
						<td bgcolor="#FFFFFF"> 
              	　</td>
					</tr>
					<tr>
						<td bgcolor="#FFFFFF" width="184" height="170" align="left"> 
						<font color="#666666">
						<table border="0" width="101%" cellspacing="0">
							<tr>
								<td> 
                <table border="0" width="252" cellspacing="0" height="253">
					<tr>
						<td bgcolor="#FFFFFF" colspan="2"> 
              	<img border="0" src="../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../../fl/lianxi.gif" width="251" height="65" align="left"></td>
					</tr>

					<tr>
						<td bgcolor="#FFFFFF" width="71" height="30" align="left"> 
              	<p style="margin-left: 5px; margin-top: 0; margin-bottom: 0">
				<font color="#666666">地　址：</font></td>
						<td bgcolor="#FFFFFF" width="177" height="50" align="left"> 
						<font color="#666666"><% =address %>              	　</font></td>
					</tr>

					<tr>
						<td bgcolor="#FFFFFF" width="71" height="25" align="left"> 
              	<p style="margin-left: 5px; margin-top: 0; margin-bottom: 0">
				<font color="#666666">联系人：</font></td>
						<td bgcolor="#FFFFFF" width="177" height="25" align="left"> 
						<font color="#666666"><% =companycon %>
              			　</font></td>
					</tr>

					<tr>
						<td bgcolor="#FFFFFF" width="71" height="25" align="left"> 
              	<p style="margin-left: 5px; margin-top: 0; margin-bottom: 0">
				<font color="#666666">电　话：</font></td>
						<td bgcolor="#FFFFFF" width="177" height="25" align="left"> 
						<font color="#666666"><% =phone %></font></td>
					</tr>

					<tr>
						<td bgcolor="#FFFFFF" width="71" height="25" align="left"> 
              	<p style="margin-left: 5px; margin-top: 0; margin-bottom: 0">
				<font color="#666666">邮　编：</font></td>
						<td bgcolor="#FFFFFF" width="177" height="25" align="left"> 
              	<font color="#666666"><% =code %></font></td>
					</tr>

					<tr>
						<td bgcolor="#FFFFFF" width="71" height="25" align="left"> 
              	<p style="margin-left: 5px; margin-top: 0; margin-bottom: 0">
				<font color="#666666">手　机：</font></td>
						<td bgcolor="#FFFFFF" width="177" height="25" align="left"> 
              	<font color="#666666"><% =beianhao %></font></td>
					</tr>

					</table>
              					</td>
							</tr>
						</table></td>
					</tr>
				</table>　<p> 
                
                </TD>
              <TD width="8">　</TD>
              <TD width="823" vAlign=top style="BORDER-TOP: #808080 1px solid;BORDER-LEFT: #808080 1px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 1px solid;"> 
                <TABLE cellSpacing=0 cellPadding=0 width="819" 
border=0 height="76">
                  <TBODY>
                    <TR> 
                      <TD width="694" align="right" height="71">
						<p align="left"><b><font color="#205081">
						<img border="0" src="../../../../../../../../fl/jiantou.jpg" width="17" height="14"></font></b><font color="#663300" face="微软雅黑"> </font>
						<font face="微软雅黑"><font color="#663300">&nbsp;</font><a href="../../../../../../../../index.asp" style="text-decoration: none"><font color="#663300">首页</font></a><font color="#663300"> 
                        &gt;&gt; </font> 
						<a href="fjzs.asp?type=<%=bigclassid%>" style="text-decoration: none">
						<font color="#663300"><%=bigclass%></font></a><font color="#663300"> 
                        &gt;&gt; <%=smallclass%> 
                        </font> 
                        </font> 
                        </TD>
                      <TD width="125">　</TD>
                    </TR>
                    <TR> 
                      <TD height="5" colspan="2" align="center" background="../../../../../../../../images/lanmufeng.gif"></TD>
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
                </TABLE>
                <table width="710" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td valign="top"> <DIV align=center> 
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
                        <TABLE width="820" border=0 cellpadding="0" cellSpacing=0 id=table5>
                          <TBODY>
                            <% 
rs.PageSize=20
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.AbsolutePage=page  
for j=1 to rs.PageSize 
%>
                            <TR height=24> 
                              <TD height="35" 
                        align=center width="112" bgcolor="#CC0000"> 
                                <b> 
                                <font face="微软雅黑" color="#FFFFFF"> 
                               职位</font></b></TD>
                              <TD height="35" 
                        align=center width="558" bgcolor="#CC0000"> 
								<font color="#FFFFFF" face="微软雅黑"><b>招聘单位
                                </b></font>
                                </TD>
                              <TD height="35" 
                        align=center width="129" bgcolor="#CC0000"> 
								<font color="#FFFFFF" face="微软雅黑"><b>发布时间
                               </b></font>
                               </TD>
                            </TR>
                            <TR height=24> 
                              <TD height="35" 
                        align=left width="112"> <font face="微软雅黑"> 
                                <% if rs("firstImageName")<>"" then response.write "<img src='images/news.gif' border=0 alt='图片'>" end if %>
                                <a href="view.asp?id=<% =rs("id") %>" target="_blank"> 
                                <% =rs("title1") %>
                                </a></font>                 
                                </TD>
                              <TD height="35" 
                        align=left width="558">  <font face="微软雅黑"> 
                                <% if rs("firstImageName")<>"" then response.write "<img src='images/news.gif' border=0 alt='图片'>" end if %>
                                <a href="view.asp?id=<% =rs("id") %>" target="_blank"> 
                                <% =rs("title") %>
                                </a></font>                                </TD>
                              <TD height="35" 
                        align=left width="129">  <font face="微软雅黑" color="#663300">( 
                                <% =year(rs("infotime")) %>
                                - 
                                <% =month(rs("infotime")) %>
                                - 
                                <% =day(rs("infotime")) %>
                                )</font><font face="微软雅黑"> </font>
                                </TD>
                            </TR>
                            <TR> 
                                    <TD height="2" colspan="3" vAlign=middle class=listb background="../../../../../../../../images/newx_fenge.gif"g></TD>
                                  </TR>
                            <%
rs.movenext
if rs.eof then exit for
next
%>
                          </TBODY>
                        </TABLE>
                        <CENTER>
                          <table width="100%" border="0">
                            <tr> 
                              <td height="40"><div align="center"> 
                                  <font color="#663300"> 
                                  <%if Page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=list.asp?type="&bigclassid&"&type2="&smallclassid&"&page=1><FONT color=#ff0000>首页</FONT></a>&nbsp;"
    response.write "<a href=list.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & Page-1 & "><FONT color=#ff0000>上一页</FONT></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=list.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & (page+1) & ">"
    response.write "<FONT color=#ff0000>下一页</FONT></a> <a href=list.asp?type="&bigclassid&"&type2="&smallclassid&"&page="&rs.pagecount&"><FONT color=#ff0000>尾页</FONT></a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#FF0000'>"&rs.recordcount&"</font></b>条信息 <b>"&rs.pagesize&"</b>条信息/页"
%>
                                </font>
                                </div></td>
                            </tr>
                          </table>
                        </CENTER>
                      </DIV>
                      <% 
end if
rs.close
set rs=nothing
%>
                      <TABLE cellSpacing=0 cellPadding=0 width=100 align=center bgColor=#ffffff 
border=0>
                        <TBODY>
                          <TR> 
                            <TD height=8></TD>
                          </TR>
                        </TBODY>
                      </TABLE></td>
                  </tr>
                </table>
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