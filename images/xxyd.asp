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
<title><% =bigclass %>--<% =sitetitle %></title>
<meta name="keywords" content="<% =bigclass %>,<% =webKeywords %>">
<meta name="description" content="<% =webDescription %>,<% =bigclass %>">
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
<BODY topMargin=0 leftmargin="0" link="#5E5E5E" vlink="#5E5E5E" alink="#5E5E5E">
<!--#include file="top.asp"-->
<div align="center">
<TABLE cellSpacing=0 cellPadding=0 width=1100 bgColor=#ffffff align="center"
border=0>
  <TBODY>
    <TR> 
      <TD width="1100">
		<TABLE width=1100 border=0 align=center cellPadding=0 cellSpacing=0    style="BORDER-TOP: #808080 1px solid;BORDER-LEFT: #808080 1px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 1px solid;">
          <TBODY>
            <TR> 
              <TD width="260" vAlign=top  style="background-repeat: repeat-x"
 bgcolor="#FFFFFF"> 
                <table border="0" width="260" cellspacing="0" height="436"   style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 0px solid;BORDER-right: #808080 0px solid;BORDER-bottom: #808080 0px solid;">
					<tr>
                      <TD height="49" align="center" bgcolor="#0099FF">
						<p align="left" class="title" style="margin-top: 0; margin-bottom: 0">
										<font color="#006699" face="微软雅黑">　</font><span style="font-weight: 700"><font color="#FFFFFF" face="微软雅黑">学习园地</font></span></TD>
                    </tr>
					<tr>
                      <TD height="99" bgcolor="#FFFFFF">
                                <TABLE cellSpacing=0 cellPadding=0 width="91%" 
                              border=0 height="87" >
                                  <TBODY>
                                    <TR> 
                                      <TD align="left" height="40" background="bj.jpg" style="font-size: 10pt; line-height: 22px; font-family: arial">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 0">
										<font color="#004993" face="微软雅黑">　</font><font face="微软雅黑"><img border="0" src="images/about1.jpg" width="19" height="27"><span class="Apple-converted-space">&nbsp;</span></font><font face="微软雅黑"><a style="text-decoration: none" href="ydlist.asp?type=64&type2=27">基础技能学习资源</a></font></TD>
                                      
                                    </TR>
                                    <TR> 
                                      <TD align="left" height="40" background="bj.jpg" style="font-size: 10pt; line-height: 22px; font-family: arial">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 0">
										<font color="#004993" face="微软雅黑">　</font><font face="微软雅黑"><img border="0" src="images/about1.jpg" width="19" height="27"><span class="Apple-converted-space">&nbsp;</span></font><font face="微软雅黑"><a style="text-decoration: none" href="ydlist.asp?type=64&type2=28">课程资源</a></font></TD>
                                      
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
              <TD width="1"></TD>
                <TD width="850" vAlign=top style="BORDER-TOP: #D9DDD7 1px solid;BORDER-LEFT: #D9DDD7 1px solid;BORDER-right: #D9DDD7 1px solid;BORDER-bottom: #D9DDD7 1px solid;">                
				<TABLE cellSpacing=0 cellPadding=0 width="850" 
border=0 height="50">
                  <TBODY>
                    <TR> 
                      <TD width="10" height="55" align="center" background="fl/cp1.gif">　</TD>
                      <TD width="404" align="right" height="55" background="fl/cp1.gif">
						<p align="left" style="margin-top: 0; margin-bottom: 0" class="shuye">
						<b>
						<font face="微软雅黑" color="#5E5E5E">
						<p class="title" style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="left"></font><font face="微软雅黑"><font color="#004993"><img border="0" src="../fl/jiantou.jpg" width="17" height="14">首页 &gt;&gt; <%=bigclass%> 
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
                <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
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
                      <TABLE width="850" border=0 cellPadding=0 cellSpacing=0 style="BORDER-bottom: #FF9900 1px solid;">
                        <TBODY>
                          <TR> 
                            <TD width="3%" height="40" align="left" class="xuesheng" bgcolor="#E9EAE8"> 
                              <div align="center">
								<font color="#005BB7">◎</font></div></TD>
                            <TD width="39%" align="left" class="xuesheng" bgcolor="#E9EAE8" ><B><SPAN style="FONT-SIZE: 14px">

							<font color="#005BB7" ><p class="xuesheng" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><%=rsSmallClass("SmallClassName")%></p></font></SPAN></B></TD>
                            <TD width="51%" align=left bgcolor="#E9EAE8">　</TD>
                            <TD width="7%" align="left" bgcolor="#E9EAE8"><a href="ydlist.asp?type=<%=bigclassid%>&type2=<%=rsSmallClass("SmallClassid")%>">
							<font color="#005BB7">
							<IMG height=11 
            src="images/cw_more.gif" 
            width=30 border=0></font></a></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                      <TABLE style="WORD-BREAK: break-all" cellSpacing=0 cellPadding=0 
      width="710" border=0>
                        <TBODY>
                          <TR> 
                            <TD height=25 rowSpan=2 align=middle vAlign=top> <span class="bh"> 
                              <%
sql2="select *  from content where BigClassName='"&bigclass&"' and SmallClassName='"&Small&"' order by id desc "
Set oRS= Server.CreateObject("ADODB.recordset")
oRS.Open sql2,conn,1,3
 if NOT oRS.EOF then 
for i=1 to 5
  if NOT oRS.EOF then
   %>
                              </span> 
							<TABLE width="850" border="0" align="center" cellPadding=0 cellSpacing=0>
                                <TBODY>
                                  <TR> 
                                    <TD width=20 height="20" class=shuye align="left">
									<p align="center" style="line-height: 200%; margin-top: 0; margin-bottom: 0">
									<font face="微软雅黑">●</font></TD>
                                    <TD width="616" height="20" class=shuye align="left">
									<p class="shuye" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><a href="tdgcview.asp?id=<% =ors("id") %>" style="text-decoration: none" > 
                                      <font color="#3D3D3D"> 
                                      <% =ors("title") %></p>
                                      </font>
                                      </a>
                                      </TD>
                                    <TD width="211" height="20" class=listbg align="left">
                                      <font color="#515753"> 

									<p class="shuye" style="line-height: 200%; margin: 0 16px; text-indent:32px" align="right"><% =year(ors("infotime")) %>-<% =month(ors("infotime")) %>- <% =day(ors("infotime")) %></TD>
                                  </TR>
                                  <TR> 
                                    <TD height="2" colspan="3" vAlign=middle class=listb background="images/newx_fenge.gif"></TD>
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
                              </span> <table width="200" height=5 border=0 align="center" cellpadding=0 cellspacing=0>
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
                        <TABLE width="850" border=0 cellpadding="0" cellSpacing=0 id=table5>
                          <TBODY>
                            <% 
rs.PageSize=20
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.AbsolutePage=page  
for j=1 to rs.PageSize 
%>
                            <TR> 
                                    <TD width="20" height="25" 
                        align=left class="shuye"><div align="center"><b><font color="#404A53">
								<img border="0" src="fl/icon03.gif" width="4" height="4"> 
                                </font> 
                                </b> 
                                </div></TD>
                              <TD height="30" 
                        align=left width="682" class="shuye"> 
                                <font face="微软雅黑" color="#404A53"> 
                                <% if rs("firstImageName")<>"" then response.write "<img src='images/news.gif' border=0 alt='图片'>" end if %>
                                </font>
                                <a href="tdgcview.asp?id=<% =rs("id") %>" style="text-decoration: none" > 
                                <font face="微软雅黑" color="#404A53"> 
                                 <%=left(rs("title"),38)%>
                                </font>
                                </a></TD>
                              <TD height="30" 
                        align=left width="94" class="shuye"> 
								<font color="#404A53" face="微软雅黑">( 
                                <% =year(rs("infotime")) %>.<% =month(rs("infotime")) %>.<% =day(rs("infotime")) %>
                               )</font></TD>
                                  </TR>
                                  <TR> 
                                    <TD width=761 height="30" class=index1 align="left" colspan="3">
									<p class="index2" style="margin:0 16px; line-height: 200%; text-indent:32px">
                                      <font color="#404A53"> 
                                      <% =ors("title1") %></p>

									</font> 
                                      <font color="#515753"> 

									<p class="index2" style="line-height: 200%; margin: 0 16px; text-indent:32px" align="right">发布时间：<% =year(ors("infotime")) %>-<% =month(ors("infotime")) %>- <% =day(ors("infotime")) %></TD>
                                  </TR>
                                  <TR> 
                                    <TD height="2" colspan="3" vAlign=middle class=listb background="images/newx_fenge.gif"></TD>
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
                          <table width="865" border="0">
                            <tr> 
                              <td height="40"><div align="center"> 
                                  <font color="#515753" face="微软雅黑"> 
                                  <%if Page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=xxyd.asp?type="&bigclassid&"&type2="&smallclassid&"&page=1><FONT color=#ff0000>首页</FONT></a>&nbsp;"
    response.write "<a href=xxyd.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & Page-1 & "><FONT color=#ff0000>上一页</FONT></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=xxyd.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & (page+1) & ">"
    response.write "<FONT color=#ff0000>下一页</FONT></a> <a href=xxyd.asp?type="&bigclassid&"&type2="&smallclassid&"&page="&rs.pagecount&"><FONT color=#ff0000>尾页</FONT></a>"
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
                      <%
	  end if
	  rsSmallClass.close
	  set rsSmallClass=nothing	
%>
                    </td>
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
      <TD width="45"></TD>
    </TR>
  </TBODY>
</TABLE>
</div>
<!--#include file="foot.asp"-->
</body>
</html>