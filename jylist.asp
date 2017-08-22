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
<style fprolloverstyle>A:hover {color: #004993; font-weight: bold}
</style>
<style>
<!--
.xuesheng    { font-size: 12pt; font-weight: bold }
.title        { font-size: 14pt }
.shuye1        { font-size: 11pt }

-->
</style>
</head>
<BODY topMargin=0 leftmargin="0" link="#515753" vlink="#515753" alink="#515753">
<!--#include file="top.asp"-->
<div align="center">
<TABLE cellSpacing=0 cellPadding=0 width=1100 bgColor=#ffffff 
border=0>
  <TBODY>
    <TR>      
      <TD width="1100">
		<TABLE width=1100 border=0 align=center cellPadding=0 cellSpacing=0  vAlign=top style="BORDER-TOP: #808080 1px solid;BORDER-LEFT: #808080 1px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 1px solid;">
          <TBODY>
            <TR> 
              <TD width="260" vAlign=top style="background-repeat: repeat-x"
 bgcolor="#FFFFFF"> 
                
                <div align="center">
                <table border="0" width="260" cellspacing="0" height="436" style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 0px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 0px solid;">
					<tbody style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
					<tr style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
                      <TD height="49" align="center" bgcolor="#0099FF" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
						<p align="left" class="title" style="margin-top: 0; margin-bottom: 0">
										<font color="#006699" face="微软雅黑" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">　</font><span style="font-weight: 700"><font color="#FFFFFF" face="微软雅黑">教学情况</font></span></TD>
                    </tr>
					<tr style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
                      <TD height="99" bgcolor="#FFFFFF" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
                                <div align="left">
                                <TABLE cellSpacing=0 cellPadding=0 width="99%" 
                              border=0 height="100%" >
                                  <TBODY>
                                    <TR> 
                                      <TD align="left" height="40" background="bj.jpg">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 0">
										<font color="#004993" face="微软雅黑">　</font><font face="微软雅黑"><img border="0" src="images/about1.jpg" width="19" height="27"><span class="Apple-converted-space">&nbsp;</span><a style="text-decoration: none" href="jylist.asp?type=73&type2=47">英语专业人才培养计划</a></font></TD>
                                      
                                    </TR>
                                    <TR> 
                                      <TD align="left" height="40" background="bj.jpg">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 0">
										<font color="#004993" face="微软雅黑">　</font><font face="微软雅黑"><img border="0" src="images/about1.jpg" width="19" height="27"><span class="Apple-converted-space">&nbsp;</span><a style="text-decoration: none" href="jylist.asp?type=73&type2=48">课堂教学</a>　　</font></TD>
                                      
                                    </TR>
                                    <TR> 
                                      <TD align="left" height="40" background="bj.jpg">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 0">
										<font color="#004993" face="微软雅黑">　</font><font face="微软雅黑"><img border="0" src="images/about1.jpg" width="19" height="27"><span class="Apple-converted-space">&nbsp;</span><a style="text-decoration: none" href="jylist.asp?type=73&type2=49">实践教学</a>　　</font></TD>
                                      
                                    </TR>
                                    <TR> 
                                      <TD align="left" height="40" background="bj.jpg">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 0">
										<font face="微软雅黑">　<img border="0" src="images/about1.jpg" width="19" height="27"><span class="Apple-converted-space">&nbsp;</span><a style="text-decoration: none" href="jylist.asp?type=73&type2=50">创新创业教育</a></font></TD>
                                      
                                    </TR>
                                    <TR> 
                                      <TD align="left" height="40" background="bj.jpg">
										<font face="微软雅黑">　</font><b><font face="微软雅黑"><img border="0" src="images/about1.jpg" width="19" height="27"><a style="text-decoration: none" href="jylist.asp?type=73&type2=51"> 
										教学资源</a></font></b></TD>
                                      
                                    </TR>
                                    <TR> 
                                      <TD align="left" height="40" background="bj.jpg">
										<font face="微软雅黑">　</font><b><font face="微软雅黑"><img border="0" src="images/about1.jpg" width="19" height="27">
										<a style="text-decoration: none" href="jylist.asp?type=73&type2=52">
										教材建设</a></font></b></TD>
                                      
                                    </TR>
                                    <TR> 
                                      <TD align="left" height="40" background="bj.jpg">
										<font face="微软雅黑">　<img border="0" src="images/about1.jpg" width="19" height="27">
										<span style="font-weight: 700">
										<a style="text-decoration: none" href="jylist.asp?type=73&type2=53">
										教学改革</a></span></font></TD>
                                      
                                    </TR>
                                    <tr>
                                      <TD align="left" height="40" background="bj.jpg">
										<font face="微软雅黑">　<img border="0" src="images/about1.jpg" width="19" height="27">
										<span style="font-weight: 700">
										<a style="text-decoration: none" href="jylist.asp?type=73&type2=54">
										授课情况</a></span></font></TD>
                                      
                                    </tr>
                                  </TBODY>
                                </TABLE>
								</div>
								<p>　</TD>
                    </tr>
					<tr style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
						<td bgcolor="#FFFFFF" height="260" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"> 
						<font color="#666666" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
						<table border="0" width="302" cellspacing="0" cellpadding="0" height="40" bgcolor="#004993" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
							<tbody style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
							<tr style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
								<td height="50" bgcolor="#0099FF" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
								<p class="title" style="margin-top: 0; margin-bottom: 0" align="left">
										<font color="#006699" face="微软雅黑" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">　</font><span style="border:0px none; margin:0px; padding:0px; font-weight: 700; "><font color="#FFFFFF" face="微软雅黑" style="border: 0px none; margin: 0px; padding: 0px">团队动态</font></span></td>
							</tr>
							<tr style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
								<td height="10" bgcolor="#FFFFFF" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
								　</td>
							</tr>
				</table>
						<table border="0" width="101%" cellspacing="0" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
							<tbody style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
							<tr style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
								<td style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"> 
                <table border="0" width="252" cellspacing="0" height="253" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
					<tbody style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
					<tr style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
						<td bgcolor="#FFFFFF" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"> 
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
                              href="gonggaoview.asp?id=<%=oRS("id")%>" 
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
					

					</table>
              					</td>
							</tr>
						</table></td>
					</tr>
					</table>
              	</div>
				</TD>
              <TD width="1" bgcolor="#808080"></TD>
              <TD width="850" vAlign=top style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 0px solid;BORDER-right: #808080 0px solid;BORDER-bottom: #808080 0px solid;"> 
                <TABLE cellSpacing=0 cellPadding=0 width="850" 
border=0 height="63">
                  <TBODY>
                    <TR> 
                      <TD width="694" align="right" height="57">
						<p align="left" class="title"><b><font color="#205081">
						<img border="0" src="fl/jiantou.jpg" width="17" height="14"></font></b><font color="#663300" face="微软雅黑"> </font>
						<font face="微软雅黑"><font color="#004993">&nbsp;</font><b><font color="#004993">首页 
                        &gt;&gt; </font> 
						<a href="cxjy.asp?type=<%=bigclassid%>" style="text-decoration: none">
						<font color="#004993"><%=bigclass%></font></a><font color="#004993"> 
                        &gt;&gt; <%=smallclass%> 
                        </font> 
                        </b> 
                        </font> 
                        </TD>
                      <TD width="282">　</TD>
                    </TR>
                    <TR> 
                      <TD height="6" colspan="2" align="center" background="images/lanmufeng.gif"></TD>
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
                            <TR height=24> 
                              <TD height="40" 
                        align=left width="716"> <font face="微软雅黑">                             
                                                             
                               <a href="cxjyview.asp?id=<% =rs("id") %>" style="text-decoration: none">
								<p class="shuye1" style="margin-top: 0; margin-bottom: 0; margin-left:16px"> 
                                      <font color="#666666"><img border="0" src="fl/icon03.gif" width="4" height="4"> <% =rs("title") %></p>
                                </font>                 

                                </TD>
                              <TD height="18" 
                        align=center width="134" class="shuye1"> 
								<p style="margin-top: 0; margin-bottom: 0"> <% =year(rs("infotime")) %>-<% =month(rs("infotime")) %>-<% =day(rs("infotime")) %> </TD>
                            </TR>
                            
                            <TR> 
                                    <TD height="2" vAlign=middle class=listb background="images/newx_fenge.gif" colspan="2"></TD>
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
                              <td height="40" class="shuye1"><div align="center"> 
                                  <font color="#666666"> 
                                  <%if Page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=jylist.asp?type="&bigclassid&"&type2="&smallclassid&"&page=1><FONT color=#ff0000>首页</FONT></a>&nbsp;"
    response.write "<a href=jylist.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & Page-1 & "><FONT color=#ff0000>上一页</FONT></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=jylist.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & (page+1) & ">"
    response.write "<FONT color=#ff0000>下一页</FONT></a> <a href=jylist.asp?type="&bigclassid&"&type2="&smallclassid&"&page="&rs.pagecount&"><FONT color=#ff0000>尾页</FONT></a>"
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
      <TD width="4"></TD>
    </TR>
  </TBODY>
</TABLE>
</div>
<!--#include file="foot.asp"-->
</body>
</html>