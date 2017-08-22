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
<style fprolloverstyle>A:hover {text-decoration: none; color: #004993; font-weight: bold}
</style>
<style>
<!--
a:link {text-decoration: none;  }
.text        { font-family: 微软雅黑; font-size: 12pt }
.anli        { font-family: 微软雅黑; font-size: 14pt; font-weight: bold }

-->
</style>
</head>
<BODY bgcolor="#FFFFFF" topMargin=0 leftmargin="0" link="#5E5E5E" vlink="#5E5E5E" alink="#5E5E5E">
<!--#include file="top.asp"-->
<TABLE cellSpacing=0 cellPadding=0 width=1100 align=center bgColor=#ffffff vAlign=top style="BORDER-TOP: #808080 1px solid;BORDER-LEFT: #808080 1px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 1px solid;"

border=0>
  <TBODY>
    <TR> 
      <TD width="1100"><TABLE cellSpacing=0 cellPadding=0 width=1100 align=center border=0>
          <TBODY>
            <TR> 
            
              <TD width="1100"  vAlign=top style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 1px solid;BORDER-right: #808080 0px solid;BORDER-bottom: #808080 0px solid;"> 
				<TABLE cellSpacing=0 cellPadding=0 width="1100" 
border=0 height="61" bgcolor="#FFFFFF">
                  <TBODY>
                    <TR> 
                      <TD width="1100" height="55" align="left" style="BORDER-BOTTOM: #CFCFCF 0px dashed" background="fl/tibj.gif"> 
                        
                      <p class="anli" style="margin-top: 0; margin-bottom: 0"> 
                        
                      <img border="0" src="fl/jiantou.jpg" width="17" height="14"><font face="微软雅黑"><a href="index.asp" style="text-decoration: none"><SPAN style="FONT-WEIGHT: bold;"><font color="#669900">
						</font><font color="#404A53">
						首页</font></SPAN></a><font color="#404A53"><b> 
                        &gt;&gt; <%=bigclass%><%if smallclass<>"" then%> &gt;&gt; <%=smallclass%><%end if%></b></font></font></TD>
                    </TR>
                     <TR> 
                      <TD height="6" align="center" background="../images/lanmufeng.gif"></TD>
                    </TR>

                  </TBODY>
                </TABLE>
<%if xsfs="图文列表" then%>
                <div align="center">
                <table width="96%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
                  <% 
page=clng(request("page"))		 
Set rs=Server.CreateObject("ADODB.RecordSet") 
if bigclass<>"" and smallclass<>"" then
sql="select * from product where ProductBigClass='"&bigclass&"' and ProductSmallClass='"&smallclass&"' order by id desc"
rs.Open sql,conn,1,1
elseif smallclass="" then
sql="select * from product where ProductBigClass='"&bigclass&"' order by id desc"
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
                                            <TD width="120" height="108" bgColor=#ffffff><a href="productview.asp?id=<%= RS("id") %>" target="_blank">
											<img src="images/wupic.jpg" width="120" height="100" border="0"></a></TD>
                                          </TR>
                                        </TBODY>
                                      </TABLE>
                                      <span class="bh"> 
                                      <% else %>
                                      </span> <TABLE width="130" height="110" border=1 
                                cellPadding=3 cellSpacing=0 borderColor=#b2b2b2>
                                        <TBODY>
                                          <TR> 
                                            <TD width="120" height="108" bgColor=#ffffff><a href="productview.asp?id=<%= RS("id") %>" target="_blank"><img src="<% Response.Write (""&RS("ProductImages")&"")%>" width="120" height="100" border="0"></a></TD>
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
                                            <TD width="17%" height=21> <DIV align=center><strong>产品名称：</strong></DIV></TD>
                                            <TD width="83%">&nbsp;<STRONG> <a href="spview.asp?id=<%= RS("id") %>" target="_blank"> 
                                              <% =RS("ProductName") %>
                                              </a></STRONG></TD>
                                          </TR>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD height=21> <DIV align=center><strong>产品规格：</strong></DIV></TD>
                                            <TD>&nbsp; 
                                              <% =RS("ProductType") %>
                                            </TD>
                                          </TR>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD height=21> <DIV align=center><strong>产品价格：</strong></DIV></TD>
                                            <TD>&nbsp; 
                                              <% =RS("ProductPrice") %>
                                            </TD>
                                          </TR>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD height=21> <DIV align=center><strong>关注指数：</strong></DIV></TD>
                                            <TD>&nbsp; 浏览：<%= rs("hits") %> 次</TD>
                                          </TR>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD height=21> <DIV align=center><strong>产品介绍：</strong></DIV></TD>
                                            <TD bgcolor="#FFFFFF">&nbsp;<strong><a href="productview.asp?id=<%= RS("id") %>" target="_blank"><font color="#FF0000">点击查看&gt;&gt;</font></a></strong></TD>
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
    response.write "<a href=zxktsp.asp?type="&bigclassid&"&type2="&smallclassid&"&page=1><font color=red>首页</font></a>&nbsp;"
    response.write "<a href=zxktsp.asp?type="&bigclassid&"&type2="&smallclass&"&page=" & Page-1 & "><font color=red>上一页</font></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=zxktsp.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & (page+1) & ">"
    response.write "<font color=red>下一页</font></a> <a href=zxktsp.asp?type="&bigclassid&"&type2="&smallclassid&"&page="&rs.pagecount&"><font color=red>尾页</font></a>"
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
				</div>
				<% else %>
                <div align="center">
                <table width="100%" border="0" cellpadding="0" cellspacing="0" height="475" bgcolor="#FFFFFF">
                  <% 
page=clng(request("page"))		 
Set rs=Server.CreateObject("ADODB.RecordSet") 
if bigclass<>"" and smallclass<>"" then
sql="select * from product where ProductBigClass='"&bigclass&"' and ProductSmallClass='"&smallclass&"' order by id desc"
rs.Open sql,conn,1,1
elseif smallclass="" then
sql="select * from product where ProductBigClass='"&bigclass&"' order by id desc"
rs.Open sql,conn,1,1
end if
if rs.eof and rs.bof then
response.Write("暂时没有内容")
else 
%>
                  <% 
rs.PageSize=16
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
                      <div align="center">
                      <TABLE cellSpacing=0 cellPadding=0 width=841>
                        <TBODY>
                          <TR> 
                            <TD width=841 align=middle>
<table border="0" width="244">
                                <% 
							   for j=1 to 4 
							   if NOT RS.EOF then
							   %>
                                <tr>
                                  <% 
							   for i=1 to 4 
							   if NOT RS.EOF then
							   %> 
                                  <td width="238">
<TABLE width="199" border=0 align="center" 
                              cellPadding=2 cellSpacing=0>
                                      <TBODY>
                                        <TR> 
                                          <TD width="195"> <span class="bh"> 
                                            <% if trim(RS("ProductImages"))="" then %>
                                            </span> 
											<TABLE width="182" height="149" border=0 
                                cellPadding=3 cellSpacing=0 borderColor=#b2b2b2>
                                              <TBODY>
                                                <TR> 
                                                  <TD width="240" height="200" bgColor=#ffffff>
													<p style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="right">
													<a href="spview.asp?id=<%= RS("id") %>">
													<img src="../images/wupic.jpg" width="245" height="188" border="0" style="border: 0px solid #808080; padding: 3px"></a></TD>
                                                </TR>
                                              </TBODY>
                                            </TABLE>
                                            <span class="bh"> 
                                            <% else %>
                                            </span> 
											<TABLE width="191" height="189" border=0 
                                cellPadding=0 cellSpacing=0 borderColor=#b2b2b2>
                                              <TBODY>
                                                <TR> 
                                                  <TD width="187" height="187" bgColor=#ffffff>
													<p align="center" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><a href="spview.asp?id=<%= RS("id") %>" >
													<img src="<% Response.Write (""&RS("ProductImages")&"")%>" width="249" height="200" border="0" style="border: 0px solid #808080; padding: 3px"></a></TD>
                                                </TR>
                                              </TBODY>
                                            </TABLE>
                                            <span class="bh"> 
                                            <% end if %>
                                            </span></TD>
                                        </TR>
                                        <TR> 
                                          <TD height="28" align="center"><STRONG>
                                            <font face="微软雅黑"> 
                                            <a style="text-decoration: none" href="spview.asp?id=<%= RS("id") %>"> 
                                            <% =left(RS("ProductName"),26) %>
                                            </a>
                                            </font>
                                            </STRONG></TD>
                                        </TR>
                                      </TBODY>
                                    </TABLE>
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
                    <td height="26" colspan="3" align="center" > 
                      <font color="#676767"> 
                      <%if Page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=zxktsp.asp?type="&bigclassid&"&type2="&smallclassid&"&page=1><font color=red>首页</font></a>&nbsp;"
    response.write "<a href=zxktsp.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & Page-1 & "><font color=red>上一页</font></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=zxktsp.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & (page+1) & ">"
    response.write "<font color=red>下一页</font></a> <a href=zxktsp.asp?type="&bigclassid&"&type2="&smallclassid&"&page="&rs.pagecount&"><font color=red>尾页</font></a>"
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
                </div>
                <% end if %></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
     
    </TR>
  </TBODY>
</TABLE>
<!--#include file="foot.asp"-->
</body>
</html>