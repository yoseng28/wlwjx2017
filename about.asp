<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">

<!--#include file="conn.asp"-->
<!--#include file="inc/function.asp"--> 
<%
id=trim(request("type"))
id=replace(id,"'","")
id=left(id,8)
	if id="" then '
  		response.write "<br><p align=center><font color='red'>对不起，你访问的文件不存在。</font></p>"
		response.end
	end if	
	if  not isNumeric(request("id")) then
  		response.write "<br><p align=center><font color='red'>对不起，你访问的文件不存在。</font></p>"
		response.end
	end if
    if len(request("id"))>8 then 
		response.write "<br><p align=center><font color='red'>对不起，你访问的文件不存在。</font></p>"
		response.end
	end if
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select title,content,bigclassname from about where id="&id&""
rs.Open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<br><p align=center><font color='red'>对不起，你访问的文件不存在。2</font></p>"
		response.end
	else
		title=rs("title")
		content=rs("content")
		bigclass=rs("bigclassname")
	end if
rs.close
set rs=nothing
%>
<html>
<head>
<link rel="shortcut icon" href="images/new/wlw.ico" >
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% =sitetitle %></title>
<meta name="keywords" content="<%=bigclass%>,<%=title%>,<% =webKeywords %>">
<meta name="description" content="<% =webDescription %>">
<meta name="robots" content="index, follow">
<meta name="googlebot" content="index, follow">
<style fprolloverstyle>A:hover {text-decoration: none; color: #CC0000; font-weight: bold}
</style>
<style>
<!--
.text        { font-family: 微软雅黑; font-size: 12pt }
.anli        { font-family: 微软雅黑; font-size: 14pt; font-weight: bold }

-->
</style>
</head>
<BODY topMargin=0 leftmargin="0" bgcolor="#FFFFFF" link="#4D4D4D" vlink="#4D4D4D" alink="#4D4D4D">
<!--#include file="top.asp"-->
<TABLE cellSpacing=0 cellPadding=0 width=1091 align=center bgColor=#ffffff style="BORDER-TOP: #808080 1px solid;BORDER-LEFT: #808080 1px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 1px solid;">
        
  <TBODY>
    <TR> 
      <TD width="1080">
		<div align="center">
		<TABLE width=1080 border=0 cellPadding=0 cellSpacing=0>
          <TBODY>
            <TR> 
              
              
              <TD width="1080" vAlign=top style="BORDER-TOP: #C5CED7 0px solid;BORDER-LEFT: #C5CED7 0px solid;BORDER-right: #C5CED7 0px solid;BORDER-bottom: #C5CED7 0px solid;"> 
                <TABLE cellSpacing=0 cellPadding=0 width="1080" 
border=0 height="59">
                  <TBODY>
                    <TR> 
                      <TD width="1080" height="52" align="center" background="fl/cp1.gif"> 
                        
                      <p align="left" class="anli" style="margin-top: 0; margin-bottom: 0"> 
                        
                      <b><font face="微软雅黑"><font color="#4D4D4D"> 
                        
                      <img border="0" src="fl/jiantou.jpg" width="17" height="14">
						</font><a href="index.asp" name="1"><font color="#4D4D4D">首页</font></a><font color="#4D4D4D"> 
						&gt;&gt; <%=bigclass%> 
                        <%if title<>"" then %>
                        &gt;&gt; <%=title%> 
                        <% end if %>
                      </font> 
                        
                      </font></b> 
                        
                      </TD>
                      <TD width="101" background="fl/cp1.gif">　</TD>
                    </TR>
                    <TR> 
                      <TD height="7" colspan="2" align="left" background="images/lanmufeng.gif"></TD>
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
                <div align="left"  class="text" >
                <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td><SPAN style="LINE-HEIGHT: 25px">
					<font color="#333333" face="微软雅黑"><%=content %></font></SPAN></td>
                  </tr>
                </table>
                </div>
                <TABLE cellSpacing=0 cellPadding=0 width=100 align=center bgColor=#ffffff 
border=0>
                  <TBODY>
                    <TR> 
                      <TD height=8></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
            </TR>
          </TBODY>
        </TABLE></div>
		</TD>
      
    </TR>
  </TBODY>
</TABLE>
<!--#include file="foot.asp"-->