<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">

<!--#include file="conn.asp"-->
<% 
'��ȡ�������
bigclassid=trim(request("type"))
smallclassid=trim(request("type2"))
bigclassid=replace(bigclassid,"'","")
smallclassid=replace(smallclassid,"'","")
bigclassid=left(bigclassid,8)
smallclassid=left(smallclassid,8)
	if bigclassid="" then '
  		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ�������1��</font></p>"
		response.end
	elseif  not isNumeric(bigclassid) then
  		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ�������2��</font></p>"
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
	    response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ�������3��</font></p>"
		response.end
	else
	    sqls="select SmallClassName from SmallClass where SmallClassID="&SmallClassID&""
        set rss=server.createobject("ADODB.Recordset")
        rss.open sqls,conn,1,1
        if not rss.eof then
          smallclass=rss("SmallClassName")
        else
        response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ�������4��</font></p>"
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
<BODY topMargin=0 leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0" link="#004993" vlink="#0099FF" alink="#0099FF">
<!--#include file="top.asp"-->
<TABLE cellSpacing=0 cellPadding=0 width=1267 align=center bgColor=#ffffff 
border=0>
  <TBODY>
    <TR> 
      <TD width="5" bgcolor="#FFFFFF"></TD>
      <TD width="1258">
		<div align="center">
			<TABLE width=1000 border=0 cellPadding=0 cellSpacing=0>
          <TBODY>
            <TR> 
              <TD width="196" vAlign=top style="background-repeat: repeat-x"
 bgcolor="#FFFFFF"> 
               
                <table border="0" width="260" cellspacing="0" height="436"   style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 0px solid;BORDER-right: #808080 0px solid;BORDER-bottom: #808080 0px solid;">
					<tr>
                      <TD height="49" align="center" bgcolor="#0099FF">
						<p align="left" class="title" style="margin-top: 0; margin-bottom: 0">
										<font color="#006699" face="΢���ź�">��</font><span style="font-weight: 700"><font color="#FFFFFF" face="΢���ź�">��������</font></span></TD>
                    </tr>
					<tr>
                      <TD height="99" bgcolor="#FFFFFF">
                                <TABLE cellSpacing=0 cellPadding=0 width="91%" 
                              border=0 height="87" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px" >
                                  <TBODY style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
                                    <TR style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"> 
                                      <TD align="left" height="48" background="bj.jpg" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 7px">
										<font color="#004993" face="΢���ź�" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">��</font><font face="΢���ź�" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"><img border="0" src="images/about1.jpg" width="19" height="27" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"><span class="Apple-converted-space">&nbsp;</span></font><font face="΢���ź�"><a style="text-decoration: none" href="download.asp?type=79&type2=40">�μ�����</a></font></TD>
                                      
                                    </TR>
                                    <TR style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"> 
                                      <TD align="left" height="13" background="bj.jpg" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 7px">
										<font color="#004993" face="΢���ź�" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">��</font><font face="΢���ź�" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"><img border="0" src="images/about1.jpg" width="19" height="27" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"><span class="Apple-converted-space">&nbsp;</span><a style="text-decoration: none" href="download.asp?type=79&type2=41">������</a></font></TD>
                                      
                                    </TR>
                                    <TR style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"> 
                                      <TD align="left" height="40" background="bj.jpg" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
										<font color="#004993" face="΢���ź�" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">��</font><font face="΢���ź�" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"><img border="0" src="images/about1.jpg" width="19" height="27" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"><span class="Apple-converted-space">&nbsp;</span><a href="download.asp?type=79&type2=42" style="text-decoration: none; font-weight: 700">��������</a></font></TD>
                                      
                                    </TR>
                                    <TR style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px"> 
                                      <TD align="left" height="40" background="bj.jpg" style="font-size: 16px; border: 0px none; margin: 0px; padding: 0px">
										��</TD>
                                      
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
										<font color="#006699" face="΢���ź�">��</font><span style="font-weight: 700"><font color="#FFFFFF" face="΢���ź�">�ŶӶ�̬</font></span></td>
							</tr>
							<tr>
								<td height="10" bgcolor="#FFFFFF">
								��</td>
							</tr>
				</table>
						<table border="0" width="101%" cellspacing="0">
							<tr>
								<td> 
                <table border="0" width="252" cellspacing="0" height="253">
					<tr>
						<td bgcolor="#FFFFFF"> 
				<div align="center" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
								<table border="1" bgcolor="#FFFFFF" cellspacing="0" bordercolor="#CFCFCF" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" width="221">
									<tbody style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
									<tr style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
										<td style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
							<body rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0" link="#004993" vlink="#0099FF" alink="#0099FF">
<div id="demo" style="border:0px none; margin:0px; padding:0px; overflow:hidden;height:209px;width:295px; font-size:15px">
<ul id="demo1" style="height: auto; text-align: left; border: 0px none; margin: 0px; padding: 0px; list-style-type:none; font-size:15px">  
<font color="#000F1A" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" face="΢���ź�">
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
						<font face="΢���ź�" style="font-size: 15px; border: 0px none; margin: 0px; padding: 0px">
						<font color="#000F1A" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">��  
						</font>  
						<A 
                              title=<%=oRS("title")%> 
                              href="newsview.asp?id=<%=oRS("id")%>" 
                               style="border:0px none; margin:0px; padding:0px; text-decoration: none; font-size:15px">
						<font color="#000F1A" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
						<%=left(oRS("title"),15)%></font></A></font></TD>
						<TD class=sd13 height=27 width="56">
						<p align="center" style="margin-top: 0; margin-bottom: 0">
						<font color="#000F1A" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" face="΢���ź�">(<% =month(ors("infotime")) %>.<% =day(ors("infotime")) %>)</font></TD>
                    </TR>

                  </TBODY>
                </TABLE>
                <font color="#FFFF00" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" face="΢���ź�">
                <b style="border: 0px none; margin: 0px; padding: 0px; font-size:15px"><%
oRS.MoveNEXT
end if
next
else Response.Write ("��������.") end if
oRS.close 
Set oRS = Nothing
%>
              </b>
              </font></ul> 
<div id="demo2" style="height: auto; text-align: left; border: 0px none; margin: 0px; padding: 0px; font-size:15px">
	<p style="margin-top: 0; margin-bottom: 0"></div>
</div> 
							<font color="#FFFF00" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" face="΢���ź�"> 
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
                ��</td>
					</tr>

					<tr>
						<td bgcolor="#FFFFFF" width="248" height="130" align="left"> 
              	��</td>
					</tr>

					</table>
              					</td>
							</tr>
						</table></td>
					</tr>
					<tr>
						<td bgcolor="#FFFFFF"> 
										��</td>
					</tr>
				</table>
              </TD>
              <TD width="4">��</TD>
              <TD width="820" vAlign=top style="border-style:solid; border-width:1px; "> 
                <TABLE cellSpacing=0 cellPadding=0 width="947" 
border=0>
                  <TBODY>
                    <TR> 
                      <TD width="10" height="45" align="center" background="fl/cp1.gif">��</TD>
                      <TD width="847" background="fl/cp1.gif" height="50"> 
						<p align="left" class="title"><font face="΢���ź�">
						<font color="#004993">��</font><font color="#FF0000"> </font>
						<b><font color="#004993">��ҳ 
                        &gt;&gt; <%=bigclass%>  
                      </font>  
                      </b></font>  
                      </TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <div align="center">
                <table width="822" border="0" cellpadding="0" cellspacing="0" style="BORDER-RIGHT: #CDD0E4 1px solid; BORDER-LEFT: #CDD0E4 1px solid; BORDER-BOTTOM: #CDD0E4 1px solid">
                  <tr> 
                    <td width="820"> <DIV align=center> 
                        <font face="΢���ź�"> 
                        <% 
page=clng(request("page"))		 
Set rs=Server.CreateObject("ADODB.RecordSet") 
if bigclass<>"" and smallclass<>"" then
sql="select * from download where downBigClass='"&bigclass&"' and downSmallClass='"&smallclass&"' order by id desc"
rs.Open sql,conn,1,1
elseif smallclass="" then
sql="select * from download where downBigClass='"&bigclass&"' order by id desc"
rs.Open sql,conn,1,1
end if
if rs.eof and rs.bof then
response.Write("��������")
else 
%>
                        </font>
                        <TABLE width="943" border=0 cellPadding=0>
                          <TBODY>
                            <TR> 
                              <TD height="35" bordercolorlight="#004993" bgcolor="#CFCFCF" class="title" width="539"> <div align="center"></div>
								<p align="left" class="shuye">
								<font color="#575757" face="΢���ź�">&nbsp;<SPAN style="FONT-WEIGHT: bold;">���ļ�����</SPAN></font></TD>
                              <TD height="35" bordercolorlight="#004993" bgcolor="#CFCFCF" class="title"> 
								<p align="center" class="shuye">����</TD>
                              <TD width="193" align="left" bordercolorlight="#004993" bgcolor="#CFCFCF" class="shuye" height="35">
								<p align="center" class="shuye"><font face="΢���ź�"><B>
								<font color="#575757">
								<SPAN style="FONT-WEIGHT: bold;">��������</SPAN></font></B></font></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE width="941" border=0 cellpadding="0" cellSpacing=0 id=table5>
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
                              <TD width="21" height="24" 
                        align=left style="BORDER-BOTTOM: #74B5EB 1px dotted"> 
                                <div align="center"><font color="#004993">��</font> 
                                </div></TD>
                              <TD height="25" 
                        align=left style="BORDER-BOTTOM: #74B5EB 1px dotted"> 
                                <font face="΢���ź�">&nbsp;<a href="downloadview.asp?id=<%= RS("id") %>" target="_blank" style="text-decoration: none"><%= left(RS("DownName"),35) %></a>&nbsp;</font></TD>
                              <TD height="25" 
                        align=left style="BORDER-BOTTOM: #74B5EB 1px dotted" width="202"> 
                                <p align="center"> 
								<a href="<% =RS("downlink") %>" style="text-decoration: none">
								<font color="#60616C">�������</font></a></font></TD>
                              <TD width="196" 
                        align=left style="BORDER-BOTTOM: #74B5EB 1px dotted">
								<p align="center">
								<font face="΢���ź�">&nbsp;<%= RS("DownTime") %></font></TD>
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
                                  <font face="΢���ź�"> 
                                  <%if Page<2 then      
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=download.asp?type="&bigclassid&"&type2="&smallclassid&"&page=1><font color=red>��ҳ</font></a>&nbsp;"
    response.write "<a href=download.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & Page-1 & "><font color=red>��һҳ</font></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=download.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & (page+1) & ">"
    response.write "<font color=red>��һҳ</font></a> <a href=download.asp?type="&bigclassid&"&type2="&smallclassid&"&page="&rs.pagecount&"><font color=red>βҳ</font></a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#FF0000'>"&rs.recordcount&"</font></b>�� <b>"&rs.pagesize&"</b>��/ҳ"
%>
                                </font>
                                </div></td>
                            </tr>
                          </table>
                        </CENTER>
                      </DIV>
                      <font face="΢���ź�">
                      <% 
end if
rs.close
set rs=nothing
%>
                    </font>
                    </td>
                  </tr>
                </table>
                </div>
                <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
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
        </TABLE></div>
		</TD>
      <TD width="4"></TD>
    </TR>
  </TBODY>
</TABLE>
<!--#include file="foot.asp"-->
</body>
</html>