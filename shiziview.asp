<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<!--#include file="conn.asp"-->
<%
id=trim(request("id"))
id=replace(id,"'","")
id=left(id,8)
	if id="" then '
  		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�</font></p>"
		response.end
	end if	
	if  not isNumeric(id) then
  		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�</font></p>"
		response.end
	end if
    if len(id)>8 then 
		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�</font></p>"
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
sql="update Product set hits=hits+1 where id="&cstr(request("id"))
conn.execute sql
sql="select * from Product where id="&id
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
response.Write("���ݿ����")
else
title=rs("ProductName")
bigclass=rs("ProductBigClass")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% =title %>-<% =sitetitle %></title>
<meta name="keywords" content="<% =title %>,<% =webKeywords %>">
<meta name="description" content="<% =title %>,<% =webDescription %>">
<meta name="robots" content="index, follow">
<meta name="googlebot" content="index, follow">
<style>
<!--
.ss          { font-size: 16pt; font-weight: bold }
.anli        { font-family: ΢���ź�; font-size: 14pt; font-weight: bold }
.anli1       { font-family: ΢���ź�; font-size: 12pt; }

.jf          { font-size: 18pt; font-weight: bold }
td{font-family:arial}td{font-size:10pt;line-height:22px}.dhl         { font-size: 16pt; font-family: ΢���ź�; color:#0066CC; line-height:100%; margin-top:0; margin-bottom:0 }
.hl          { line-height: 200%; font-family: ΢���ź�; margin-top: 0; margin-bottom: 0 font-size:10pt; font-weight:bold}
.hl1          { line-height: 100%; font-family: ΢���ź�; margin-top: 0; margin-bottom: 0 }
-->
</style>
<style fprolloverstyle>A:hover {color: #CC0000; font-weight: bold}
.text        { font-family: ΢���ź�; font-size: 11pt }
</style>
</head>
<BODY bgcolor="#FFFFFF" topMargin=0 leftmargin="0" link="#5E5E5E" vlink="#5E5E5E" alink="#5E5E5E">

<BODY topMargin=0 leftmargin="0">
<!--#include file="top.asp"-->
<TABLE cellSpacing=0 cellPadding=0 width=1100 
 align=center 
border=0>
  <TBODY>
    <TR> 
      <TD width="0"></TD>
      <TD width="1100">
		<div align="center">
			<TABLE cellSpacing=0 cellPadding=0 width=1100 border=1 style="BORDER-TOP: #808080 1px solid;BORDER-LEFT: #808080 1px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 1px solid;">          <TBODY>
            <TR> 
              <TD width="260" vAlign=top bgcolor="#FFFFFF"  style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 0px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 0px solid;"> 
                <table border="0" width="240" cellspacing="0" height="32">
					<tr>
                      <TD height="32" align="left">
		                <table border="0" width="240" cellspacing="0" height="644">
							<tr>
		                      <TD height="60" align="center" bgcolor="#0099FF">
								<p class="anli" align="left"><font color="#4D4D4D">��</font><span style="font-weight: 700"><font color="#FFFFFF" face="΢���ź�">�Ŷӳ�Ա</font></span></TD>
		                    </tr>
							<tr>
								<td bgcolor="#FFFFFF"> 
		               				<table border="0" width="288" cellspacing="0" height="91">
               			 				<tr>
                          					<TD align="left" height="91" class="text">
			                                <TABLE cellSpacing=0 cellPadding=0 width="91%" border=0 height="42" >
			                                  <TBODY>
			                                    <TR> 
			                                      <TD align="left" height="42" background="bj.jpg" class="anli1">
													<p class="xuesheng" style="margin-top: 0; margin-bottom: 0">
													<font color="#004993" face="΢���ź�">
														<img border="0" src="images/about1.jpg" width="19" height="27">
														<span class="Apple-converted-space">&nbsp;</span>
														<a href="shizi.asp?type=48" style="text-decoration: none">��ʦ���</a>
													</font>
													</p>
			                                      </TD>
			                                    </TR>
			                                  </TBODY>
                                			</TABLE>
											</TD>
                        				</tr>
									</table>
									<p style="margin-top: -3px; margin-bottom: -3px" align="center">
									<table border="0" width="95%" height="142" cellspacing="0" cellpadding="0">
										<tr>
											<td bgcolor="#FFFFFF" height="142"> 
                								<table border="0" width="260" cellspacing="0" height="436"   style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 0px solid;BORDER-right: #808080 0px solid;BORDER-bottom: #808080 0px solid;">
													<tr>
														<td bgcolor="#FFFFFF" height="260"> 
															<font color="#666666">
															<table border="0" width="302" cellspacing="0" cellpadding="0" height="40" bgcolor="#004993">
															</table>
														</td>
													</tr>
												</table>
              								</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
            		  </TD>
                    </tr>
					</table>
				</TD>
              <TD width="842" vAlign=top bgcolor="#FFFFFF"> 
				<table border="0" width="97%">
					<tr>
						<td> 
				<div align="center">
				<TABLE cellSpacing=0 cellPadding=0 width="836" 
border=0 height="60">
                  <TBODY>
                    <TR> 
                      <TD width="836"  style="BORDER-BOTTOM: #CFCFCF 0px

 dashed" background="fl/tibj.gif" height="49">
						<p style="line-height: 200%; margin-top: 0px; margin-bottom:0; margin-right:40px" align="left" class="anli">
						<SPAN style="FONT-WEIGHT: bold;FONT-SIZE: 14px">
						<font color="#669900">��<img border="0" src="fl/jiantou.jpg" width="17" height="14"> </font></SPAN>
						<font face="΢���ź�"><a href="index.asp" style="text-decoration: none">
						<font color="#004993">��ҳ</font></a>
						<font color="#004993"> 
                        &gt;&gt;<a style="text-decoration: none; color: #004993;" href="shizi.asp?type=48"><%= rs("ProductBigClass") %></a><%if rs("ProductSmallClass")<>"" then%>
                        &gt;&gt; <%= rs("ProductSmallClass") %><%end if%>
						&gt;&gt; <%= rs("ProductName") %></b></font></font></TD>
                    </TR>
                  	<tr>
                      <TD height="5" align="center" background="images/lanmufeng.gif"></TD>
                    </tr>
					<tr>
                      <TD height="23" align="center"></TD>
                    </tr>
                  </TBODY>
                </TABLE>
				</div>
                <div align="center">
                <table width="835" border="0" cellpadding="0" cellspacing="0" height="25">
                  <tr> 
                    <td height="23" valign="top" align="left">
                <table width="835" border="0" align="center" cellpadding="0" cellspacing="0" height="82">
                  <tr> 
                    <td height="82" valign="top">
                      <TABLE width=834 
                        align=center cellPadding=0 cellSpacing=0>
                        <TBODY>
                          <TR> 
                            <TD style="PADDING-LEFT: 4px"> 
							<div align="center">
								<TABLE cellSpacing=0 cellPadding=0 width="98%" border=0 bgcolor="#4E2727">
                                <TBODY>
                                  <TR> 
                                    <TD bgColor=#FFFFFF> 
									<TABLE width="100%" 
                                border=0 cellPadding=0 cellSpacing=0 height="42">
                                        <TBODY>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD width="100%" height=50 colspan="4" bgcolor="#FFFFFF"> <DIV align=center>
												
											<font color="#4D4D4D"><STRONG> 
                                              <p style="margin-top: 0; margin-bottom: 0; line-height:100%" class="ss"> 
                                              <% =RS("ProductName") %>
                                              </STRONG></font></DIV>
											<DIV align=center>
												
											<font color="#4D4D4D">
											<p style="margin-top: 0; margin-bottom: 0; line-height:100%" class="ss"> 
		<img border="0" src="fl/fot02.jpg" width="820" height="24">��</font></DIV></TD>
                                          </TR>

  <TR bgcolor="#FFFFFF"> 
                                            <TD height=32 bgcolor="#FFFFFF" width="13%"> 
											<DIV align=right>
												<p style="line-height: 100%; margin-top: 0; margin-bottom: 0">
												<font color="#4D4D4D" face="΢���ź�">
												<strong>����ʱ�䣺</strong></font></DIV></TD>
                                            <TD height=32 width="16%"bgcolor="#FFFFFF" align="left"> <% =RS("ProductTime") %></TD>
                                            <TD height=32 bgcolor="#FFFFFF"> 
											<DIV align=right>
												<p style="line-height: 100%; margin-top: 0; margin-bottom: 0">
												<font color="#4D4D4D" face="΢���ź�">
												<strong>
												��עָ����</strong></font></DIV></TD>
                                            <TD bgcolor="#FFFFFF" align="left">
											<p style="line-height: 100%; margin-top: 0; margin-bottom: 0">
											<font color="#4D4D4D" face="΢���ź�">&nbsp; <%= rs("hits") %> ��</font></TD>
											
                                          </TR>

                                        
                                        </TBODY>
                                      </TABLE>
									<p style="margin-top: 0; margin-bottom: 0">
		<img border="0" src="fl/fot02.jpg" width="820" height="24"></TD>
                                  </TR>
                                </TBODY>
                              </TABLE>
							</div>
							</TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                      
                          </TR>
                        </TBODY>
                      </TABLE>
                      
                          </TR>
                        <tr>
                            <TD height=22 align="left">	
                                   <div align="center">
                <table width="818" border="0" cellpadding="0" cellspacing="0" height="25">
                  <tr> 
                    <td height="23" valign="top">
                <table width="710" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td height="174" valign="top">
                      <TABLE width=788 
                        align=center cellPadding=0 cellSpacing=0>
                        <TBODY>
                          <TR> 
                            <TD width=258 align=middle> <TABLE cellSpacing=0 
                              cellPadding=2 border=0>
                                <TBODY>
                                  <TR> 
                                    <TD width="251"> <span class="bh"> 
                                      <% if RS("ProductImages")="" then %>
                                      </span> 
									<TABLE width="250" height="210" border=0 cellSpacing=0 borderColor=#666666>
                                        <TBODY>
                                          <TR> 
                                            <TD bgColor=#FFE9BB>
											<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
											<img src="../images/wupic.jpg" width="359" height="298" border="0"></TD>
                                          </TR>
                                        </TBODY>
                                      </TABLE>
                                      <span class="bh"> 
                                      <% else %>
                                      </span> 
									<TABLE borderColor=#666666 cellSpacing=0 border=0 cellpadding="0">
                                        <TBODY>
                                          <TR> 
                                            <TD bgColor=#FFE9BB>
											<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
											<img src="<% Response.Write (""&RS("ProductImages")&"")%>" width="359" height="302" border="0"></TD>
                                          </TR>
                                        </TBODY>
                                      </TABLE>
                                      <span class="bh"> 
                                      <% end if %>
                                      </span></TD>
                                  </TR>
                                </TBODY>
                              </TABLE></TD>
                            <TD style="PADDING-LEFT: 4px"> 
							<div align="center">
								<TABLE cellSpacing=0 cellPadding=0 width="98%" border=0 bgcolor="#4E2727">
                                <TBODY>
                                  <TR> 
                                    <TD bgColor=#dfff80> 
									<TABLE width="100%" 
                                border=0 cellPadding=0 cellSpacing=0 height="218">
                                        <TBODY>
                                          <TR bgcolor="#FFFFFF"> 
                                              <TD width="568" height="54" bgColor=#ffffff>
														<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">��</p>
														<STRONG>
														<p class="anli" style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="left">
														</STRONG>														
														<font color="#2B2B2B">
														<STRONG><% =left(RS("ProductName"),20) %>
                                            			</STRONG>														
                                            			</font>
                                            			</p>														
														<p class="anli1" style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="left">
														<font color="#2B2B2B">
														<% =left(RS("ProductPrice5"),220) %></font></p>

														</TD>
                                          </TR>

                                        
                                        </TBODY>
                                      </TABLE></TD>
                                  </TR>
                                </TBODY>
                              </TABLE>
								<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">��</div>
							</TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                      
                          </TR>
                        </TBODY>
                      </TABLE>
                      
                          </TR>
                        <tr>
                            <TD height=22>	
							<p align="left" style="margin-top: 0; margin-bottom: 0">
		<img border="0" src="fl/fot02.jpg" width="820" height="14"><STRONG><font color="#EBEBEB">����</font></STRONG><div align="center">
								<table border="0" width="99%" height="32">
									<tr>
										<td background="../fl/cpjs.png">
										<p align="left" class="anli1" style="margin-top: 0; margin-bottom: 0">
										<b><font color="#4D4D4D">��ϸ����</font></b></td>
									</tr>
								</table>
							</div>
                            <div align="center">
								<table border="0" width="100%" height="32">
									<tr>
										<td height="30">
										<div align="center">
                      <TABLE width=796 height=10 border=0 cellPadding=0 cellSpacing=1 bgcolor="#FFFFFF">
                        <TBODY>
                          <TR vAlign=top> 
                            <TD 
                              height=1 bgcolor="#FFFFFF" align="left">
							<font color="#4D4D4D" face="΢���ź�"><%=rs("ProductContent") %></font></TD>
                          </TR>
                        </TBODY>
                      </TABLE>
                      
										<p>
		<img border="0" src="fl/fot02.jpg" width="820" height="14"></p>
                                                  
										</div>
										</td>
									</tr>
								</table>
                      
							</div>
							</TD>
                          </tr>
                  <tr> 
                    <td height="0" valign="top" align="left">
                      <div align="left">
                      
<script type=text/javascript> 
var speed=6 
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
								</font>
					</td>
				</tr>
			</table>
                                                                      
                              <font face="΢���ź�">
                                                                      
                              </div></font></td>
							</tr>
							</table>
						
						</div>							</TD>
                          </TR>
                      </TABLE>
                      </div>
                      </div>
                      
                          </TR>
                        </TBODY>
                      </TABLE>
						</div></td>
					</tr>
				</table>
				</td>
                  </tr>
                </table>
                </div>
                
  </TBODY>
</TABLE>
<% 
		end if
		rs.close
		set rs=nothing
		 %><!--#include file="foot.asp"-->