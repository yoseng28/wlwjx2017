<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<!--#include file="conn.asp"-->
<%
id=trim(request("id"))
id=replace(id,"'","")
id=left(id,8)
	if id="" then '
  		response.write "<br><p align=center><font color='red'>¶Ô²»Æð£¬Äã·ÃÎÊµÄÎÄ¼þ²»´æÔÚ¡£</font></p>"
		response.end
	end if	
	if  not isNumeric(id) then
  		response.write "<br><p align=center><font color='red'>¶Ô²»Æð£¬Äã·ÃÎÊµÄÎÄ¼þ²»´æÔÚ¡£</font></p>"
		response.end
	end if
    if len(id)>8 then 
		response.write "<br><p align=center><font color='red'>¶Ô²»Æð£¬Äã·ÃÎÊµÄÎÄ¼þ²»´æÔÚ¡£</font></p>"
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
response.Write("Êý¾Ý¿â³ö´í")
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
.xuesheng    { font-size: 12pt; font-weight: bold }

.ss          { font-size: 16pt; font-weight: bold }
.anli        { font-family: Î¢ÈíÑÅºÚ; font-size: 14pt; font-weight: bold }
.jf          { font-size: 18pt; font-weight: bold }
td{font-family:arial}td{font-size:10pt;line-height:22px}.dhl         { font-size: 16pt; font-family: Î¢ÈíÑÅºÚ; color:#0066CC; line-height:100%; margin-top:0; margin-bottom:0 }
.hl          { line-height: 200%; font-family: Î¢ÈíÑÅºÚ; margin-top: 0; margin-bottom: 0 font-size:10pt; font-weight:bold}
.hl1          { line-height: 100%; font-family: Î¢ÈíÑÅºÚ; margin-top: 0; margin-bottom: 0 }
-->
</style>
<style fprolloverstyle>A:hover {color: #CC0000; font-weight: bold}
.text        { font-family: Î¢ÈíÑÅºÚ; font-size: 11pt }
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
      <TD width="1100
">
		<div align="center">
			<TABLE cellSpacing=0 cellPadding=0 width=1100 border=1 style="BORDER-TOP: #808080 1px solid;BORDER-LEFT: #808080 1px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 1px solid;">          <TBODY>
            <TR> 
              <TD width="260" vAlign=top bgcolor="#FFFFFF"  style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 0px solid;BORDER-right: #808080 1px solid;BORDER-bottom: #808080 0px solid;"> 
               
                <table border="0" width="240" cellspacing="0" height="32">
					<tr>
                      <TD height="32" align="left">
                <table border="0" width="240" cellspacing="0" height="644">
					<tr>
                      <TD height="50" align="center" bgcolor="#0099FF">
						<p class="anli" align="left"><font color="#4D4D4D">¡¡</font><font color="#FFFFFF">³É¹ûÕ¹Ê¾</font></TD>
                    </tr>
                                   

					<tr>
						<td bgcolor="#FFFFFF"> 
                <table border="0" width="305" cellspacing="0" height="114">
                <tr>
                                      <TD align="left" height="114" class="text">
                                <TABLE cellSpacing=0 cellPadding=0 width="99%" 
                              border=0 height="85" >
                                  <TBODY>
                                    <TR> 
                                      <TD align="left" height="44" background="bj.jpg" style="font-size: 10pt; line-height: 22px; font-family: arial">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 0; line-height:200%">
										<font color="#004993" face="Î¢ÈíÑÅºÚ">¡¡</font><font face="Î¢ÈíÑÅºÚ"><img border="0" src="images/about1.jpg" width="19" height="27"><span class="Apple-converted-space">&nbsp;</span><a href="zjsp.asp?type=65&type2=30" style="text-decoration: none">½ÌÑÐ³É¹û</a></font></TD>
                                      
                                    </TR>
                                    <TR> 
                                      <TD align="left" height="41" background="bj.jpg" style="font-size: 10pt; line-height: 22px; font-family: arial">
										<p class="xuesheng" style="margin-top: 0; margin-bottom: 0; line-height:200%">
										<font color="#004993" face="Î¢ÈíÑÅºÚ">¡¡</font><font face="Î¢ÈíÑÅºÚ"><img border="0" src="images/about1.jpg" width="19" height="27"><span class="Apple-converted-space">&nbsp;</span><a href="zjsp.asp?type=65&type2=31" style="text-decoration: none">´´ÐÂ³É¹û</a></font></TD>
                                      
                                    </TR>
                                  </TBODY>
                                </TABLE>
										</TD>
                                      
                                    </tr>

					</table>
				<table border="0" width="95%" height="142" cellspacing="0" cellpadding="0">
					<tr>
						<td bgcolor="#FFFFFF" height="142"> 
              	
                <table border="0" width="260" cellspacing="0" height="436"   style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 0px solid;BORDER-right: #808080 0px solid;BORDER-bottom: #808080 0px solid;">
					<tr>
						<td bgcolor="#FFFFFF" height="260"> 
						<font color="#666666">
						<table border="0" width="302" cellspacing="0" cellpadding="0" height="40" bgcolor="#004993">
							<tr>
								<td height="50" bgcolor="#0099FF">
								<p class="anli" style="margin-top: 0; margin-bottom: 0" align="left">
										<font color="#006699" face="Î¢ÈíÑÅºÚ">¡¡</font><span style="font-weight: 700"><font color="#FFFFFF" face="Î¢ÈíÑÅºÚ">ÍÅ¶Ó¶¯Ì¬</font></span></td>
							</tr>
							<tr>
								<td height="10" bgcolor="#FFFFFF">
								¡¡</td>
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
<font color="#000F1A" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" face="Î¢ÈíÑÅºÚ">
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
						<font face="Î¢ÈíÑÅºÚ" style="font-size: 15px; border: 0px none; margin: 0px; padding: 0px">
						<font color="#000F1A" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">¡¤  
						</font>  
						<A 
                              title=<%=oRS("title")%> 
                              href="newsview.asp?id=<%=oRS("id")%>" 
                               style="border:0px none; margin:0px; padding:0px; text-decoration: none; font-size:15px">
						<font color="#000F1A" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px">
						<%=left(oRS("title"),15)%></font></A></font></TD>
						<TD class=sd13 height=27 width="56">
						<p align="center" style="margin-top: 0; margin-bottom: 0">
						<font color="#000F1A" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" face="Î¢ÈíÑÅºÚ">(<% =month(ors("infotime")) %>.<% =day(ors("infotime")) %>)</font></TD>
                    </TR>

                  </TBODY>
                </TABLE>
                <font color="#FFFF00" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" face="Î¢ÈíÑÅºÚ">
                <b style="border: 0px none; margin: 0px; padding: 0px; font-size:15px"><%
oRS.MoveNEXT
end if
next
else Response.Write ("ÔÝÎÞÄÚÈÝ.") end if
oRS.close 
Set oRS = Nothing
%>
              </b>
              </font></ul> 
<div id="demo2" style="height: auto; text-align: left; border: 0px none; margin: 0px; padding: 0px; font-size:15px">
	<p style="margin-top: 0; margin-bottom: 0"></div>
</div> 
							<font color="#FFFF00" style="border: 0px none; margin: 0px; padding: 0px; font-size:15px" face="Î¢ÈíÑÅºÚ"> 
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
                ¡¡</td>
					</tr>

					<tr>
						<td bgcolor="#FFFFFF" width="248" height="130" align="left"> 
              	¡¡</td>
					</tr>

					</table>
              					</td>
							</tr>
						</table></td>
					</tr>
					<tr>
						<td bgcolor="#FFFFFF"> 
										¡¡</td>
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
						<font color="#669900">¡¡<img border="0" src="fl/jiantou.jpg" width="17" height="14"> </font></SPAN>
						<font face="Î¢ÈíÑÅºÚ"><a href="http://www.dfmzsz.com" style="text-decoration: none">
						<font color="#404A53">Ê×Ò³</font></a><font color="#404A53"> 
                        &gt;&gt; <%= rs("ProductBigClass") %> <%if rs("ProductSmallClass")<>"" then%>
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
		<img border="0" src="fl/fot02.jpg" width="820" height="24">¡¡</font></DIV></TD>
                                          </TR>

  <TR bgcolor="#FFFFFF"> 
                                            <TD height=32 bgcolor="#FFFFFF" width="13%"> 
											<DIV align=right>
												<p style="line-height: 100%; margin-top: 0; margin-bottom: 0">
												<font color="#4D4D4D" face="Î¢ÈíÑÅºÚ">
												<strong>·¢²¼Ê±¼ä£º</strong></font></DIV></TD>
                                            <TD height=32 width="16%"bgcolor="#FFFFFF" align="left"> <% =RS("ProductTime") %></TD>
                                            <TD height=32 bgcolor="#FFFFFF"> 
											<DIV align=right>
												<p style="line-height: 100%; margin-top: 0; margin-bottom: 0">
												<font color="#4D4D4D" face="Î¢ÈíÑÅºÚ">
												<strong>
												¹Ø×¢Ö¸Êý£º</strong></font></DIV></TD>
                                            <TD bgcolor="#FFFFFF" align="left">
											<p style="line-height: 100%; margin-top: 0; margin-bottom: 0">
											<font color="#4D4D4D" face="Î¢ÈíÑÅºÚ">&nbsp; <%= rs("hits") %> ´Î</font></TD>
											
                                          </TR>

                                        
                                        </TBODY>
                                      </TABLE>
									<p style="line-height: 100%; margin-top: 0; margin-bottom: 0">
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
                            <TD height=22 align="left">	<div align="center">
								<table border="0" width="100%" height="32">
									<tr>
										<td height="30">
										<div align="center">
                      <TABLE width=796 height=10 border=0 cellPadding=0 cellSpacing=1 bgcolor="#FFFFFF">
                        <TBODY>
                        <TR vAlign=top> 
                            <TD bgcolor="#FFFFFF" align="left"><p style="text-align: center"><object height="460" width="605" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" border="1">
<param name="movie" value="Flvplayer.swf" />
<param name="quality" value="High" />
<param name="allowFullScreen" value="true" />
<param name="FlashVars" value="vcastr_file=<% =left(RS("ProductPrice4"),220) %>&amp;BufferTime=2&amp;isautoplay=1&amp;iscontinue=1" />
<param name="bgcolor" value="#ffffff" /><embed height="460" width="605" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" quality="High" flashvars="vcastr_file=<% =left(RS("ProductPrice4"),220) %>&amp;BufferTime=2&amp;isautoplay=1&amp;iscontinue=1" allowfullscreen="true" src="Flvplayer.swf" bgcolor="#ffffff"></embed></object></p>
							¡¡</TD>
                          </TR>
                          <TR vAlign=top> 
                            <TD 
                              height=1 bgcolor="#FFFFFF" align="left">
							<font color="#4D4D4D" face="Î¢ÈíÑÅºÚ"><%=rs("ProductContent") %></font></TD>
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
                                                                      
                              <font face="Î¢ÈíÑÅºÚ">
                                                                      
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