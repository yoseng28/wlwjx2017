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
'��ȡ��Ʒ��ʾ��ʽ
sqlxs="select *  from BigClass where ClassType='product.asp' and BigClassName='"&bigclass&"' order by BigClassID desc "
Set oRSxs= Server.CreateObject("ADODB.recordset")
oRSxs.Open sqlxs,conn,1,3
 if NOT oRSxs.EOF then 
xsfs=oRSxs("xsfs")
else
		response.write "<br><p align=center><font color='red'>�Բ�������ʵ��ļ������ڡ�</font></p>"
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
<style fprolloverstyle>A:hover {text-decoration: none; color: #CC0000; font-weight: bold}
</style>
<style>
<!--
a:link {text-decoration: none;  }
.text        { font-family: ΢���ź�; font-size: 12pt }
.anli        { font-family: ΢���ź�; font-size: 14pt; font-weight: bold }

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
              <TD width="252" vAlign=top bgcolor="#ffffff"> 
               
                <table border="0" width="240" cellspacing="0" height="32">
					<tr>
                      <TD height="32" align="left">
                <table border="0" width="240" cellspacing="0" height="644">
					<tr>
                      <TD height="60" align="center" background="fl/fanwei.gif">
						<p class="anli" align="left"><font color="#4D4D4D">����Ʒ����</font></TD>
                    </tr>
                                   

					<tr>
						<td bgcolor="#FFFFFF"> 
                <table border="0" width="258" cellspacing="0" height="262">
                <tr>
                                      <TD align="left" height="262" class="text">
                <table border="0" width="258" cellspacing="0" height="247">
                <tr>
                                      <TD align="left" height="37" background="bj.jpg" class="text">
										<p style="margin-top: 0; margin-bottom: 0">
										<font face="΢���ź�">&nbsp;<img border="0" src="images/about1.jpg" width="19" height="27"> </font><a style="text-decoration: none" href="product.asp?type=31&type2=30">
										<font face="΢���ź�">���ǽ�綯��Ͱ</font></a></TD>
                                      
                                    </tr>

					<tr>
                                      <TD align="left" height="42" background="bj.jpg" class="text">
										<p style="margin-top: 0; margin-bottom: 0">
										<font face="΢���ź�">&nbsp;<img border="0" src="images/about1.jpg" width="19" height="27"> </font><a style="text-decoration: none" href="product.asp?type=31&type2=31">
										<font face="΢���ź�">���ǽ���ˮ������</font></a></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="41" background="bj.jpg" class="text">
										<font face="΢���ź�">&nbsp;<img border="0" src="images/about1.jpg" width="19" height="27"> </font>
										<font color="#006699" face="΢���ź�">
										<a style="text-decoration: none" href="product.asp?type=31&type2=39">
										��Դ��������ԡ��</a></font></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="43" background="bj.jpg" class="text">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="΢���ź�">&nbsp;<img border="0" src="images/about1.jpg" width="19" height="27"> </font>
										<font color="#CDCED0" face="΢���ź�">
										<a style="text-decoration: none" href="product.asp?type=31&type2=45">
										���ض�����ʽˮ���ʽ��Ͱ</a></font></TD>
                                      
                                    </tr>
                                    	<tr>
                                      <TD align="left" height="43" background="bj.jpg" class="text">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="΢���ź�">&nbsp;<img border="0" src="images/about1.jpg" width="19" height="27"> </font>
										<font color="#CDCED0" face="΢���ź�">
										<a style="text-decoration: none" href="product.asp?type=31&type2=83">
										���ض���Ͱ</a></font></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="41" background="bj.jpg" class="text">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="΢���ź�">&nbsp;<img border="0" src="images/about1.jpg" width="19" height="27"> </font>
										<a href="product.asp?type=31&type2=46" style="text-decoration: none">
										<font face="΢���ź�">
										ϴ���裬ˮ��ͷ</font></a></TD>
                                      
                                    </tr>
					<tr>
                                      <TD align="left" height="43" background="bj.jpg" class="text">
										<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
										<font face="΢���ź�">&nbsp;<img border="0" src="images/about1.jpg" width="19" height="27"> </font>
										<font face="΢���ź�" color="#4D4D4D">
										<a href="product.asp?type=31&type2=72" style="text-decoration: none">��ԡ��</a></font></TD>
                                      
                                    </tr>
					
					</table>
										</TD>
                                      
                                    </tr>

					</table>
				<p style="margin-top: -3px; margin-bottom: -3px" align="center">
				��<table border="0" width="95%" height="142" cellspacing="0" cellpadding="0">
					<tr>
						<td bgcolor="#FFFFFF" height="142"> 
              	
				<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">
				<font color="#404A53">
				<img border="0" src="fl/lianxi.gif" width="259" height="65"></font></p>
				<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">
				<font face="΢���ź�" color="#666666" >
				��ַ��<% =address %><br>
				�ֻ���<% =phone  %></font></p>
				<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">
				<font face="΢���ź�" color="#666666">�绰��</font><font face="΢���ź�" color="#404A53"><% =fax %></font><p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
							<font color="#404A53">����</font><font face="΢���ź�" color="#404A53">ҵ����ѯ1��</font><font face="΢���ź�"><a href="http://wpa.qq.com/msgrd?v=1&uin=<% =qq1 %>&site=<% =sitetitle %>&menu=yes" target="blank"><font color="#404A53"><img src="http://wpa.qq.com/pa?p=1:<% =qq1 %>:1" border="0"></font></a></font></p>
							<p style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="΢���ź�">
							<font color="#404A53">
							����ҵ����ѯ2��</font><a href="http://wpa.qq.com/msgrd?v=1&uin=<% =qq2 %>&site=<% =sitetitle %>&menu=yes" target="blank"><font color="#404A53"><img src="http://wpa.qq.com/pa?p=1:<% =qq2 %>:1" border="0"></font></a></font></p>
							<p style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="΢���ź�">
							<font color="#404A53">
							����ҵ����ѯ3��</font><b><a href="http://wpa.qq.com/msgrd?v=1&uin=<% =qq3 %>&site=<% =sitetitle %>&menu=yes" target="blank"><font color="#404A53"><img src="http://wpa.qq.com/pa?p=1:<% =qq2 %>:1" border="0"></font></a></b></font><p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
				<img border="0" src="fl/2wm.jpg" width="265" height="148"><p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
				��<p style="line-height: 200%; margin-top: 0; margin-bottom: 0"> 
              	��</td>
					</tr>
					</table>
										</td>
					</tr>
				</table>
                		</TD>
                    </tr>
					</table>
                
                </TD>
            
              <TD width="847"  vAlign=top style="BORDER-TOP: #808080 0px solid;BORDER-LEFT: #808080 1px solid;BORDER-right: #808080 0px solid;BORDER-bottom: #808080 0px solid;"> 
				<TABLE cellSpacing=0 cellPadding=0 width="845" 
border=0 height="61" bgcolor="#FFFFFF">
                  <TBODY>
                    <TR> 
                      <TD width="845" height="55" align="left" style="BORDER-BOTTOM: #CFCFCF 0px dashed" background="fl/tibj.gif"> 
                        
                      <p class="anli" style="margin-top: 0; margin-bottom: 0"> 
                        
                      <img border="0" src="fl/jiantou.jpg" width="17" height="14"><font face="΢���ź�"><a href="index.asp" style="text-decoration: none"><SPAN style="FONT-WEIGHT: bold;"><font color="#669900">
						</font><font color="#404A53">
						��ҳ</font></SPAN></a><font color="#404A53"><b> 
                        &gt;&gt; <%=bigclass%><%if smallclass<>"" then%> &gt;&gt; <%=smallclass%><%end if%></b></font></font></TD>
                    </TR>
                     <TR> 
                      <TD height="6" align="center" background="../images/lanmufeng.gif"></TD>
                    </TR>

                  </TBODY>
                </TABLE>
<%if xsfs="ͼ���б�" then%>
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
response.Write("��ʱû������")
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
                                  <td>��</td>
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
                                            <TD width="17%" height=21> <DIV align=center><strong>��Ʒ���ƣ�</strong></DIV></TD>
                                            <TD width="83%">&nbsp;<STRONG> <a href="productview.asp?id=<%= RS("id") %>" target="_blank"> 
                                              <% =RS("ProductName") %>
                                              </a></STRONG></TD>
                                          </TR>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD height=21> <DIV align=center><strong>��Ʒ���</strong></DIV></TD>
                                            <TD>&nbsp; 
                                              <% =RS("ProductType") %>
                                            </TD>
                                          </TR>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD height=21> <DIV align=center><strong>��Ʒ�۸�</strong></DIV></TD>
                                            <TD>&nbsp; 
                                              <% =RS("ProductPrice") %>
                                            </TD>
                                          </TR>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD height=21> <DIV align=center><strong>��עָ����</strong></DIV></TD>
                                            <TD>&nbsp; �����<%= rs("hits") %> ��</TD>
                                          </TR>
                                          <TR bgcolor="#FFFFFF"> 
                                            <TD height=21> <DIV align=center><strong>��Ʒ���ܣ�</strong></DIV></TD>
                                            <TD bgcolor="#FFFFFF">&nbsp;<strong><a href="productview.asp?id=<%= RS("id") %>" target="_blank"><font color="#FF0000">����鿴&gt;&gt;</font></a></strong></TD>
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
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=product.asp?type="&bigclassid&"&type2="&smallclassid&"&page=1><font color=red>��ҳ</font></a>&nbsp;"
    response.write "<a href=product.asp?type="&bigclassid&"&type2="&smallclass&"&page=" & Page-1 & "><font color=red>��һҳ</font></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=product.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & (page+1) & ">"
    response.write "<font color=red>��һҳ</font></a> <a href=product.asp?type="&bigclassid&"&type2="&smallclassid&"&page="&rs.pagecount&"><font color=red>βҳ</font></a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#FF0000'>"&rs.recordcount&"</font></b>�� <b>"&rs.pagesize&"</b>��/ҳ"
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
response.Write("��ʱû������")
else 
%>
                  <% 
rs.PageSize=28
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
							   for j=1 to 6 
							   if NOT RS.EOF then
							   %>
                                <tr>
                                  <% 
							   for i=1 to 3 
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
											<TABLE width="182" height="149" border=1 
                                cellPadding=3 cellSpacing=0 borderColor=#b2b2b2>
                                              <TBODY>
                                                <TR> 
                                                  <TD width="240" height="200" bgColor=#ffffff>
													<p style="line-height: 200%; margin-top: 0; margin-bottom: 0" align="right">
													<a href="productview.asp?id=<%= RS("id") %>">
													<img src="../images/wupic.jpg" width="245" height="185" border="0" style="border: 1px solid #808080; padding: 3px"></a></TD>
                                                </TR>
                                              </TBODY>
                                            </TABLE>
                                            <span class="bh"> 
                                            <% else %>
                                            </span> 
											<TABLE width="191" height="189" border=1 
                                cellPadding=0 cellSpacing=0 borderColor=#b2b2b2>
                                              <TBODY>
                                                <TR> 
                                                  <TD width="187" height="187" bgColor=#ffffff>
													<p align="center" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><a href="productview.asp?id=<%= RS("id") %>" >
													<img src="<% Response.Write (""&RS("ProductImages")&"")%>" width="249" height="188" border="0" style="border: 1px solid #808080; padding: 3px"></a></TD>
                                                </TR>
                                              </TBODY>
                                            </TABLE>
                                            <span class="bh"> 
                                            <% end if %>
                                            </span></TD>
                                        </TR>
                                        <TR> 
                                          <TD height="28" align="center"><STRONG>
                                            <font face="΢���ź�"> 
                                            <a style="text-decoration: none" href="productview.asp?id=<%= RS("id") %>"> 
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
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=product.asp?type="&bigclassid&"&type2="&smallclassid&"&page=1><font color=red>��ҳ</font></a>&nbsp;"
    response.write "<a href=product.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & Page-1 & "><font color=red>��һҳ</font></a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=product.asp?type="&bigclassid&"&type2="&smallclassid&"&page=" & (page+1) & ">"
    response.write "<font color=red>��һҳ</font></a> <a href=product.asp?type="&bigclassid&"&type2="&smallclassid&"&page="&rs.pagecount&"><font color=red>βҳ</font></a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#FF0000'>"&rs.recordcount&"</font></b>�� <b>"&rs.pagesize&"</b>��/ҳ"
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