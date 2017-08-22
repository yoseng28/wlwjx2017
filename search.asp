<!--#include file="conn.asp"-->
<%
key=request("key")
leixing=request("leixing")
key=replace(key,"'","")
leixing=replace(leixing,"'","")
key=left(key,12)
leixing=left(leixing,12)
if key="" then
   response.write "<script>alert('查找字符串不能为空！');history.back();</Script>"
   response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../../../../../../images/style.css" type=text/css rel=stylesheet>
<title><% =sitetitle %></title>
<meta name="Keywords" content="<% =webKeywords %>">
<meta name="Description" content="<% =webDescription %>">
<meta name="robots" content="index, follow">
<meta name="googlebot" content="index, follow">
</head>
<BODY bgcolor="#214984" topMargin=0 leftmargin="0">
<!--#include file="top.asp"-->
<TABLE cellSpacing=0 cellPadding=0 width=980 align=center bgColor=#ffffff 
border=0>
  <TBODY>
    <TR> 
      <TD width="5" bgcolor="#FFFFFF"></TD>
      <TD width="960"><TABLE width=960 border=0 align=center cellPadding=0 cellSpacing=0>
          <TBODY>
            <TR> 
              <TD width="235" vAlign=top background="../../../../../../../images/class_bg4.gif" style="background-repeat: repeat-x"
 bgcolor="#FFFFFF"> <TABLE style="MARGIN: 10px 0px 5px" cellSpacing=0 
                  cellPadding=0 width=100% border=0>
                  <TBODY>
                    <TR> 
                      <TD width="6%" height="25">　</TD>
                      <TD width="94%">
						<IMG height=16 
                        src="../../../../../../../images/class_left0.gif" width=11 
                        align=absMiddle><STRONG><font color="FF9900">关于我们</font></STRONG> 
                        <STRONG>About Us</STRONG></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE width=90% border=0 align="center" cellPadding=5 cellSpacing=1 
                  bgColor=#dddddd>
                  <TBODY>
                    <TR> 
                      <TD width=186 bgColor=#ffffff> <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
                          <TBODY>
                            <TR> 
                              <TD height=8></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <%
sql2="select *  from about order by id asc "
Set oRS= Server.CreateObject("ADODB.recordset")
oRS.Open sql2,conn,1,3
 if NOT oRS.EOF then 
for i=1 to 10
  if NOT oRS.EOF then
   %>
                        <TABLE width=100% 
              border=0 align=center cellPadding=0 cellSpacing=0>
                          <TBODY>
                            <TR> 
                              <TD width="12%" height="25"> <div align="center"><font color="#0066CC">·</font></div></TD>
                              <TD width="88%"><a href="about.asp?type=<%=oRS("id")%>"><SPAN style="FONT-SIZE: 14px"><font color="#0066CC"><%=oRS("title")%></font></SPAN></a></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <%
oRS.MoveNEXT
end if
next
else Response.Write ("暂无内容.") end if
oRS.close 
Set oRS = Nothing

%>
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
                </TABLE>
                <TABLE style="MARGIN: 10px 0px 5px" cellSpacing=0 
                  cellPadding=0 width=100% border=0>
                  <TBODY>
                    <TR> 
                      <TD width="12" height="25">　</TD>
                      <TD width="186">
						<IMG height=16 
                        src="../../../../../../../images/class_left0.gif" width=11 
                        align=absMiddle><STRONG><font color="FF9900">站内搜索</font></STRONG> 
                        <STRONG>Search</STRONG></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE width=90% border=0 align="center" cellPadding=5 cellSpacing=1 
                  bgColor=#dddddd>
                  <TBODY>
                    <TR> 
                      <TD width=186 bgColor=#ffffff> <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
                          <TBODY>
                            <TR> 
                              <TD height=8></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <TABLE width=100% border=0 align="center" cellPadding=2 cellSpacing=0>
                          <SCRIPT language=JavaScript>
<!--
function postSearch(theform) {
	if(theform.key.value=="") {
		alert("请输入搜索关键字！");
		theform.key.focus();
		return false;
	}
	if(theform.leixing.value=="1") {
		alert("请选择搜索类型！");
		theform.leixing.focus();
		return false;
	}
}
-->
</SCRIPT>
                          <form name="form2" method="post" onsubmit="javascript:return postSearch(this)" action="../../../../../../../search.asp">
                            <TBODY>
                              <TR> 
                                <TD width="60" rowspan="2" align=center>
								<img src="../../../../../../../images/class_search.gif" width="42" height="40"></TD>
                                <TD width="155" height="25" align=left><font color="#666666">&nbsp;请输入关键字：</font> 
                                </TD>
                              </TR>
                              <TR> 
                                <TD height="25" align=middle><INPUT class=box03 id=key2 size=15 
                        name=key> <input name="leixing" type="hidden" id="leixing" value="ProductName"></TD>
                              </TR>
                              <TR> 
                                <TD height="25" align=middle>　</TD>
                                <TD align=middle>
								<INPUT id=ImageButton1 type=image alt="" 
                        src="../../../../../../../../images/go.gif" border=0 
                      name=ImageButton12></TD>
                              </TR>
                            </TBODY>
                          </form>
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
                </TABLE>
                <TABLE cellSpacing=0 cellPadding=0 width=100 align=center 
border=0>
                  <TBODY>
                    <TR> 
                      <TD height=8></TD>
                    </TR>
                  </TBODY>
                </TABLE></TD>
              <TD width="1">&nbsp;</TD>
              <TD width="734" vAlign=top style="BORDER-TOP: #C5CED7 1px solid;BORDER-LEFT: #C5CED7 1px solid;BORDER-right: #C5CED7 1px solid;BORDER-bottom: #C5CED7 1px solid;"> 
                <TABLE cellSpacing=0 cellPadding=0 width="724" 
border=0>
                  <TBODY>
                    <TR> 
                      <TD width="10" height="29" align="center">　</TD>
                      <TD width="296"> <font color="FF9900">内容搜索</font> </TD>
                      <TD width="404" align="right"><font color="#FF0000">⊙ </font>
						<a href="../../index.asp">首页</a> 
                        &gt;&gt; 内容搜索 </TD>
                      <TD width="14">　</TD>
                    </TR>
                    <TR> 
                      <TD height="5" colspan="4" align="center" background="../../../../../../../images/lanmufeng.gif"></TD>
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
                <% 
on error resume next
Set rs= Server.CreateObject("ADODB.Recordset")
if leixing="ProductName" then
sql="select * from ProductContent where ProductName Like '%"& key &"%' order by id desc"
elseif leixing="ProductContent" then
sql="select * from ProductContent where ProductContent Like '%"& key &"%' order by id desc"
else
end if
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<br><br><br><p align='center'>对不起，没有找到相关内容，请尝试用其他关键字试试</p>"
else
%>
                <table width="100%" border="0">
                  <tr> 
                    <td height="25"><div align="center"><SPAN style="FONT-SIZE: 14px">您搜索的关键字市　<font color="#FF0000"><strong><%=key%></strong></font>，共为您找到　<font color="#FF0000"><%=rs.recordcount%>　</font>条内容 
                        </SPAN> </div></td>
                  </tr>
                </table>
                <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
                  <tr bgcolor="#e3f8fb"> 
                    <td width="15%" height="28" align="center"> <div align="center"><span class="style7">栏 
                        目</span></div></td>
                    <td width="63%" height="28" align="center"> <div align="center"><span class="style7">标 
                        题</span></div></td>
                    <td width="22%" align="center"> <div align="center"><span class="style7">时间</span></div></td>
                  </tr>
                  <%
i=0
do while not rs.eof
%>
                  <tr bgcolor="#FFFFFF"> 
                    <td style="BORDER-BOTTOM: #c0c0c0 1px dotted" height="25" align="center"><a href="class.asp?type=<%=rs("ProductBigClass")%>"><%=rs("ProductBigClass")%></a></td>
                    <td style="BORDER-BOTTOM: #c0c0c0 1px dotted"><a href="productview.asp?id=<%=rs("id")%>"  target="_blank"><%=replace(rs("ProductName"),key,"<font color=#FF0000><strong>"&key&"</strong></font>")%></a></td>
                    <td style="BORDER-BOTTOM: #c0c0c0 1px dotted" align="center"> 
                      <div align="left"><%=rs("infotime")%> </div></td>
                  </tr>
                  <%
rs.movenext
i=i+1                                                         
loop
%>
                  <tr bgcolor="#FFFFFF"> 
                    <td height="24" colspan="3"><div align="center"></div></td>
                  </tr>
                  <% 
end if
rs.close
set rs=nothing
%>
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
