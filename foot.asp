
<head>
<style>
<!--
.foot        { font-size: 14px }
.sg1          { font-size: 12pt; font-family: ΢���ź�; font-weight: bold }
-->
</style>
</head>

<div align="center">

                <table border="0" width="1105" height="72" cellspacing="0" cellpadding="0">
					<tr>
						<td bgcolor="#FFFFFF" height="49" background="../��ѧ������������/fl/cpzx.gif"> 
              	<p class="sg1" style="margin-top: 0; margin-bottom: 0" align="left">&nbsp;<font color="#3399FF">�������</font><p class="sg1" style="margin-top: 0; margin-bottom: 0" align="left">
		<img border="0" src="fl/fot02.jpg" width="100%" height="14"></td>
					</tr>
					<tr>
                      <TD> 
                <p align="center" style="margin-top: 0; margin-bottom: 0"> 
                
                        <%
sql2="select *  from frendlink where reply='-1' "
Set oRS= Server.CreateObject("ADODB.recordset")
oRS.Open sql2,conn,1,3
 if NOT oRS.EOF then 
%>
                        <TABLE border="0" align="left" cellpadding="0" cellSpacing=0>
                          <%
 do while not oRS.eof
  if NOT oRS.EOF then 
   %>
                          <tr> 
                            <% for m=1 to 6
  if NOT oRS.EOF then
%>
                            <td valign="middle"> <div align="center"> 
                                <TABLE border=0 align="center" cellPadding=0 cellSpacing=0 width="188">
                                  <TBODY>
                                    <TR> 
                                      <TD width="188" height=32 align="center">
										<p align="center" style="margin-top: 0; margin-bottom: 0">
										<font color="#FF3300">&nbsp;</font><a href="<% =ors("weburl")  %>" target="_blank" style="text-decoration: none""><font color="#FF3300"><%=left(oRS("webname"),38)%></font></a><font color="#FF3300">
										</font> 
                                      </TD>
                                    </TR>
                                  </TBODY>
                                </TABLE>
                              </div></td>
                            <% oRS.MoveNEXT
end if
next %>
                          </tr>
                          <%
end if
loop
%>
                        </table>
                        <%
else Response.Write ("��������.") end if
oRS.close 
Set oRS = Nothing
%>
                        </TD>
                    </tr>
				</table>

<TABLE width="100%" height="95" border=0 cellPadding=0 cellSpacing=0 bgcolor="#008000">
  <TBODY>
    <TR> 
      <TD align=middle height=95 bgcolor="#3399FF"> 
        <TABLE width="100%" height="131" border=0 cellPadding=0 cellSpacing=0 bgcolor="#484848" >
          <TBODY>
            <TR> 
              <TD align=center height=16 bgcolor="#FFFFFF">
		<img border="0" src="fl/fot02.jpg" width="100%" height="14"></TD>
            </TR>
            <TR> 
              <TD align=center height=2></TD>
            </TR>
            <tr>
              <TD align=center width="100%" bgcolor="#3399FF">
				<table border="0" width="100%" height="96">
					<tr>
              <TD align=center height=25 class="foot" > 
				<font face="΢���ź�" color="#DDEEFF"> 
				<p style="margin-top: 0; margin-bottom: 0; "> 
				<span style="font-weight: 400">���ƣ�<% =company %>����ַ: <% =address %>�����ʱࣺ<% =code %> </span> </font> 
				<font color="#DDEEFF"> </font></font></TD>
            		</tr>
					<tr>
              <TD align=center height=25 class="foot"> 
				<p style="margin-top: 0; margin-bottom: 0; "> 
				<font color="#FFFFFF"><span style="font-weight: 400">
				<a style="text-decoration: none" >�绰</a></span></font><font face="΢���ź�" color="#DDEEFF"><a style="text-decoration: none" href="tel:09312135926"><span style="font-weight: 400"><font color="#FFFFFF">: 
                <% =fax %>&nbsp;&nbsp;
                &nbsp;</font></span></a></font><span style="font-weight: 400"><font face="΢���ź�" color="#FFFFFF">�ֻ���</font><font face="΢���ź�" color="#DDEEFF"><font color="#DDEEFF"><a href="tel:13609316443" style="text-decoration: none"><font color="#FFFFFF"><% =phone %>����</font></a></font><font face="΢���ź�" color="#FFFFFF">���䣺<% =email %>&nbsp;&nbsp; 
               ��</font></font></span></TD>
            		</tr>
					<tr>
              <TD align=center height=25 class="foot"> 
				<p style="margin-top: 0; margin-bottom: 0"> 
				<span style="font-weight: 400"> 
				<font color="#DDEEFF" face="΢���ź�">
				��վ��Ȩ��</font></span><span style="color: rgb(221, 238, 255); font-family: ΢���ź�; font-size: 14px; font-style: normal; font-variant: normal; font-weight: normal; letter-spacing: normal; line-height: 18.2px; orphans: auto; text-align: -webkit-center; text-indent: 0px; text-transform: none; white-space: normal; widows: 1; word-spacing: 0px; -webkit-text-stroke-width: 0px; display: inline !important; float: none; background-color: rgb(51, 153, 255);">Ӣ������ν�ѧ�Ŷ�</span><span style="font-weight: 400"><font color="#DDEEFF" face="΢���ź�">&nbsp;&nbsp;&nbsp;&nbsp; ��վ������</font><font color="#484848" face="΢���ź�"><a target="_blank" href="http://www.lzkeyuan.cn" style="text-decoration: none"><font color="#DDEEFF">��Դ����</font></a><font color="#DDEEFF">&nbsp;&nbsp;&nbsp;</font></font></span></TD>
            		</tr>
					</table>
				
            </TR>
          		</tr>
            <TR> 
              <TD align=center width="100%" bgcolor="#484848">
				��</TR>
          </TBODY>
        </TABLE></TD>
    </TR>
  </TBODY>
</TABLE>
</div>
