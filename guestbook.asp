<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">

<!--#include file="conn.asp"-->
<% 
bigclass=request("type")
smallclass=request("smallclass")
 %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% =sitetitle %></title>
<script language="javascript">
<!--
function check(){
if(document.form1.name.value.length<1){
  	    alert("����д���⣡");
		  document.form1.name.focus();
		return false;
		}else if(document.form1.homepage.value.length<1){
		alert("����д����������");
		  document.form1.homepage.focus();
		  return false;
		}else if(document.form1.oicq.value.length<1){
		alert("����д������ϵ�绰��");
		  document.form1.oicq.focus();
  ������return false;
		}else if(document.form1.form01.value.length<1){
		alert("����д������ϵ��ϵ��ַ��");
		  document.form1.form01.focus();

		return false;
		} else if(document.form1.TxbContent.value.length<1){
		alert("����д������Ϣ���ݣ�");
		  document.form1.TxbContent.focus();
		return false;
	} else{
     return true;
   }
}
-->
</script>  
<style>
<!--
.itable {
	background-position: top;
	table-layout:fixed;
	word-wrap:break-word;
	word-break:break-all;
}
.liuyan      { font-family: ΢���ź�; font-size: 11pt; line-height: 200%; margin: 0 }
TD {
	COLOR: #000000
}
.distance {
	LINE-HEIGHT: 20px
}
.style14 {
	COLOR: #378cb5
}
td{font-family:arial}.dhl         { font-size: 11pt; font-family: ����; color:#0066CC; line-height:100%; margin-top:0; margin-bottom:0 }
.hl          { line-height: 200%; font-family: ΢���ź�; margin-top: 0; margin-bottom: 0 }
.text        { font-family: ΢���ź�; font-size: 12pt }
-->
</style>
<style fprolloverstyle>A:hover {color: red; font-weight: bold}
</style>
</head>
<BODY topMargin=0 leftmargin="0" link="#484848" vlink="#484848" alink="#484848">
<!--#include file="top.asp"-->
<div align="center">
<TABLE cellSpacing=0 cellPadding=0 width=1100 bgColor=#ffffff   border=0 bordercolor="#7F7F7F"  vAlign=top style="BORDER-TOP: #C5CED7 1px solid;BORDER-LEFT: #C5CED7 1px solid;BORDER-right: #C5CED7 1px solid;BORDER-bottom: #C5CED7 1px solid;"
border=0>
  <TBODY>
    <TR> 
      
      <TD width="996">
		<div align="center">
			<TABLE cellSpacing=0 cellPadding=0 width=1095  border=0>
          <TBODY>
            <TR> 
              <TD width="250" vAlign=top bgcolor="#ffffff" style="BORDER-TOP: #C5CED7 0px solid;BORDER-LEFT: #C5CED7 0px solid;BORDER-right: #C5CED7 1px solid;BORDER-bottom: #C5CED7 0px solid;"
border=0>
                <table border="0" width="100%" cellspacing="0" bordercolor="#993300" cellpadding="0">
					<tr>
                      <TD height="65" bgcolor="#FFFFFF">
						<img border="0" src="fl/about.gif" width="245" height="65"></TD>
                    </tr>
					<tr>
						<td bgcolor="#FFFFFF"> 
                                <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                              border=0 height="153" >
                                  <TBODY>
                                    <TR> 
                                      <TD align="left" height="40" background="images/bj.jpg" class="text">
										<font face="΢���ź�">
										��<img border="0" src="images/about1.jpg" width="19" height="18">��<a href="about.asp?type=29" style="text-decoration: none">��ͼ��</a></font></TD>
                                      
                                    </TR>
                                    <TR> 
                                      <TD align="left" height="40" background="images/bj.jpg" class="text">
										<font face="΢���ź�">
										��<img border="0" src="images/about1.jpg" width="19" height="18">��<a href="about.asp?type=95" style="text-decoration: none">��;���</a></font></TD>
                                      
                                    </TR>
                                    <TR> 
                                      <TD align="left" height="40" background="images/bj.jpg" class="text">
										<font face="΢���ź�">
										��<img border="0" src="images/about1.jpg" width="19" height="18">��<a href="about.asp?type=5" style="text-decoration: none">��ϵ����</a></font></TD>
                                      
                                    </TR>
                                    <TR> 
                                      <TD align="left" height="40" background="images/bj.jpg" class="text">
										<font face="΢���ź�">
										��<img border="0" src="images/about1.jpg" width="19" height="18">��<a href="guestbook.asp" style="text-decoration: none">���߷���</a></font></TD>
                                      
                                    </TR>
                              
                                  </TBODY>
                                </TABLE>
								</td>
					</tr>
					<tr>
                      <TD height="40" align="center" style="font-size: 10pt; line-height: 22px; font-family: arial" width="205">
                <table border="0" width="95%" height="240" cellspacing="0" cellpadding="0">
					<tr>
						<td bgcolor="#FFFFFF" height="240" style="font-size: 10pt; line-height: 22px; font-family: arial"> 
              	<p align="center">
				<img border="0" src="fl/lianxi.gif" width="245" height="65"><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">
				<font face="΢���ź�" color="#404A53">
				��ַ��<% =address %><br>
				�绰��<% =fax %></font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">
				<font face="΢���ź�" color="#404A53">�ֻ���<% =beianhao %></font><p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
							<font color="#404A53">����</font><font face="΢���ź�" color="#404A53">ҵ����ѯ1��</font><font face="΢���ź�"><a href="http://wpa.qq.com/msgrd?v=1&uin=<% =qq1 %>&site=<% =sitetitle %>&menu=yes" target="blank"><font color="#404A53"><img src="http://wpa.qq.com/pa?p=1:<% =qq1 %>:1" border="0"></font></a></font></p>
							<p style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="΢���ź�">
							<font color="#404A53">
							����ҵ����ѯ2��</font><a href="http://wpa.qq.com/msgrd?v=1&uin=<% =qq2 %>&site=<% =sitetitle %>&menu=yes" target="blank"><font color="#404A53"><img src="http://wpa.qq.com/pa?p=1:<% =qq2 %>:1" border="0"></font></a></font></p>
							<p style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="΢���ź�">
							<font color="#404A53">
							����ҵ����ѯ3��</font><b><a href="http://wpa.qq.com/msgrd?v=1&uin=<% =qq3 %>&site=<% =sitetitle %>&menu=yes" target="blank"><font color="#404A53"><img src="http://wpa.qq.com/pa?p=1:<% =qq2 %>:1" border="0"></font></a></b></font><p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
				<img border="0" src="fl/2wm.jpg" width="265" height="148"></td>
					</tr>
					</table>
						</TD>
                    </tr>
					<tr>
						<td bgcolor="#FFFFFF" style="font-size: 10pt; line-height: 22px; font-family: arial" bordercolor="#993300"> 
                                <img border="0" src="fl/ad01.gif" width="250" height="255"></td>
					</tr>
				</table>
				</TD>
              <TD width="4" vAlign=top bgcolor="#800000"></TD>
              <TD width="853" vAlign=top bgcolor="#FFFFFF"> 
                <table border="0" width="100%">
					<tr>
						<td> 
                <TABLE cellSpacing=0 cellPadding=0 width="847" 
border=0 height="59">
                  <TBODY>
                    <TR> 
                      <TD width="10" height="52" align="center">��</TD>
                      <TD width="832"> 
                        
                      <p align="left" class="text"> 
                        
                      <font face="΢���ź�"> 
                        
                      <b><font color="#4D4D4D"> 
                        
                      <img border="0" src="fl/jiantou.jpg" width="17" height="14">
						</font> 
                        
                      <a href="index.asp" style="text-decoration: none">
						<font color="#FF3300">��ҳ</font></a></b><font color="#FF3300"><b> &gt;&gt; 
						���߷���</b></font></font></TD>
                      <TD width="102">��</TD>
                    </TR>
                    <TR> 
                      <TD height="7" colspan="3" align="left" background="images/lanmufeng.gif"></TD>
                    </TR>
                  </TBODY>
                </TABLE>
                		</td>
					</tr>
				</table>
                <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td>
					<form action="guestadd.asp" method="POST" name=form1 target="_parent">
                        <table align=center border=0 cellspacing=0 
                          width="76%">
                          <tbody>
                            <tr> 
                              <td height="30"><div align="center">
								<p class="text"><font color="#666666"><b>
								������д��Ϣ���ύ</b></font></div></td>
                            </tr>
                            <tr> 
                              <td height="25" align="left" class="text">
								<p class="text">
								<font face="΢���ź�" color="#666666"><b>��Ϣ����:</b> 
                                </font>
								<font color="#666666" face="΢���ź�"> 
                                <input 
                              name=name class=bg id=name3 size=20 maxlength="50"> 
                                * </font></td>
                            </tr>
                            <tr> 
                              <td height="25" align="left" class="text">
								<font face="΢���ź�" color="#666666"><b>�ա�����:</b> 
                                </font>
								<font color="#666666" face="΢���ź�"> 
                                <input name="homepage" type="text" class=bg id="homepage3" value="<% =session("netname") %>" size="20" maxlength="50"> 
                                *</font></td>
                            </tr>
                            <tr> 
                              <td height="24" align="left" class="text">
								<font face="΢���ź�" color="#666666"><b>��ϵ�绰: 
                                </b> 
								</font>
								<font color="#666666" face="΢���ź�"> 
                                <input 
                              name=oicq class=bg id=oicq3 size=20 maxlength="15">
								*</font></td>
                            </tr>
                            <tr> 
                              <td height="26" align="left" class="text">
								<font face="΢���ź�" color="#666666"><b>��ϵ��ַ:</b> 
                                </font>
								<font color="#666666" face="΢���ź�"> 
                                <input class=bg name="from01" type="text" id="from012" size="20" maxlength="80"> 
								</font>
								</td>
                            </tr>
                            <tr> 
                              <td height="25" align="left" class="text"> 
								<font face="΢���ź�" color="#666666">&nbsp;&nbsp;&nbsp; <b>&nbsp;&nbsp; �ѣ�: </b> 
                                </font><font color="#666666" face="΢���ź�"> 
                                <input class=bg name="email" type="text" id="email3" size="20" maxlength="50"></font></td>
                            </tr>
                            <tr> 
                              <td height="25" align="left" class="text">
								<b><font color="#666666" face="΢���ź�">����:</font><font color="#FF3300" face="΢���ź�">��</font></b><font color="#FF3300" face="΢���ź�">�ڴ�˵����Ҫ��ѯ����ϸ���ݣ�</font></td>
                            </tr>
                            <tr> 
                              <td valign="top" class="text"> <table class=p11 cellspacing=0 cellpadding=1 
                        width="60%" border=0>
                                  <tbody>
                                    <tr> 
                                      <td valign=top>
										<font face="΢���ź�">
										<textarea id=textarea2 style="WIDTH: 750px; HEIGHT: 176px" name=TxbContent rows="1" cols="20"></textarea> 
                                      	</font> 
                                      </td>
                                    </tr>
                                  </tbody>
                                </table></td>
                            </tr>
                            <tr> 
                              <td height="30" class="text"> <div align="left"> 
                                  <font face="΢���ź�"> 
                                  <input type="submit" name="Submit" value="�ᡡ�����š�Ϣ" onClick="return check();">
                                  ���� 
                                  <input type="reset" name="Submit2" value="�� д">
                                  <input name="classname" type="hidden" id="classname" value="webguest">
                                	</font>
                                </div></td>
                            </tr>
                          </tbody>
                        </table>
                      </FORM></td>
                  </tr>
                </table>
                <div align="center">
                <table width="98%" border="0" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td style="color: #666666; font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11px"> 
                      <font face="΢���ź�"> 
                      <% 
	page = CLng(request("page"))							'����CLng������pageֵת��ΪLong��
	judge=request("judge")
	judge2=request("judge2")
	judge3=0
sql2="select *  from guestbook where reply='-1'and classname='webguest' order by id desc "
Set rs= Server.CreateObject("ADODB.recordset")
rs.Open sql2,conn,1,3
%>
                      <%
			if not (rs.EOF or rs.BOF) then
			rs.MoveFirst
			rs.PageSize = 8
			If page < 1 Then page = 1 
			If page > rs.PageCount Then page = rs.PageCount  
			rs.AbsolutePage = page
			For ipage = 1 To rs.PageSize%>
                      </font>
                      <table width="101%" border="0" align="center" cellspacing="0" class="itable">
                        <tr> 
                          <td width="254" background="images/bg_line.gif" class="liuyan" align="left">
							<strong>
							<font color="#CC0000">���Ա��⣺ </font> </strong>
							<font color="#CC0000"><%=rs("name") %></font></td>
                          <td width="140" background="images/bg_line.gif" class="liuyan" align="left">��</td>
                          <td width="358" background=images/bg_line.gif" class="liuyan" align="left">
							<p align="right"><span class="style14">
							<font color="#C5CED7"><b>�ύʱ�䣺</b><%=rs("time01")%></font></span></td>
                        </tr>
                        <tr> 
                          <td colspan="3" background="images/bg_line.gif" class="liuyan" align="left">
							<span class="style14"><font color="#9751FF">
							<b>�������ݣ�</b> <%=rs("guestcontent")%></font></span></td>
                        </tr>
                        <tr background="images/bg_line.gif"> 
                          <td colspan="3" class="liuyan" align="left">
							<strong>
							<font color="#993300">����Ա�ظ���</font></strong><span class="style14"><font color="#000000"> 
                            <% =rs("replycontent") %>
                            </font></span></td>
                        </tr>
                        <tr> 
                          <td colspan="3" class="distance" height="1" bgcolor="#D9D9FF"></td>
                        </tr>
                      </table>
                      <div align="center"> 
                        <font face="΢���ź�"> 
                        <%   
				rs.MoveNext      
				If rs.EOF Then Exit For   
				Next   
			%>
                      </font>
                      </div>
                      <table width="588" border="0" align="center" cellspacing="0">
                        <tr> 
                          <td width="586" style="color: #666666; font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px" class="hl">
							<font face="΢���ź�">&nbsp;&nbsp;&nbsp; <div align="center">
								<p class="liuyan"><font color="#000000">
								����<%=rs.recordCount%> 
                              	������ ҳ����<%=page%> /&nbsp; <%=rs.PageCount%>&nbsp;&nbsp; 
                              <%
				If page = 1 Then 
				   Response.Write "��һҳ����һҳ" 
				End If
				If page <> 1 Then 
				   Response.Write "<a href=guestbook.asp?page=1 class='yellow'><font color=#0000FF>��һҳ</font></a>" 
				   Response.Write "��<a href=guestbook.asp?page=" & (page - 1) & " class='yellow'><font color=#0000FF>��һҳ</font></a>" 
				End If
				If page <> RS.PageCount Then 
				   Response.Write "��<a href=guestbook.asp?page=" & (page + 1) & " class='yellow'><font color=#0000FF>��һҳ</font></a>" 
				   Response.Write "��<a href=guestbook.asp?page=" & RS.PageCount & " class='yellow'><font color=#0000FF>���һҳ</font></a>" 
				End If 
				If page = RS.PageCount Then 
				   Response.Write "����һҳ�����һҳ" 
				End If
			  %>
                              </font> &nbsp; </font> </div></td>
                        </tr>
                      </table>
                      <font face="΢���ź�">
                      <%else%>
                      </font>
                      <table width="75%" border="0" align="center" cellspacing="0">
                        <tr> 
                          <td style="color: #666666; font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11px">
							<font color="#000000" face="΢���ź�">Ŀǰû���û����ԣ�</font></td>
                        </tr>
                      </table>
                      <font face="΢���ź�">
                      <%end if%>
                    </font>
                    </td>
                  </tr>
                </table>
                </div>
                </TD>
            </TR>
          </TBODY>
        </TABLE></div>
		
    </TR>
  </TBODY>
</TABLE>
</div>
<!--#include file="foot.asp"-->
</body>
</html>
