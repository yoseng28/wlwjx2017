<!--#include file="adminconn.asp"-->
<%
dim urllist,fso,fout
if request("xiugai")="yes" then
'��ȡ�ύ����
company=Server.HTMLEncode(Request("company"))
sitetitle=Server.HTMLEncode(Request("sitetitle"))
webDescription=request.form("webDescription")
webKeywords=request.form("webKeywords")
gongsurl=request.form("gongsurl")
weblogo=request.form("weblogo")
weblogotype=right(weblogo,3)
dataurl=request.form("dataurl")
backurl = right(dataurl,instrrev(dataurl,"/"))
address=Server.HTMLEncode(Request("address"))
code=Server.HTMLEncode(Request("code"))
companycon=request.form("companycon")
phone=Server.HTMLEncode(Request("phone"))
fax=Server.HTMLEncode(Request("fax"))
email=request.form("email")
reply=request.form("reply")
kehu=request.form("kehu")
chengxuup=request.form("chengxuup")
companycontent=Server.HTMLEncode(Request("companycontent"))
companycontent = Replace( companycontent, vbCrLf, "<br>")
companycontent = Replace( companycontent, "'", "''")
qq1=trim(request.form("qq1"))
qq2=trim(request.form("qq2"))
qq3=trim(request.form("qq3"))
beianhao=request.form("beianhao")
ipaddress=request.form("ipaddress")
index1=request.form("index1")
index2=request.form("index2")
index3=request.form("index3")
index4=request.form("index4")
index5=request.form("index5")
index6=request.form("index6")
index7=request.form("index7")
index8=request.form("index8")



'���������б� 
urllist=urllist & chr(60) & "%" & VbCrLf
urllist=urllist & "company=" & chr(34) & ""&company&"" & chr(34) & VbCrLf
urllist=urllist & "sitetitle=" & chr(34) & ""&sitetitle&"" & chr(34) & VbCrLf
urllist=urllist & "gongsurl=" & chr(34) & ""&gongsurl&"" & chr(34) & VbCrLf
urllist=urllist & "webDescription=" & chr(34) & ""&webDescription&"" & chr(34) & VbCrLf
urllist=urllist & "webKeywords=" & chr(34) & ""&webKeywords&"" & chr(34) & VbCrLf
urllist=urllist & "weblogo=" & chr(34) & ""&weblogo&"" & chr(34) & VbCrLf
urllist=urllist & "weblogotype=" & chr(34) & ""&weblogotype&"" & chr(34) & VbCrLf
urllist=urllist & "dataurl=" & chr(34) & ""&dataurl&"" & chr(34) & VbCrLf
urllist=urllist & "backurl=" & chr(34) & ""&backurl&"" & chr(34) & VbCrLf
urllist=urllist & "admindataurl=" & chr(34) & "../" & ""&dataurl&"" & chr(34) & VbCrLf
urllist=urllist & "address=" & chr(34) & ""&address&"" & chr(34) & VbCrLf
urllist=urllist & "code=" & chr(34) & ""&code&"" & chr(34) & VbCrLf
urllist=urllist & "companycon=" & chr(34) & ""&companycon&"" & chr(34) & VbCrLf
urllist=urllist & "phone=" & chr(34) & ""&phone&"" & chr(34) & VbCrLf
urllist=urllist & "fax=" & chr(34) & ""&fax&"" & chr(34) & VbCrLf
urllist=urllist & "email=" & chr(34) & ""&email&"" & chr(34) & VbCrLf
urllist=urllist & "reply=" & chr(34) & ""&reply&"" & chr(34) & VbCrLf
urllist=urllist & "kehu=" & chr(34) & ""&kehu&"" & chr(34) & VbCrLf
urllist=urllist & "chengxuup=" & chr(34) & ""&chengxuup&"" & chr(34) & VbCrLf
urllist=urllist & "companycontent=" & chr(34) & ""&companycontent&"" & chr(34) & VbCrLf
urllist=urllist & "qq1=" & chr(34) & ""&qq1&"" & chr(34) & VbCrLf
urllist=urllist & "qq2=" & chr(34) & ""&qq2&"" & chr(34) & VbCrLf
urllist=urllist & "qq3=" & chr(34) & ""&qq3&"" & chr(34) & VbCrLf
urllist=urllist & "beianhao=" & chr(34) & ""&beianhao&"" & chr(34) & VbCrLf
urllist=urllist & "ipaddress=" & chr(34) & ""&ipaddress&"" & chr(34) & VbCrLf
urllist=urllist & "index1=" & chr(34) & ""&index1&"" & chr(34) & VbCrLf
urllist=urllist & "index2=" & chr(34) & ""&index2&"" & chr(34) & VbCrLf
urllist=urllist & "index3=" & chr(34) & ""&index3&"" & chr(34) & VbCrLf
urllist=urllist & "index4=" & chr(34) & ""&index4&"" & chr(34) & VbCrLf
urllist=urllist & "index5=" & chr(34) & ""&index5&"" & chr(34) & VbCrLf
urllist=urllist & "index6=" & chr(34) & ""&index6&"" & chr(34) & VbCrLf
urllist=urllist & "index7=" & chr(34) & ""&index7&"" & chr(34) & VbCrLf
urllist=urllist & "index8=" & chr(34) & ""&index8&"" & chr(34) & VbCrLf
urllist=urllist & "index9=" & chr(34) & ""&index9&"" & chr(34) & VbCrLf
urllist=urllist & "index2=" & chr(34) & ""&index2&"" & chr(34) & VbCrLf

urllist=urllist & "%" & chr(62) & VbCrLf

'----����Ŀ¼�ļ��кͷ����ļ�����
dim filename
'----���ɷ����ļ�����
filename="../inc/config.asp"

Set fso = Server.CreateObject("Scripting.FileSystemObject")
  Set fout = fso.Createtextfile(server.mappath(filename),true)
 fout.writeline urllist
fout.close 
Response.Write"<script language=JavaScript>"
Response.Write"alert(""��ϲ.�޸ĳɹ�"");"
Response.Write"window.location='SiteSetup.asp'"
Response.Write"</script>"
end if
%>	
<script language="JavaScript" type="text/JavaScript">
function ConfirmDelBig()
{
   if(confirm("��ȷ��Ҫ�޸��ύ��������"))
     return true;
   else
     return false;
	 
}
</script>
<html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/style.css" rel="stylesheet" type="text/css">
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>��</td>
  </tr>
</table>
<TABLE WIDTH="90%" BORDER="0" align="center" CELLPADDING="0" CELLSPACING="0">
  
  <TR BGCOLOR="#eeeeee">
    <TD 
HEIGHT="30" bgcolor="#FFFFFF"> 
      <form name="form1" method="post" action="?xiugai=yes">
        <table width="100%" border="1" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
          <tr bgcolor="#f9f9f9"> 
            <td height="25" colspan="3" background="images/b1.gif"> <div align="center"><strong><font color="#FFFFFF">��վ��������</font></strong></div></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center">��˾����</div></td>
            <td colspan="2"> <input name="company" type="text" id="company" value="<% =company %>" size="30">
              ������Ĺ�˾����(�255��)���磺<font color="#FF0000">������˾</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td width="10%" height="25"> <div align="center">��վ����</div></td>
            <td colspan="2" bgcolor="#f9f9f9"> <input name="sitetitle" type="text" id="sitetitle" value="<% =sitetitle %>" size="30">
              ��������վ����</td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td width="10%" height="25"> <div align="center">��˾��ַ</div></td>
            <td colspan="2"> <input name="gongsurl" type="text" id="sitetitle3" value="<% =gongsurl %>" size="30"> 
              <input name="dataurl" type="hidden" id="dataurl" value="<% =dataurl %>" size="30" onfocus="document.form1.address.focus();">
              ���뺬 http:// �磺<font color="#FF0000">http://www.lzkeyuan.cn</font></td>
          </tr>
		            <tr bgcolor="#f9f9f9"> 
            <td height="25" bgcolor="#f9f9f9"><div align="center">��վ����</div></td>
            <td colspan="2" bgcolor="#f9f9f9"><input name="webDescription" type="text" id="webDescription" value="<% =webDescription %>" size="30" maxlength="100">
              100�����ڣ������ǹ�˾��Ӫ��Χ����վ���ݽ��ܡ�</td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25" bgcolor="#f9f9f9"><div align="center">��վ�ؼ���</div></td>
            <td colspan="2" bgcolor="#f9f9f9"><input name="webKeywords" type="text" id="sitetitle4" value="<% =webKeywords %>" size="30" maxlength="80">
              80�����ڣ�����ؼ���֮���á� ,���ָ�磺<font color="#FF0000">��桢��������˾�����</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td width="10%" height="25"> <div align="center">��վͼƬ</div></td>
            <td width="33%"> <input name="weblogo" type="text" id="gongsurl" value="<% =weblogo %>" size="30"> 
            </td>
            <td width="57%"> <iframe style="top:2px" ID="UploadFiles" src="shangchuan_index.asp?picID=0" frameborder=0 scrolling=no width="320" height="25"></iframe></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center">�ء���ַ</div></td>
            <td colspan="2"> <input name="address" type="text" id="address" value="<% =address %>" size="30">
              ��˾<br>
			��ϵ��ַ���磺<font color="#FF0000">�����С���·����</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center">�ʡ�����</div></td>
            <td colspan="2"> <input name="code" type="text" id="code" value="<% =code %>" size="30">
              �������룬�磺<font color="#FF0000">730000</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center">�� ϵ ��</div></td>
            <td colspan="2"> <input name="companycon" type="text" id="companycon" value="<% =companycon %>" size="30">
              �����빫˾<br>
			��ϵ���������磺<font color="#FF0000">���� </font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center">�硡����</div></td>
            <td colspan="2"> <input name="phone" type="text" id="phone" value="<% =phone %>" size="30">
              ��˾<br>
			��ϵ�绰������绰�ö��Ÿ������磺<font color="#FF0000">0931-0000000��0931-0000000</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center">��ϵ�绰</div></td>
            <td colspan="2"> <input name="fax" type="text" id="fax" value="<% =fax %>" size="30">
              ��˾<br>
			���棬�磺<font color="#FF0000">0931-0000000</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25" bgcolor="#f9f9f9"> <div align="center">�ʡ�����</div></td>
            <td colspan="2"> <input name="email" type="text" id="email" value="<% =email %>" size="30">
              ��ϵE_mail,�磺<font color="#FF0000">wsxs126@163.com</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25" bgcolor="#f9f9f9"><div align="center">Ͷ�ߵ绰</div></td>
            <td colspan="2"> <input name="beianhao" type="text" id="beianhao" value="<% =beianhao %>" size="30">
              ICP������ ,�磺<font color="#FF0000">¤ICP��07500075��</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="12" bgcolor="#f9f9f9"><div align="center">ҵ��QQ</div></td>
            <td colspan="2"> <input name="qq1" type="text" id="qq1" value="<% =qq1 %>" size="30">
              ��һ��ҵ��QQ��<font color="#FF0000">��������QQ���룬�磺664126937</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="12" bgcolor="#f9f9f9"><div align="center">ҵ��QQ</div></td>
            <td colspan="2"> <input name="qq2" type="text" id="qq2" value="<% =qq2 %>" size="30">
              �ڶ���ҵ��QQ��<font color="#FF0000">��������QQ����</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25" bgcolor="#f9f9f9"><div align="center">ҵ��QQ</div></td>
            <td colspan="2"><input name="qq3" type="text" id="qq3" value="<% =qq3 %>" size="30">
              �ڶ���ҵ��QQ��<font color="#FF0000">��������QQ����</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="12"> <div align="center">��������</div></td>
            <td colspan="2"> <div align="left"> 
                <select name="reply" id="reply">
                  <option value="no" <% if reply="no" then %>selected<% end if %>>��Ҫ����Ա��˺���ʾ</option>
                  <option value="yes"<% if reply="yes" then %>selected<% end if %>>����Ҫ��ˣ�ֱ����ʾ</option>
                </select>
                �����Ƿ���Ҫ��˲�����ʾ������ѡ��</div></td>
          </tr>
          <tr bgcolor="#f9f9f9">
            <td height="12"><div align="center">�����ͷ�</div></td>
            <td colspan="2"> <select name="kehu" id="kehu">
                <option value="no" <% if kehu="no" then %>selected<% end if %>>����Ҫ�����ͷ�ͼ��</option>
                <option value="yes"<% if kehu="yes" then %>selected<% end if %>>��Ҫ�����ͷ�ͼ��</option>
              </select>
              ��ҳ�Ҳม�������߿ͷ���</td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center"><font color="#FF0000">��ȫ����</font></div></td>
            <td colspan="2"><input name="ipaddress" type="text" id="ipaddress" value="<% =ipaddress %>" size="30">
              ���ʺ�̨��IP��ַ�������ڵ�IP��ַ�ǣ� <font color="#FF0000"> 
              <% =request.servervariables("REMOTE_ADDR") %>
              </font> ����Ϊ������,<strong><font color="#FF0000">����������</font></strong>��</td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="13"><div align="center">��ҳ��Ŀ1</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "���������Ŀ��"
		else
		%><select name="index1" size="1" id="index1">
                <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		    selectclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
                <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		      rs.movenext
    	    loop
			%>
              </select> 
              <%
		end if
        rs.close
			%>
              <SCRIPT type=text/javascript>form1.index1.value = "<% =index1 %>";
		</SCRIPT>
              ������Ҫ����ҳ��ʾ�ĵ�һ����Ŀ���ƣ���ģ����Ч��</td>
          </tr>
          <tr>
            <td height="13"><div align="center">��ҳ��Ŀ2</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "���������Ŀ��"
		else
		%><select name="index2" size="1" id="index2">
                <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		    selectclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
                <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		      rs.movenext
    	    loop
			%>
              </select> 
              <%
		end if
        rs.close
			%>
              <SCRIPT type=text/javascript>form1.index2.value = "<% =index2 %>";
		</SCRIPT>
              ������Ҫ����ҳ��ʾ�ĵڶ�����Ŀ���ƣ���ģ����Ч��</td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="13"><div align="center">��ҳ��Ŀ3</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "���������Ŀ��"
		else
		%><select name="index3" size="1" id="index3">
                <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		    selectclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
                <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		      rs.movenext
    	    loop
			%>
              </select> 
              <%
		end if
        rs.close
			%>
              <SCRIPT type=text/javascript>form1.index3.value = "<% =index3 %>";
		</SCRIPT>
              ������Ҫ����ҳ��ʾ�ĵ�������Ŀ���ƣ���ģ����Ч��</td>
          </tr>
          <tr>
            <td height="13"><div align="center">��ҳ��Ŀ4</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "���������Ŀ��"
		else
		%><select name="index4" size="1" id="index4">
                <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		    selectclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
                <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		      rs.movenext
    	    loop
			%>
              </select> 
              <%
		end if
        rs.close
			%>
              <SCRIPT type=text/javascript>form1.index4.value = "<% =index4 %>";
		</SCRIPT>
              ������Ҫ����ҳ��ʾ�ĵ��ĸ���Ŀ���ƣ���ģ����Ч��</td>
          </tr>
			<tr>
            <td height="13"><div align="center">��ҳ��Ŀ5</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "���������Ŀ��"
		else
		%><select name="index5" size="1" id="index5">
                <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		    selectclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
                <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		      rs.movenext
    	    loop
			%>
              </select> 
              <%
		end if
        rs.close
			%>
              <SCRIPT type=text/javascript>form1.index5.value = "<% =index5 %>";
		</SCRIPT>
              ������Ҫ����ҳ��ʾ�ĵ������Ŀ���ƣ���ģ����Ч��</td>
          </tr>
			<tr>
            <td height="13"><div align="center">��ҳ��Ŀ6</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "���������Ŀ��"
		else
		%><select name="index6" size="1" id="index6">
                <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		    selectclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
                <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		      rs.movenext
    	    loop
			%>
              </select> 
              <%
		end if
        rs.close
			%>
              <SCRIPT type=text/javascript>form1.index6.value = "<% =index6 %>";
		</SCRIPT>
              ������Ҫ����ҳ��ʾ�ĵ�������Ŀ���ƣ���ģ����Ч��</td>
          </tr>
			<tr>
            <td height="13"><div align="center">��ҳ��Ŀ7</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "���������Ŀ��"
		else
		%><select name="index7" size="1" id="index7">
                <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		    selectclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
                <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		      rs.movenext
    	    loop
			%>
              </select> 
              <%
		end if
        rs.close
			%>
              <SCRIPT type=text/javascript>form1.index7.value = "<% =index7 %>";
		</SCRIPT>
              ������Ҫ����ҳ��ʾ�ĵ��߸���Ŀ���ƣ���ģ����Ч��</td>
          </tr>
			<tr>
            <td height="13"><div align="center">��ҳ��Ŀ8</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "���������Ŀ��"
		else
		%><select name="index8" size="1" id="index8">
                <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		    selectclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
                <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		      rs.movenext
    	    loop
			%>
              </select> 
              <%
		end if
        rs.close
			%>
              <SCRIPT type=text/javascript>form1.index2.value = "<% =index2 %>";
		</SCRIPT>
              ������Ҫ����ҳ��ʾ�ĵڶ�����Ŀ���ƣ���ģ����Ч��</td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td> <div align="center">��˾���</div></td>
            <td colspan="2"> <textarea name="companycontent" cols="70" rows="8" id="companycontent"><% companycontent= Replace( companycontent, "<br>",vbCrLf ) %><% =companycontent %></textarea>
              ������250���ڹ�˾<br>
			����</td>
          </tr>
        </table>
        <div align="center"><br>
          <input type="submit" name="Submit" value=" �� �� �� �� " onClick="return ConfirmDelBig();">
          ����
          <input type="reset" name="Submit2" value=" �� �� �� д ">
        </div>
      </form></TD>
  </TR>
</TABLE>


</body>
</html>