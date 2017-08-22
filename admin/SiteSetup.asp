<!--#include file="adminconn.asp"-->
<%
dim urllist,fso,fout
if request("xiugai")="yes" then
'读取提交数据
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



'生成数据列表 
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

'----定义目录文件夹和分类文件名称
dim filename
'----生成分类文件名称
filename="../inc/config.asp"

Set fso = Server.CreateObject("Scripting.FileSystemObject")
  Set fout = fso.Createtextfile(server.mappath(filename),true)
 fout.writeline urllist
fout.close 
Response.Write"<script language=JavaScript>"
Response.Write"alert(""恭喜.修改成功"");"
Response.Write"window.location='SiteSetup.asp'"
Response.Write"</script>"
end if
%>	
<script language="JavaScript" type="text/JavaScript">
function ConfirmDelBig()
{
   if(confirm("您确定要修改提交的内容吗！"))
     return true;
   else
     return false;
	 
}
</script>
<html>
<head>
<title>管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/style.css" rel="stylesheet" type="text/css">
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>　</td>
  </tr>
</table>
<TABLE WIDTH="90%" BORDER="0" align="center" CELLPADDING="0" CELLSPACING="0">
  
  <TR BGCOLOR="#eeeeee">
    <TD 
HEIGHT="30" bgcolor="#FFFFFF"> 
      <form name="form1" method="post" action="?xiugai=yes">
        <table width="100%" border="1" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
          <tr bgcolor="#f9f9f9"> 
            <td height="25" colspan="3" background="images/b1.gif"> <div align="center"><strong><font color="#FFFFFF">网站参数设置</font></strong></div></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center">公司名称</div></td>
            <td colspan="2"> <input name="company" type="text" id="company" value="<% =company %>" size="30">
              输入你的公司名称(最长255字)，如：<font color="#FF0000">××公司</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td width="10%" height="25"> <div align="center">网站名称</div></td>
            <td colspan="2" bgcolor="#f9f9f9"> <input name="sitetitle" type="text" id="sitetitle" value="<% =sitetitle %>" size="30">
              请输入网站名称</td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td width="10%" height="25"> <div align="center">公司网址</div></td>
            <td colspan="2"> <input name="gongsurl" type="text" id="sitetitle3" value="<% =gongsurl %>" size="30"> 
              <input name="dataurl" type="hidden" id="dataurl" value="<% =dataurl %>" size="30" onfocus="document.form1.address.focus();">
              必须含 http:// 如：<font color="#FF0000">http://www.lzkeyuan.cn</font></td>
          </tr>
		            <tr bgcolor="#f9f9f9"> 
            <td height="25" bgcolor="#f9f9f9"><div align="center">网站描述</div></td>
            <td colspan="2" bgcolor="#f9f9f9"><input name="webDescription" type="text" id="webDescription" value="<% =webDescription %>" size="30" maxlength="100">
              100字以内，可以是公司经营范围或网站内容介绍。</td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25" bgcolor="#f9f9f9"><div align="center">网站关键字</div></td>
            <td colspan="2" bgcolor="#f9f9f9"><input name="webKeywords" type="text" id="sitetitle4" value="<% =webKeywords %>" size="30" maxlength="80">
              80字以内，多个关键字之间用“ ,”分割，如：<font color="#FF0000">广告、体育、公司、礼佛</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td width="10%" height="25"> <div align="center">网站图片</div></td>
            <td width="33%"> <input name="weblogo" type="text" id="gongsurl" value="<% =weblogo %>" size="30"> 
            </td>
            <td width="57%"> <iframe style="top:2px" ID="UploadFiles" src="shangchuan_index.asp?picID=0" frameborder=0 scrolling=no width="320" height="25"></iframe></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center">地　　址</div></td>
            <td colspan="2"> <input name="address" type="text" id="address" value="<% =address %>" size="30">
              公司<br>
			联系地址，如：<font color="#FF0000">××市××路×号</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center">邮　　编</div></td>
            <td colspan="2"> <input name="code" type="text" id="code" value="<% =code %>" size="30">
              邮政编码，如：<font color="#FF0000">730000</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center">联 系 人</div></td>
            <td colspan="2"> <input name="companycon" type="text" id="companycon" value="<% =companycon %>" size="30">
              请输入公司<br>
			联系人姓名，如：<font color="#FF0000">张三 </font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center">电　　话</div></td>
            <td colspan="2"> <input name="phone" type="text" id="phone" value="<% =phone %>" size="30">
              公司<br>
			联系电话，多个电话用逗号隔开，如：<font color="#FF0000">0931-0000000，0931-0000000</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center">联系电话</div></td>
            <td colspan="2"> <input name="fax" type="text" id="fax" value="<% =fax %>" size="30">
              公司<br>
			传真，如：<font color="#FF0000">0931-0000000</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25" bgcolor="#f9f9f9"> <div align="center">邮　　箱</div></td>
            <td colspan="2"> <input name="email" type="text" id="email" value="<% =email %>" size="30">
              联系E_mail,如：<font color="#FF0000">wsxs126@163.com</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25" bgcolor="#f9f9f9"><div align="center">投诉电话</div></td>
            <td colspan="2"> <input name="beianhao" type="text" id="beianhao" value="<% =beianhao %>" size="30">
              ICP备案号 ,如：<font color="#FF0000">陇ICP备07500075号</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="12" bgcolor="#f9f9f9"><div align="center">业务QQ</div></td>
            <td colspan="2"> <input name="qq1" type="text" id="qq1" value="<% =qq1 %>" size="30">
              第一个业务QQ，<font color="#FF0000">这里输入QQ号码，如：664126937</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="12" bgcolor="#f9f9f9"><div align="center">业务QQ</div></td>
            <td colspan="2"> <input name="qq2" type="text" id="qq2" value="<% =qq2 %>" size="30">
              第二个业务QQ，<font color="#FF0000">这里输入QQ号码</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25" bgcolor="#f9f9f9"><div align="center">业务QQ</div></td>
            <td colspan="2"><input name="qq3" type="text" id="qq3" value="<% =qq3 %>" size="30">
              第二个业务QQ，<font color="#FF0000">这里输入QQ号码</font></td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="12"> <div align="center">留言设置</div></td>
            <td colspan="2"> <div align="left"> 
                <select name="reply" id="reply">
                  <option value="no" <% if reply="no" then %>selected<% end if %>>需要管理员审核后显示</option>
                  <option value="yes"<% if reply="yes" then %>selected<% end if %>>不需要审核，直接显示</option>
                </select>
                留言是否需要审核才能显示，请点击选择</div></td>
          </tr>
          <tr bgcolor="#f9f9f9">
            <td height="12"><div align="center">浮动客服</div></td>
            <td colspan="2"> <select name="kehu" id="kehu">
                <option value="no" <% if kehu="no" then %>selected<% end if %>>不需要浮动客服图标</option>
                <option value="yes"<% if kehu="yes" then %>selected<% end if %>>需要浮动客服图标</option>
              </select>
              网页右侧浮动“在线客服”</td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="25"> <div align="center"><font color="#FF0000">安全设置</font></div></td>
            <td colspan="2"><input name="ipaddress" type="text" id="ipaddress" value="<% =ipaddress %>" size="30">
              访问后台的IP地址，你现在的IP地址是： <font color="#FF0000"> 
              <% =request.servervariables("REMOTE_ADDR") %>
              </font> ，空为不限制,<strong><font color="#FF0000">不建议设置</font></strong>。</td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="13"><div align="center">首页栏目1</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
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
              请输入要在首页显示的第一个栏目名称，该模板有效。</td>
          </tr>
          <tr>
            <td height="13"><div align="center">首页栏目2</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
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
              请输入要在首页显示的第二个栏目名称，该模板有效。</td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td height="13"><div align="center">首页栏目3</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
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
              请输入要在首页显示的第三个栏目名称，该模板有效。</td>
          </tr>
          <tr>
            <td height="13"><div align="center">首页栏目4</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
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
              请输入要在首页显示的第四个栏目名称，该模板有效。</td>
          </tr>
			<tr>
            <td height="13"><div align="center">首页栏目5</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
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
              请输入要在首页显示的第五个栏目名称，该模板有效。</td>
          </tr>
			<tr>
            <td height="13"><div align="center">首页栏目6</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
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
              请输入要在首页显示的第六个栏目名称，该模板有效。</td>
          </tr>
			<tr>
            <td height="13"><div align="center">首页栏目7</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
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
              请输入要在首页显示的第七个栏目名称，该模板有效。</td>
          </tr>
			<tr>
            <td height="13"><div align="center">首页栏目8</div></td>
            <td colspan="2"> 
              <%
		set rs=server.createobject("adodb.recordset")
        sql = "select BigClassName from BigClass where ClassType='class.asp' "
        rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
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
              请输入要在首页显示的第二个栏目名称，该模板有效。</td>
          </tr>
          <tr bgcolor="#f9f9f9"> 
            <td> <div align="center">公司简介</div></td>
            <td colspan="2"> <textarea name="companycontent" cols="70" rows="8" id="companycontent"><% companycontent= Replace( companycontent, "<br>",vbCrLf ) %><% =companycontent %></textarea>
              请输入250字内公司<br>
			介绍</td>
          </tr>
        </table>
        <div align="center"><br>
          <input type="submit" name="Submit" value=" 提 交 修 改 " onClick="return ConfirmDelBig();">
          　　
          <input type="reset" name="Submit2" value=" 重 新 填 写 ">
        </div>
      </form></TD>
  </TR>
</TABLE>


</body>
</html>