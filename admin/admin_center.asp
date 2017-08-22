<%@ Language="VBScript" CODEPAGE="936"%>

<% Option Explicit %>
<%
  if session("aleave")="" then
      response.redirect "adminlogin.asp"
	  response.end
  end if
%>
<% 
Response.Buffer = True
Dim startime
	 startime=timer()
Dim hx
Set hx = New Cls_AspCheck

class Cls_AspCheck

Public FileName,WebName,WebUrl,SysName,SysNameE,SysVersion

'检查组件是否被支持
Public Function IsObjInstalled(strClassString)
	On Error Resume Next
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err Then
		IsObjInstalled = False
	else	
		IsObjInstalled = True
	end if
	Set xTestObj = Nothing
End Function

'检查组件版本
Public Function getver(Classstr)
	On Error Resume Next
	Dim xTestObj
	Set xTestObj = Server.CreateObject(Classstr)
	If Err Then
		getver=""
	else	
	 	getver=xTestObj.version
	end if
	Set xTestObj = Nothing
End Function

Public Function GetObjInfo(startnum,endnum)
	dim i,Outstr
	for i=startnum to endnum
      	Outstr = Outstr & "<tr bgcolor=#FFFFFF align=center height=18><TD align=left>&nbsp;" & theTestObj(i,0) & ""
      	Outstr = Outstr & "<font color=#888888>&nbsp;"&theTestObj(i,1)&"</font>"
      	Outstr = Outstr & "</td>"
    	If Not IsObjInstalled(theTestObj(i,0)) Then 
      	Outstr = Outstr & "<td align=left>&nbsp;<font color=red><b>×</b></font></td>"
    	Else
      	Outstr = Outstr & "<td align=left>&nbsp;<font color=green><b>√</b></font> " & getver(theTestObj(i,0)) & "</td>"
		End If
      	Outstr = Outstr & "</tr>" & vbCrLf
	next
	Response.Write(Outstr)
End Function

Public Function formatvariables(str)
on error resume next
str = cstr(server.htmlencode(str))
formatvariables=replace(str,chr(10),"<br>")
End Function

Public Sub ShowFooter()
	dim Endtime,Runtime,OutStr
	Endtime=timer()
	OutStr = "<table border=0 cellpadding=0 cellspacing=1 class=tableBorder><tr><td align=center>"
	OutStr = OutStr & "<br><p>Copyright 2006 网站管理系统</p>"& vbcrlf
 	Runtime=FormatNumber((endtime-startime)*1000,2) 
	if Runtime>0 then
		if Runtime>1000 then
		  	OutStr = OutStr & "页面执行时间：约"& FormatNumber(runtime/1000,2) & "秒"
		else
			OutStr = OutStr & "页面执行时间：约"& Runtime & "毫秒"
		end if	
	end if				
	OutStr = OutStr & "</p></td></tr></table>"
	Response.Write(OutStr)
End Sub
End class

Dim theTestObj(25,1)

	theTestObj(0,0) = "Scripting.FileSystemObject"
	theTestObj(0,1) = "(FSO 文本文件读写)"
	theTestObj(1,0) = "ADODB.Connection"
	theTestObj(1,1) = "(ADO 数据对象)"  
	theTestObj(2,0) = "JMail.SmtpMail"
	theTestObj(2,1) = "(Dimac JMail 邮件收发)</a>"
	theTestObj(3,0) = "CDONTS.NewMail"
	theTestObj(3,1) = "(虚拟 SMTP 发信)"
	theTestObj(4,0) = "SoftArtisans.ImageGen"
	theTestObj(4,1) = "(SA 的图像读写组件)"
	theTestObj(5,0) = "W3Image.Image"
	theTestObj(5,1) = "(Dimac 的图像读写组件)"


%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>com</TITLE>
<link href="images/style.css" rel="stylesheet" type="text/css">
<style>
<!--
A       { COLOR: #376ed0; TEXT-DECORATION: none}
A:hover { COLOR: green}
body,td,span { font-size: 9pt}
.input  { BACKGROUND-COLOR: #ffffff;BORDER:#CCCCCC 1px solid;FONT-SIZE: 9pt}
.backc  { BACKGROUND-COLOR: #CCCCCC;BORDER:#CCCCCC 1px solid;FONT-SIZE: 9pt;color:white}
.PicBar { background-color: #CCCCCC; border: 1px solid #376ed0; height: 12px;}
.tableBorder {BORDER-RIGHT: #183789 1px solid; BORDER-TOP: #183789 1px solid; BORDER-LEFT: #183789 1px solid; BORDER-BOTTOM: #183789 1px solid; BACKGROUND-COLOR: #ffffff; WIDTH: 600;}
.divcenter {position:absolute;height:30px;z-index:1000;}
-->
</STYLE>
<SCRIPT language="JavaScript" runat="server">
	function getEngVerJs(){
		try{
			return ScriptEngineMajorVersion() +"."+ScriptEngineMinorVersion()+"."+ ScriptEngineBuildVersion() + " ";
		}catch(e){
			return "服务器不支持此项检测";
		}
		
	}
</SCRIPT>
<SCRIPT language="VBScript" runat="server">
	Function getEngVerVBS()
		getEngVerVBS=ScriptEngineMajorVersion() &"."&ScriptEngineMinorVersion() &"." & ScriptEngineBuildVersion() & " "
	End Function
</SCRIPT>
<script language="javascript">
<!--
function Checksearchbox()
{
if(form1.classname.value == "")
{
	alert("请输入你要检测的组件名！");
	form1.classname.focus();
	return false;
}
}
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
eval("txt" + sid + ".innerHTML=\"<a href='#' title='关闭此项'><font face='Wingdings' color=#FFFFFF>x</font></a>\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
eval("txt" + sid + ".innerHTML=\"<a href='#' title='打开此项'><font face='Wingdings' color=#FFFFFF>y</font></a>\";");
}
}
-->
</SCRIPT>
</HEAD>
<BODY leftmargin="50">
<table width="600" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="50" align="center"> <p><font color="#0099FF" size="6"><strong>服务器信息 
        <a href="../help/index.htm" target="_blank"><font color="#FF0000">点击查看帮助文档</font></a></strong></font></p>
      </td>
  </tr>
</table>
<br>
<%
dim action
action=request("action")
if action="testzujian" then
call ObjTest2
end if

Call menu
Call SystemTest
Call ObjTest
Call CalculateTest
Call DriveTest
Call SpeedTest
hx.ShowFooter
Set hx= nothing

%>
<%Sub menu%>
选项：<a href="#SystemTest">服务器有关参数</a> | <a href="#ObjTest">服务器组件情况</a> | <a href="#SpeedTest">服务器连接速度</a> 
<%End Sub%>
<%Sub smenu(i)%>
<a href="#top" title="返回顶部"><font face='Webdings' color=#FFFFFF>5</font></a> <span id=txt<%=i%> name=txt<%=i%>><a href='#' title='关闭此项'><font face='Wingdings' color=#FFFFFF>x</font></a></span> 
<%End Sub%>
<%Sub SystemTest
on error resume next
%>
<a name="SystemTest"></a> 
<table width="90%" border="0" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr> 
    <td height="25" align="center" bgcolor="#CCCCCC" onclick="showsubmenu(0)" background="images/b1.gif"><font color=#FFFFFF><strong>服务器有关参数</strong></font> 
      <%Call smenu(0)%></td>
  </tr>
  <tr> 
    <td style="display" id='submenu0'><table border=0 width=100% cellspacing=1 cellpadding=3 bgcolor="#376ed0">
        <tr bgcolor="#FFFFFF" height=18> 
          <td width="130">&nbsp;服务器名</td>
          <td width="170" height="18">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
          <td width="130" height="18">&nbsp;服务器操作系统</td>
          <td width="170" height="18">&nbsp;<%=Request.ServerVariables("OS")%></td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td>&nbsp;服务器IP</td>
          <td>&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
          <td>&nbsp;服务器端口</td>
          <td>&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td>&nbsp;服务器时间</td>
          <td>&nbsp;<%=now%></td>
          <td>&nbsp;服务器CPU数量</td>
          <td>&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 
            个</td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td>&nbsp;IIS版本</td>
          <td height="18">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
          <td height="18">&nbsp;脚本超时时间</td>
          <td height="18">&nbsp;<%=Server.ScriptTimeout%> 秒</td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td>&nbsp;Application变量</td>
          <td height="18">&nbsp;<%Response.Write(Application.Contents.Count & "个 ")
		  if Application.Contents.count>0 then Response.Write("[<a href=""?action=showapp"">遍历Application变量</a>]")%>
          </td>
          <td height="18">&nbsp;Session变量<br> </td>
          <td height="18">&nbsp;<%Response.Write(Session.Contents.Count&"个 ")
		  if Session.Contents.count>0 then Response.Write("[<a href=""?action=showsession"">遍历Session变量</a>]")%>
          </td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td height="18">&nbsp;<a href="?action=showvariables">所有服务器参数</a></td>
          <td height="18">&nbsp;<%Response.Write(Request.ServerVariables.Count&"个 ")
		  if Request.ServerVariables.Count>0 then Response.Write("[<a href=""?action=showvariables"">遍历服务器参数</a>]")%>
          </td>
          <td height="18">&nbsp;服务器环境变量</td>
          <td height="18">&nbsp;<%
			dim WshShell,WshSysEnv
			Set WshShell = server.CreateObject("WScript.Shell")
			Set WshSysEnv = WshShell.Environment
			if err then
				Response.Write("服务器不支持WScript.Shell组件")
				err.clear
			else
				Response.Write(WshSysEnv.count &"个 ")
				if WshSysEnv.count>0 then Response.Write("[<a href=""?action=showwsh"">遍历环境变量</a>]") 
		 	end if
		  %>		  
		  </td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td align=left>&nbsp;服务器解译引擎</td>
          <td height="18" colspan="3">&nbsp;JScript: <%= getEngVerJs() %> | VBScript: <%=getEngVerVBS()%></td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td align=left bgcolor="#FFFFFF">&nbsp;本文件实际路径</td>
          <td height="8" colspan="3" bgcolor="#FFFFFF">&nbsp;<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
        </tr>
      </table>
      <%
if action="showapp" or action="showsession" or action="showvariables" or action="showwsh" then
	showvariable(action)
end if
%>
    </td>
  </tr>
</table>
<br>
<%
End Sub

Sub showvariable(action)
%>
<table width="90%" border="0" cellpadding="3" cellspacing="1" bgcolor="#376ed0">
  <tr bgcolor="#CCCCCC"> 
    <td colspan="2" background="images/b1.gif"> &nbsp;&nbsp; 
      <%
	on error resume next
	dim Item,xTestObj,outstr
	if action="showapp" then
		Response.Write("<font face='Webdings'>4</font> 遍历Application变量")
		set xTestObj=Application.Contents
	elseif action="showsession" then
		Response.Write("<font face='Webdings'>4</font> 遍历Session变量")
		set xTestObj=Session.Contents
	elseif action="showvariables" then
		Response.Write("<font face='Webdings'>4</font> 遍历服务器参数")
		set xTestObj=Request.ServerVariables
	elseif action="showwsh" then
		Response.Write("<font face='Webdings'>4</font> 遍历环境变量")
		dim WshShell
		Set WshShell = server.CreateObject("WScript.Shell")
		set xTestObj=WshShell.Environment
	end if
		Response.Write "(<a href="&hx.FileName&">关闭</a>)"
	%>
    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="130">变量名</td>
    <td width="470">值</td>
  </tr>
  <%
	if err then
		outstr = "<tr bgcolor=#FFFFFF><td colspan=2>没有符合条件的变量</td></tr>"
		err.clear
	else
		dim w
		if action="showwsh" then
			for each Item in xTestObj
				w=split(Item,"=")
				outstr = outstr & "<tr bgcolor=#FFFFFF>"
				outstr = outstr & "<td>" & w(0) & "</td>" 
				outstr = outstr & "<td>" & w(1) & "</td>" 
				outstr = outstr & "</tr>" 		
			next
		else
			dim i
			for each Item in xTestObj
				outstr = outstr & "<tr bgcolor=#FFFFFF>"
				outstr = outstr & "<td>" & Item & "</td>" 				
				outstr = outstr & "<td>"
				if IsArray(xTestObj(Item)) then		
					for i=0 to ubound(xTestObj(Item))-1
					outstr = outstr & hx.formatvariables(xTestObj(Item)(i)) & "<br>"
					next
				else
					outstr = outstr & hx.formatvariables(xTestObj(Item))
				end if			
				outstr = outstr & "</td>"
				outstr = outstr & "</tr>" 
			next
		end if
	end if
		Response.Write(outstr)	
		set xTestObj=nothing
		%>
</table>
<%End Sub%>
<%Sub ObjTest%>
<a name="ObjTest"></a>
<table border="0" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr> 
    <td height="25" align="center" bgcolor="#CCCCCC" onclick="showsubmenu(1)" background="images/b1.gif"><strong><font color="#FFFFFF">服务器</font></strong><font color=#FFFFFF><strong>组件情况</strong></font> 
      <%Call smenu(1)%>
    </td>
  </tr>
  <tr> 
    <td style="display" id='submenu1'><table border=0 width=100% cellspacing=1 cellpadding=3 bgcolor="#376ed0">

        <tr bgcolor="#FFFFFF" height=18> 
          <td width=450 align="center">组 件 名 称</td>
          <td width=150 align="center">支持及版本</td>
        </tr>
        <%hx.GetObjInfo 0,5%>
      </table>
      <table border=0 width=100% cellspacing=1 cellpadding=3  bgcolor="#376ed0">
        <tr> 
          <td height="25" background="images/b1.gif" bgcolor="#CCCCCC">&nbsp;&nbsp;<font face='Webdings'>4</font> 
            其他组件支持情况检测 </td>
        </tr>
        <FORM action=?action=testzujian method=post id=form1 name=form1 onSubmit="JavaScript:return Checksearchbox();">
          <tr> 
            <td height=30 bgcolor="#FFFFFF">输入你要检测的组件的ProgId或ClassId 
              <input class=input type=text value="" name="classname" size=40> 
              <INPUT type=submit value="确定" class=backc id=submit1 name=submit1> 
            </td>
          </tr>
        </FORM>
      </table></td>
  </tr>
</table>
<br>
<%
End Sub
Sub ObjTest2
	Dim strClass
    strClass = Trim(Request.Form("classname"))
    If strClass <> "" then
	
    Response.Write "<br>您指定的组件的检查结果："
      If Not hx.IsObjInstalled(strClass) then 
        Response.Write "<br><font color=red>很遗憾，该服务器不支持" & strclass & "组件！</font>"
      Else
        Response.Write "<br><font color=green>"
		Response.Write " 恭喜！该服务器支持" & strclass & "组件。"
		If hx.getver(strclass)<>"" then
		Response.Write " 该组件版本是：" & hx.getver(strclass)
		End if
		Response.Write "</font>"
      End If
      Response.Write "<br>"
    end if
	
	Response.Write "<p><a href="&hx.FileName&">返回</a></p>"
	Response.End
End Sub
Sub CalculateTest
%>
<br>
<%
End Sub
Sub DriveTest
	On Error Resume Next
	Dim fo,d,xTestObj
	set fo=Server.Createobject("Scripting.FileSystemObject")
	set xTestObj=fo.Drives
%>
<br>
<%
End Sub
Sub SpeedTest
Response.Flush()
%>
<a name="SpeedTest"></a>
<%	if action="SpeedTest" then%>
<div id="testspeed"> 
  <table width="200" border="0" cellspacing="0" cellpadding="0" class="divcenter">
    <tr> 
      <td height="30" align=center><p><font color="#376ed0"><span id=txt1>网速测试中，请稍候...</span></font></p></td>
    </tr>
  </table>
</div>
<%	end if%>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#376ed0" class="tableBorder">
  <tr> 
    <td height="25" align="center" bgcolor="#CCCCCC" onclick="showsubmenu(3)" background="images/b1.gif"><font color="#FFFFFF"><strong>服务器连接速度</strong></font> 
      <%smenu(3)%>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#F8F9FC" style="display" id='submenu3'> <table width="100%" border="0" cellspacing=1 cellpadding=3 bgcolor="#376ed0">
        <tr bgcolor="#FFFFFF"> 
          <td width="80">接入设备</td>
          <td width="420">&nbsp;连接速度(理想值)</td>
          <td width="100">下载速度(理想值)</td>
        </tr>
<tr bgcolor="#FFFFFF"> 
          <td>56k Modem</td>
          <td><img align=absmiddle class=PicBar width='1%'> 56 Kbps</td><td>&nbsp;7.0 k/s</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>64k ISDN</td>
          <td><img align=absmiddle class=PicBar width='1%'> 64 Kbps</td><td>&nbsp;8.0 k/s</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>512k ADSL</td>
          <td><img align=absmiddle class=PicBar width='5%'> 512 Kbps</td><td>&nbsp;64.0 k/s</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="19">1.5M Cable</td>
          <td><img align=absmiddle class=PicBar width='15%'> 1500 Kbps</td><td>&nbsp;187.5 k/s</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>5M FTTP</td>
          <td><img align=absmiddle class=PicBar width='50%'> 5000 Kbps</td><td>&nbsp;625.0 k/s</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>当前连接速度</td>
          <%
	if action="SpeedTest" then
		dim i
		Response.Write("<script language=""JavaScript"">var tSpeedStart=new Date();</script>")	
		Response.Write("<!--") & chr(13) & chr(10)
		for i=1 to 1000
		Response.Write("####################################################################################################") & chr(13) & chr(10)
		next
		Response.Write("-->") & chr(13) & chr(10)
		Response.Write("<script language=""JavaScript"">var tSpeedEnd=new Date();</script>") & chr(13) & chr(10)		
		Response.Write("<script language=""JavaScript"">")
		Response.Write("var iSpeedTime=0;iSpeedTime=(tSpeedEnd - tSpeedStart) / 1000;")
		Response.Write("if(iSpeedTime>0) iKbps=Math.round(Math.round(100 * 8 / iSpeedTime * 10.5) / 10); else iKbps=10000 ;")
		Response.Write("var iShowPer=Math.round(iKbps / 100);")		
		Response.Write("if(iShowPer<1) iShowPer=1;  else if(iShowPer>82)   iShowPer=82;")
		Response.Write("</script>") & chr(13) & chr(10)		
		Response.Write("<script language=""JavaScript"">") 
		Response.Write("document.write('<td><img align=absmiddle class=PicBar width=""' + iShowPer + '%""> ' + iKbps + ' Kbps');")
		Response.Write("</script>") & chr(13) & chr(10)
		Response.Write("</td><td>&nbsp;<a href='?action=SpeedTest' title=测试连接速度><u>")
		Response.Write("<script language=""JavaScript"">")
		Response.Write("document.write(Math.round(iKbps/8*10)/10+ ' k/s');")		
		Response.Write("</script>") & chr(13) & chr(10)				
		Response.Write 	"</u></a></td>"
%>
<script>
txt1.innerHTML="网速测试完毕!"
testspeed.style.visibility="hidden"
</script>
<%
	else
		Response.Write "<td></td><td>&nbsp;<a href='?action=SpeedTest' title=测试连接速度><u>开始测试</u></a></td>"
	end if
%>
        </tr>
      </table></td>
  </tr>
</table>
<%End Sub%>
</BODY>
</HTML>
