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

'�������Ƿ�֧��
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

'�������汾
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
      	Outstr = Outstr & "<td align=left>&nbsp;<font color=red><b>��</b></font></td>"
    	Else
      	Outstr = Outstr & "<td align=left>&nbsp;<font color=green><b>��</b></font> " & getver(theTestObj(i,0)) & "</td>"
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
	OutStr = OutStr & "<br><p>Copyright 2006 ��վ����ϵͳ</p>"& vbcrlf
 	Runtime=FormatNumber((endtime-startime)*1000,2) 
	if Runtime>0 then
		if Runtime>1000 then
		  	OutStr = OutStr & "ҳ��ִ��ʱ�䣺Լ"& FormatNumber(runtime/1000,2) & "��"
		else
			OutStr = OutStr & "ҳ��ִ��ʱ�䣺Լ"& Runtime & "����"
		end if	
	end if				
	OutStr = OutStr & "</p></td></tr></table>"
	Response.Write(OutStr)
End Sub
End class

Dim theTestObj(25,1)

	theTestObj(0,0) = "Scripting.FileSystemObject"
	theTestObj(0,1) = "(FSO �ı��ļ���д)"
	theTestObj(1,0) = "ADODB.Connection"
	theTestObj(1,1) = "(ADO ���ݶ���)"  
	theTestObj(2,0) = "JMail.SmtpMail"
	theTestObj(2,1) = "(Dimac JMail �ʼ��շ�)</a>"
	theTestObj(3,0) = "CDONTS.NewMail"
	theTestObj(3,1) = "(���� SMTP ����)"
	theTestObj(4,0) = "SoftArtisans.ImageGen"
	theTestObj(4,1) = "(SA ��ͼ���д���)"
	theTestObj(5,0) = "W3Image.Image"
	theTestObj(5,1) = "(Dimac ��ͼ���д���)"


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
			return "��������֧�ִ�����";
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
	alert("��������Ҫ�����������");
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
eval("txt" + sid + ".innerHTML=\"<a href='#' title='�رմ���'><font face='Wingdings' color=#FFFFFF>x</font></a>\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
eval("txt" + sid + ".innerHTML=\"<a href='#' title='�򿪴���'><font face='Wingdings' color=#FFFFFF>y</font></a>\";");
}
}
-->
</SCRIPT>
</HEAD>
<BODY leftmargin="50">
<table width="600" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="50" align="center"> <p><font color="#0099FF" size="6"><strong>��������Ϣ 
        <a href="../help/index.htm" target="_blank"><font color="#FF0000">����鿴�����ĵ�</font></a></strong></font></p>
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
ѡ�<a href="#SystemTest">�������йز���</a> | <a href="#ObjTest">������������</a> | <a href="#SpeedTest">�����������ٶ�</a> 
<%End Sub%>
<%Sub smenu(i)%>
<a href="#top" title="���ض���"><font face='Webdings' color=#FFFFFF>5</font></a> <span id=txt<%=i%> name=txt<%=i%>><a href='#' title='�رմ���'><font face='Wingdings' color=#FFFFFF>x</font></a></span> 
<%End Sub%>
<%Sub SystemTest
on error resume next
%>
<a name="SystemTest"></a> 
<table width="90%" border="0" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr> 
    <td height="25" align="center" bgcolor="#CCCCCC" onclick="showsubmenu(0)" background="images/b1.gif"><font color=#FFFFFF><strong>�������йز���</strong></font> 
      <%Call smenu(0)%></td>
  </tr>
  <tr> 
    <td style="display" id='submenu0'><table border=0 width=100% cellspacing=1 cellpadding=3 bgcolor="#376ed0">
        <tr bgcolor="#FFFFFF" height=18> 
          <td width="130">&nbsp;��������</td>
          <td width="170" height="18">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
          <td width="130" height="18">&nbsp;����������ϵͳ</td>
          <td width="170" height="18">&nbsp;<%=Request.ServerVariables("OS")%></td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td>&nbsp;������IP</td>
          <td>&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
          <td>&nbsp;�������˿�</td>
          <td>&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td>&nbsp;������ʱ��</td>
          <td>&nbsp;<%=now%></td>
          <td>&nbsp;������CPU����</td>
          <td>&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 
            ��</td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td>&nbsp;IIS�汾</td>
          <td height="18">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
          <td height="18">&nbsp;�ű���ʱʱ��</td>
          <td height="18">&nbsp;<%=Server.ScriptTimeout%> ��</td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td>&nbsp;Application����</td>
          <td height="18">&nbsp;<%Response.Write(Application.Contents.Count & "�� ")
		  if Application.Contents.count>0 then Response.Write("[<a href=""?action=showapp"">����Application����</a>]")%>
          </td>
          <td height="18">&nbsp;Session����<br> </td>
          <td height="18">&nbsp;<%Response.Write(Session.Contents.Count&"�� ")
		  if Session.Contents.count>0 then Response.Write("[<a href=""?action=showsession"">����Session����</a>]")%>
          </td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td height="18">&nbsp;<a href="?action=showvariables">���з���������</a></td>
          <td height="18">&nbsp;<%Response.Write(Request.ServerVariables.Count&"�� ")
		  if Request.ServerVariables.Count>0 then Response.Write("[<a href=""?action=showvariables"">��������������</a>]")%>
          </td>
          <td height="18">&nbsp;��������������</td>
          <td height="18">&nbsp;<%
			dim WshShell,WshSysEnv
			Set WshShell = server.CreateObject("WScript.Shell")
			Set WshSysEnv = WshShell.Environment
			if err then
				Response.Write("��������֧��WScript.Shell���")
				err.clear
			else
				Response.Write(WshSysEnv.count &"�� ")
				if WshSysEnv.count>0 then Response.Write("[<a href=""?action=showwsh"">������������</a>]") 
		 	end if
		  %>		  
		  </td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td align=left>&nbsp;��������������</td>
          <td height="18" colspan="3">&nbsp;JScript: <%= getEngVerJs() %> | VBScript: <%=getEngVerVBS()%></td>
        </tr>
        <tr bgcolor="#FFFFFF" height=18> 
          <td align=left bgcolor="#FFFFFF">&nbsp;���ļ�ʵ��·��</td>
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
		Response.Write("<font face='Webdings'>4</font> ����Application����")
		set xTestObj=Application.Contents
	elseif action="showsession" then
		Response.Write("<font face='Webdings'>4</font> ����Session����")
		set xTestObj=Session.Contents
	elseif action="showvariables" then
		Response.Write("<font face='Webdings'>4</font> ��������������")
		set xTestObj=Request.ServerVariables
	elseif action="showwsh" then
		Response.Write("<font face='Webdings'>4</font> ������������")
		dim WshShell
		Set WshShell = server.CreateObject("WScript.Shell")
		set xTestObj=WshShell.Environment
	end if
		Response.Write "(<a href="&hx.FileName&">�ر�</a>)"
	%>
    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="130">������</td>
    <td width="470">ֵ</td>
  </tr>
  <%
	if err then
		outstr = "<tr bgcolor=#FFFFFF><td colspan=2>û�з��������ı���</td></tr>"
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
    <td height="25" align="center" bgcolor="#CCCCCC" onclick="showsubmenu(1)" background="images/b1.gif"><strong><font color="#FFFFFF">������</font></strong><font color=#FFFFFF><strong>������</strong></font> 
      <%Call smenu(1)%>
    </td>
  </tr>
  <tr> 
    <td style="display" id='submenu1'><table border=0 width=100% cellspacing=1 cellpadding=3 bgcolor="#376ed0">

        <tr bgcolor="#FFFFFF" height=18> 
          <td width=450 align="center">�� �� �� ��</td>
          <td width=150 align="center">֧�ּ��汾</td>
        </tr>
        <%hx.GetObjInfo 0,5%>
      </table>
      <table border=0 width=100% cellspacing=1 cellpadding=3  bgcolor="#376ed0">
        <tr> 
          <td height="25" background="images/b1.gif" bgcolor="#CCCCCC">&nbsp;&nbsp;<font face='Webdings'>4</font> 
            �������֧�������� </td>
        </tr>
        <FORM action=?action=testzujian method=post id=form1 name=form1 onSubmit="JavaScript:return Checksearchbox();">
          <tr> 
            <td height=30 bgcolor="#FFFFFF">������Ҫ���������ProgId��ClassId 
              <input class=input type=text value="" name="classname" size=40> 
              <INPUT type=submit value="ȷ��" class=backc id=submit1 name=submit1> 
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
	
    Response.Write "<br>��ָ��������ļ������"
      If Not hx.IsObjInstalled(strClass) then 
        Response.Write "<br><font color=red>���ź����÷�������֧��" & strclass & "�����</font>"
      Else
        Response.Write "<br><font color=green>"
		Response.Write " ��ϲ���÷�����֧��" & strclass & "�����"
		If hx.getver(strclass)<>"" then
		Response.Write " ������汾�ǣ�" & hx.getver(strclass)
		End if
		Response.Write "</font>"
      End If
      Response.Write "<br>"
    end if
	
	Response.Write "<p><a href="&hx.FileName&">����</a></p>"
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
      <td height="30" align=center><p><font color="#376ed0"><span id=txt1>���ٲ����У����Ժ�...</span></font></p></td>
    </tr>
  </table>
</div>
<%	end if%>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#376ed0" class="tableBorder">
  <tr> 
    <td height="25" align="center" bgcolor="#CCCCCC" onclick="showsubmenu(3)" background="images/b1.gif"><font color="#FFFFFF"><strong>�����������ٶ�</strong></font> 
      <%smenu(3)%>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#F8F9FC" style="display" id='submenu3'> <table width="100%" border="0" cellspacing=1 cellpadding=3 bgcolor="#376ed0">
        <tr bgcolor="#FFFFFF"> 
          <td width="80">�����豸</td>
          <td width="420">&nbsp;�����ٶ�(����ֵ)</td>
          <td width="100">�����ٶ�(����ֵ)</td>
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
          <td>��ǰ�����ٶ�</td>
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
		Response.Write("</td><td>&nbsp;<a href='?action=SpeedTest' title=���������ٶ�><u>")
		Response.Write("<script language=""JavaScript"">")
		Response.Write("document.write(Math.round(iKbps/8*10)/10+ ' k/s');")		
		Response.Write("</script>") & chr(13) & chr(10)				
		Response.Write 	"</u></a></td>"
%>
<script>
txt1.innerHTML="���ٲ������!"
testspeed.style.visibility="hidden"
</script>
<%
	else
		Response.Write "<td></td><td>&nbsp;<a href='?action=SpeedTest' title=���������ٶ�><u>��ʼ����</u></a></td>"
	end if
%>
        </tr>
      </table></td>
  </tr>
</table>
<%End Sub%>
</BODY>
</HTML>
