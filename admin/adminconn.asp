<%@ Language=VBScript CODEPAGE=936%>
<% Option Explicit %>
<%
' ============================================
' 常用全局变量
' ============================================
Dim company, sitetitle, gongsurl,weblogo,weblogotype,selectclass,index1,index2,index3,index4,index5,index6,index7,index8,index9,reuserpassword,dataurl,backurl,admindataurl,address,code,companycon,phone,fax,email,reply,companycontent,companyldao,bankid,bankuser,bankname,ipaddress,beianhao,qq1,qq2,qq3,kehu,webDescription,webKeywords,chengxuup
%>
<!-- #include file=../inc/config.asp -->
<!-- #include file=function.asp -->
<!--#include file="md5.asp" -->
<%
'==================================================================
'打开数据库地址
dim StrSQL,conn
StrSQL="DBQ="+server.mappath(""&admindataurl&"")+";DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
conn.open StrSQL
sub CloseConn()
	conn.close
	set conn=nothing
end sub
function decrypt(dcode)	
dim texts
dim i
for i=1 to len(dcode)
texts=texts & chr(asc(mid(dcode,i,2))-i)
next
decrypt=texts
end function
function encrypt(ecode)
Dim texts
dim i
for i=1 to len(ecode)
texts=texts & chr(asc(mid(ecode,i,2))+i)
next
encrypt = texts
end function
%>
<% 
'=================================================
'本例根据远程主机地址来进行判断,如果为本地
'地址则进入欢迎页面,否则显示出错信息.
Dim ipaddress1
if ipaddress<>"" then 
ipaddress1 = request.servervariables("REMOTE_ADDR")
if ipaddress1<>ipaddress then 
    '否则显示出错信息.
   	response.write "<br><p align=center><font color='red'>对不起，你的IP地址不能访问该网站。</font></p>"
	response.end

end if
end if
%>

<%
'判断是否登陆和来源
Dim ComeUrl,cUrl,AdminName

ComeUrl=lcase(trim(request.ServerVariables("HTTP_REFERER")))
if ComeUrl="" then
	response.write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。</font></p>"
	response.end
else
	cUrl=trim("http://" & Request.ServerVariables("SERVER_NAME"))
	if mid(ComeUrl,len(cUrl)+1,1)=":" then
		cUrl=cUrl & ":" & Request.ServerVariables("SERVER_PORT")
	end if
	cUrl=lcase(cUrl & request.ServerVariables("SCRIPT_NAME"))
	if lcase(left(ComeUrl,instrrev(ComeUrl,"/")))<>lcase(left(cUrl,instrrev(cUrl,"/"))) then
		response.write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许从外部链接地址访问本系统的后台管理页面。</font></p>"
		response.end
	end if
end if

Dim admin,sql,rs
admin=replace(session("admin"),"'","")
reuserpassword=replace(session("password"),"'","")
if admin="" then
	call CloseConn()
	response.redirect "adminlogin.asp"
	response.End()
end if
sql="select admin from Admin where admin='" & session("admin") & "' and password='"&md5(reuserpassword,32)&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
  rs.close
  response.Redirect("adminlogin.asp")
  response.End()
end if
%>