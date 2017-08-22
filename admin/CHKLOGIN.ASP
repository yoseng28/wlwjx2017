<!--#include file="../anquan/Check_SI.asp" -->
<!-- #include file=../inc/config.asp -->
<!--#include file="md5.asp" -->
<% if ipaddress<>"" then 
'本例根据远程主机地址来进行判断,如果为本地
'地址则进入欢迎页面,否则显示出错信息.
dim address
address = request.servervariables("REMOTE_ADDR")
if address<>ipaddress then 
    '否则显示出错信息.
   response.write "对不起,你无权查看该网站"

end if
end if
%>
<%
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
dim rs
 
admin1=request("adminname")
password1=request("adminpassword")
admin1=replace(admin1,"'","")
password1=replace(password1,"'","")
 
set rs=server.CreateObject("ADODB.RecordSet")
rs.open "select * from admin where admin='" & admin1 & "' and password='"&md5(password1,32)&"'",conn,1
if rs.eof and rs.bof then
response.write"<SCRIPT language=JavaScript>alert('用户名或密码不正确！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
else
if rs("aleave")="lookuser" then
response.write"<SCRIPT language=JavaScript>alert('对不起，你没有权限登陆，请和管理员联系！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
session("admin")=rs("admin")
session("password")=password1
session("aleave")=rs("aleave")
response.redirect "admin.asp"
end if
rs.close
set rs=nothing
%>
