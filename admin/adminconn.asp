<%@ Language=VBScript CODEPAGE=936%>
<% Option Explicit %>
<%
' ============================================
' ����ȫ�ֱ���
' ============================================
Dim company, sitetitle, gongsurl,weblogo,weblogotype,selectclass,index1,index2,index3,index4,index5,index6,index7,index8,index9,reuserpassword,dataurl,backurl,admindataurl,address,code,companycon,phone,fax,email,reply,companycontent,companyldao,bankid,bankuser,bankname,ipaddress,beianhao,qq1,qq2,qq3,kehu,webDescription,webKeywords,chengxuup
%>
<!-- #include file=../inc/config.asp -->
<!-- #include file=function.asp -->
<!--#include file="md5.asp" -->
<%
'==================================================================
'�����ݿ��ַ
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
'��������Զ��������ַ�������ж�,���Ϊ����
'��ַ����뻶ӭҳ��,������ʾ������Ϣ.
Dim ipaddress1
if ipaddress<>"" then 
ipaddress1 = request.servervariables("REMOTE_ADDR")
if ipaddress1<>ipaddress then 
    '������ʾ������Ϣ.
   	response.write "<br><p align=center><font color='red'>�Բ������IP��ַ���ܷ��ʸ���վ��</font></p>"
	response.end

end if
end if
%>

<%
'�ж��Ƿ��½����Դ
Dim ComeUrl,cUrl,AdminName

ComeUrl=lcase(trim(request.ServerVariables("HTTP_REFERER")))
if ComeUrl="" then
	response.write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ��������ֱ�������ַ���ʱ�ϵͳ�ĺ�̨����ҳ�档</font></p>"
	response.end
else
	cUrl=trim("http://" & Request.ServerVariables("SERVER_NAME"))
	if mid(ComeUrl,len(cUrl)+1,1)=":" then
		cUrl=cUrl & ":" & Request.ServerVariables("SERVER_PORT")
	end if
	cUrl=lcase(cUrl & request.ServerVariables("SCRIPT_NAME"))
	if lcase(left(ComeUrl,instrrev(ComeUrl,"/")))<>lcase(left(cUrl,instrrev(cUrl,"/"))) then
		response.write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ����������ⲿ���ӵ�ַ���ʱ�ϵͳ�ĺ�̨����ҳ�档</font></p>"
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