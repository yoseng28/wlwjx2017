<!-- #include file="conn.asp"-->
<%
if session("vote")<>"" then
response.write "<script>alert('���Ѿ�Ͷ��Ʊ�ˣ��벻Ҫ�ظ�ͶƱ');history.back();</script>"
response.end
end if
on error resume next
dim id,rs,v,chk,vv,vvv
id=request.form("id")
id=replace(id,"'","")
id=left(id,8)
v=request.form("v")
chk=request.form("chk")
if id<>"" and v<>"" then
select case chk
case "1" '��ѡ
conn.execute("update V_vote set vcount=vcount+1 where id=cint('"&v&"')")
case "2" '��ѡ
vvv=replace(v," ","")
vv=split(vvv,",")
for i= 0 to ubound(vv)
conn.execute("update V_vote set vcount=vcount+1 where id=cint('"&vv(i)&"')")
next
end select
session("vote")="ͶƱ"
response.write "<script>alert('ͶƱ�ɹ�');history.back();</script>"
response.end
else
response.write "<script>alert('ͶƱѡ���Ϊ��');history.back();</script>"
response.end
end if
%>