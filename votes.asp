<!-- #include file="conn.asp"-->
<%
if session("vote")<>"" then
response.write "<script>alert('您已经投过票了，请不要重复投票');history.back();</script>"
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
case "1" '单选
conn.execute("update V_vote set vcount=vcount+1 where id=cint('"&v&"')")
case "2" '多选
vvv=replace(v," ","")
vv=split(vvv,",")
for i= 0 to ubound(vv)
conn.execute("update V_vote set vcount=vcount+1 where id=cint('"&vv(i)&"')")
next
end select
session("vote")="投票"
response.write "<script>alert('投票成功');history.back();</script>"
response.end
else
response.write "<script>alert('投票选项不能为空');history.back();</script>"
response.end
end if
%>