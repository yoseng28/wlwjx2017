<%
if replace(session("admin"),"'","")="" then
response.write"<div align=center><font color=red><strong><br><br><br><br>��ʾ���Ƿ�������</strong></font></div>"
response.end 
end if
%>