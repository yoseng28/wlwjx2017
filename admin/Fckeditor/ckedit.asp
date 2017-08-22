<%
if replace(session("admin"),"'","")="" then
response.write"<div align=center><font color=red><strong><br><br><br><br>提示：非法操作！</strong></font></div>"
response.end 
end if
%>