<!-- #include file=inc/config.asp -->
<%
dim conn,connstr
on error resume next
connstr="DBQ="+server.mappath(""&dataurl&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open connstr
%>
<%
Function check_request(StrRequest)
check_request=false
if StrRequest<>"" then
   if instr(StrRequest,"'")>=1 or instr(StrRequest,";")>=1 or instr(StrRequest,".")>=1 or instr(StrRequest,",")>=1 or instr(StrRequest,"+")>=1 or instr(StrRequest,"=")>=1 or instr(StrRequest,"-")>=1 or instr(StrRequest,"(")>=1 or instr(StrRequest,")")>=1 then
      check_request=true
      exit function
   end if
end if
end function
%>