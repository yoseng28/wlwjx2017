<!--#include file="adminconn.asp"-->
<%
dim SmallClassID,Action,BigClassName, SmallClassName, OldSmallClassName,FoundErr,ErrMsg
SmallClassID=trim(Request("SmallClassID"))
Action=trim(Request("Action"))
BigClassName=trim(Request.form("BigClassName"))
SmallClassName=trim(Request.form("SmallClassName"))
OldSmallClassName=trim(request.form("OldSmallClassName"))
if SmallClassID="" then
	response.Redirect("ClassManage.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from SmallClass where SmallClassID="&SmallClassID&"",conn,1,3
if rs.Bof or rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>������С�಻���ڣ�</li>"
else
	if Action="Modify" then
		if SmallClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>����С��������Ϊ�գ�</li>"
		end if
		if FoundErr<>True then
			rs("SmallClassName")=SmallClassName
     		rs.update
			rs.Close
    	 	set rs=Nothing
			if SmallClassName<>OldSmallClassName then
				conn.execute "update content set SmallClassName='" & SmallClassName & "' where BigClassName='" & BigClassName & "' and SmallClassName='" & OldSmallClassName & "'"
	     	end if	
			call CloseConn()
    	 	Response.Redirect "ClassManage.asp"
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<LINK href="images/style.css" type=text/css rel=stylesheet>
<table width="80%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br>
      <a href="ClassManage.asp"><strong>�� �� �� ��</strong></a> <br> <form name="form1" method="post" action="ClassModifySmall.asp">
        <p>&nbsp;</p>
        <table width="350" border="1" align="center" cellpadding="4" cellspacing="1" bordercolor="#376ed0" class="border">
          <tr bgcolor="#999999" class="title"> 
            <td height="25" colspan="2" align="center" background="images/b1.gif"><strong>�޸�С������</strong></td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td width="126" height="22"><strong>�������ࣺ</strong></td>
            <td width="204"><%=rs("BigClassName")%> 
              <input name="SmallClassID" type="hidden" id="SmallClassID" value="<%=rs("SmallClassID")%>"> 
              <input name="OldSmallClassName" type="hidden" id="OldSmallClassName" value="<%=rs("SmallClassName")%>"> 
              <input name="BigClassName" type="hidden" id="BigClassName" value="<%=rs("BigClassName")%>"></td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td height="22"><strong>С�����ƣ�</strong></td>
            <td> 
              <input name="SmallClassName" type="text" id="SmallClassName" value="<%=rs("SmallClassName")%>" size="20" maxlength="30"></td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td height="22" align="center">&nbsp;</td>
            <td align="center"> 
              <div align="left"> 
                <input name="Action" type="hidden" id="Action4" value="Modify">
                <input name="Save" type="submit" id="Save" value=" �� �� ">
              </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<%
end if
end if
rs.close
set rs=nothing
%>
