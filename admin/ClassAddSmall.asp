<!--#include file="adminconn.asp"-->
<%
dim Action,BigClassName,SmallClassName,FoundErr,ErrMsg
Action=trim(Request("Action"))
BigClassName=trim(request("BigClassName"))
SmallClassName=trim(request("SmallClassName"))
if Action="Add" then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From SmallClass Where BigClassName='" & BigClassName & "' AND SmallClassName='" & SmallClassName & "'",conn,1,3
		if not rs.EOF then
		%>
<script language="JavaScript" type="text/JavaScript">

    alert("С�������ڸô������Ѿ����ڣ����������룡");

</script>
<%
else
     		rs.addnew
			rs("BigClassName")=BigClassName
    	 	rs("SmallClassName")=SmallClassName
     		rs.update
	     	rs.Close
    	 	set rs=Nothing
     		call CloseConn()
			Response.Redirect "ClassManage.asp"  
		end if
	end if
%>
<script language="JavaScript" type="text/JavaScript">
function checkSmall()
{
  if (document.form2.BigClassName.value=="")
  {
    alert("������Ӵ������ƣ�");
	document.form1.BigClassName.focus();
	return false;
  }

  if (document.form2.SmallClassName.value=="")
  {
    alert("С�����Ʋ���Ϊ�գ�");
	document.form2.SmallClassName.focus();
	return false;
  }
}
</script>
<LINK href="images/style.css" type=text/css rel=stylesheet>
<table width="80%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br>
      <a href="ClassManage.asp"><strong>�� �� �� ��</strong></a> <br> <form name="form2" method="post" action="ClassAddSmall.asp" onsubmit="return checkSmall()">
        <table width="350" border="1" align="center" cellpadding="4" cellspacing="1" bordercolor="#376ed0" class="border">
          <tr bgcolor="#999999" class="title"> 
            <td height="25" colspan="2" align="center" background="images/b1.gif"><strong>���С��</strong></td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td width="126" height="22"><strong>�������ࣺ</strong></td>
            <td width="218"> <strong> 
              <select name="BigClassName">
                <%
	dim rsBigClass
	set rsBigClass=server.CreateObject("adodb.recordset")
	rsBigClass.open "Select * From BigClass",conn,1,1
	if rsBigClass.bof and rsBigClass.bof then
		response.write "<option>����������´���</option>"
	else
		do while not rsBigClass.eof
			if rsBigClass("BigClassName")=BigClassName then
				response.write "<option value='"& rsBigClass("BigClassName") & "' selected>" & rsBigClass("BigClassName") & "</option>"
			else
				response.write "<option value='"& rsBigClass("BigClassName") & "'>" & rsBigClass("BigClassName") & "</option>"
			end if
			rsBigClass.movenext
		loop
	end if
	rsBigClass.close
	set rsBigClass=nothing
	%>
              </select>
              </strong></td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td width="126" height="22"><strong>С�����ƣ�</strong></td>
            <td> <strong> 
              <input name="SmallClassName" type="text" size="20" maxlength="30">
              </strong></td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td height="22" align="center">&nbsp; </td>
            <td height="22" align="center" bgcolor="#f9f9f9"> 
              <div align="left"> 
                <strong> 
                <input name="Action" type="hidden" id="Action3" value="Add">
                <input name="Add" type="submit" value=" �� �� ">
                </strong></div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
