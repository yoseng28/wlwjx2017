<!--#include file="adminconn.asp"-->

<%
dim Action,BigClassName,FoundErr,ErrMsg,ClassType,xianshi,xuehao,rsabout,xsfs,linkfs,wblj
Action=trim(Request("Action"))
BigClassName=trim(request("BigClassName"))
ClassType=trim(request("ClassType"))
xianshi=trim(request("xianshi"))
xuehao=request("xuehao")
xsfs=request("xsfs")
linkfs=trim(request("linkfs"))
wblj=trim(request("wblj"))
if wblj="��" and linkfs="" then
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('��ѡ������ⲿ���ӣ���δ�����ⲿ���ӵ�ַ�������������ⲿ��ַ��');" & Chr(13)
		response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
end if
if Action="Add" then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From BigClass Where BigClassName='" & BigClassName & "'",conn,1,3
		if not (rs.bof and rs.EOF) then
		%>
<script language="JavaScript" type="text/JavaScript">

    alert("���������Ѿ����ڣ����������룡");

</script>
<%
			
		else
    	 	rs.addnew
     		rs("BigClassName")=BigClassName
			rs("ClassType")=ClassType	
			rs("xianshi")=xianshi		
			rs("xuehao")=xuehao
			rs("xsfs")=xsfs
			rs("linkfs")=linkfs
			if wblj="��" then rs("wblj")=true
    	 	rs.update
     		rs.Close
	     	set rs=Nothing
			
		if ClassType="about.asp" then
		Set rsabout=Server.CreateObject("Adodb.RecordSet")
		rsabout.open "Select * From about Where title='" & BigClassName & "'",conn,1,3
		if not (rsabout.bof and rsabout.EOF) then
		%>
<script language="JavaScript" type="text/JavaScript">

    alert("���������Ѿ����ڣ����������룡");

</script>
<%
			
		else
    	 	rsabout.addnew
     		rsabout("title")=BigClassName
			rsabout("content")=BigClassName
			rsabout("abouttype")=1
    	 	rsabout.update
     		rsabout.Close
	     	set rsabout=Nothing
    	  
	end if
	end if
	call CloseConn()
	Response.Redirect "ClassManage.asp" 					
	end if
	end if
%>

<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("�������Ʋ���Ϊ�գ�");
    document.form1.BigClassName.focus();	
    return false;
  }  
}
</script>
<LINK href="images/style.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
<table width="794" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="794" align="center" valign="top"> <br>
      <a href="ClassManage.asp"><strong>�� �� �� ��</strong></a> <br>
      <form name="form1" method="post" action="ClassAddBig.asp" onsubmit="return checkBig()">
        <table width="786" border="1" align="center" cellpadding="4" cellspacing="1" bordercolor="#376ed0" class="border">
          <tr bgcolor="#999999" class="title"> 
            <td height="25" colspan="3" align="center" background="images/b1.gif"><strong>��Ӵ���</strong></td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td width="158" height="9"> 
              <div align="right"><strong>�� �� �� �� ��</strong></div></td>
            <td width="150"> 
              <input name="BigClassName" type="text" size="20" maxlength="30"> 
            </td>
            <td width="442">��</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td width="158" height="4"> 
              <div align="right"><strong>�� Ŀ �� �� ��</strong></div></td>
            <td bgcolor="#f9f9f9"> <select name="ClassType" id="ClassType" onchange="dis(document.form1.ClassType.value)">
                <option value="class.asp" selected>��ͨҳ��</option>
                <option value="about.asp">������Ϣ����</option>
                <option value="product.asp">ͼƬ����</option>
				<option value="download.asp">�������</option>
                <option value="guestbook.asp">�������</option>
                <option value="rencai.asp">��Ƹҳ��</option>
              </select> </td>
            <td bgcolor="#f9f9f9">��</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg" id="xsfs" style='display:none'> 
            <td width="158" height="1"> 
              <div align="right"><strong>�� ʾ ��ʽ ��</strong></div></td>
            <td><select name="xsfs" size="1" id="select2">
                <option value="ͼƬ��ʽ" selected>ͼƬ��ʽ</option>
                <option value="ͼ���б�">ͼ���б�</option>
              </select></td>
            <td><font color="#FF0000">ֻ��ͼƬģ�� ����Ч</font></td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td width="158" height="2"> 
              <div align="right"><strong>��ҳ��Ŀ��ʾ��</strong></div></td>
            <td> <select name="xianshi" size="1" id="select">
                <option selected value="ok">��ʾ</option>
                <option value="no">����ʾ</option>
              </select></td>
            <td>��ҳ�����Ƿ���ʾ</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td height="22" align="center"><div align="right"><strong>�� ʾ ˳ ��</strong></div></td>
            <td height="22" align="center"> <div align="left"> 
                <input name="xuehao" type="text" id="xuehao3" size="10" maxlength="5">
                <span class="style1">ֻ��������</span></div></td>
            <td bgcolor="#f9f9f9">����ԽСԽ��ǰ��ʾ</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"> 
            <td height="22" align="center"><div align="right"><strong>�Ƿ��ⲿ���ӣ�</strong></div></td>
            <td height="22" bgcolor="#f9f9f9"><select name="wblj" size="1" id="wblj"  onchange="wdis(document.form1.wblj.value)">
                <option value="��" selected>��</option>
                <option value="��">��</option>
              </select> </td>
            <td bgcolor="#f9f9f9">���ѡ���ǣ�����������ⲿ���ӵ�ַ��</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg"  id="wbljdz" style='display:none'> 
            <td height="22" align="right" bgcolor="#f9f9f9"><strong>�ⲿ���ӵ�ַ��</strong></td>
            <td height="22" bgcolor="#f9f9f9"> 
              <input name="linkfs" type="text" id="linkfs2" size="25" maxlength="100"> 
            </td>
            <td bgcolor="#f9f9f9">ע���ַ�б��뺬�� <font color="#FF0000">http://</font>,�磺<font color="#FF0000">http://www.lzkeyuan.cn</font> 
              ��</td>
          </tr>
          <tr bgcolor="#f9f9f9" class="tdbg">
            <td height="22" align="center">��</td>
            <td height="22" align="center"> <input name="Action" type="hidden" id="Action2" value="Add"> 
              <input name="Add" type="submit" value=" �� �� "></td>
            <td align="center">��</td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<script language="JavaScript">
<!--
function dis(vid)
{
   if(vid=="product.asp")
   {
   eval("xsfs.style.display=\"\";");
   form1.class_cell.value=1; 
   }
   else
   {
    eval("xsfs.style.display=\"none\";");
   form1.class_cell.value=0; 
   }
}
function wdis(vid)
{
   if(vid=="��")
   {
   eval("wbljdz.style.display=\"\";");
   form1.class_cell.value=1; 
   }
   else
   {
    eval("wbljdz.style.display=\"none\";");
   form1.class_cell.value=0; 
   }
}
//-->
</SCRIPT>
