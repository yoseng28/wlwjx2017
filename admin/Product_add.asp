<!--#include file="adminconn.asp"-->
<!--#include file="Fckeditor/Fckeditor.Asp" -->
<%
Dim count,ranNum,iddata,Content
set rs=server.createobject("adodb.recordset")
sql = "select * from SmallClass order by SmallClassID asc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("SmallClassName"))%>","<%= trim(rs("BigClassName"))%>","<%= trim(rs("SmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.addNEWS.SmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.addNEWS.SmallClassName.options[document.addNEWS.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    

</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="renderer" content="webkit"> 
<meta name="renderer" content="ie-comp"> 
<meta name="renderer" content="ie-stand">
<link rel="stylesheet" href="images/style.css" type="text/css">
<title>�������</title>
</head>

<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>��</td>
  </tr>
</table>
<table width="780" border="1" align="center" cellpadding="2" cellspacing="1" bordercolor="#376ed0">
  <form name="addNEWS" method="post" action="Product_addok.asp">
    <tr> 
      <td height="25" colspan="3" align="center" background="images/b1.gif"><span class="style1"><font color="#FFFFFF">�� 
        �� �� ��</font></span> </td>
    </tr>
    <tr> 
      <td width="85" height="20" bgcolor="#f9f9f9"><font color="#FF0000">*</font>���ƣ�</td>
      <td colspan="2" bgcolor="#f9f9f9"> <input name="ProductName" type="text" class="input" id="ProductName" size="30">&nbsp; 
		�磺����&nbsp; ͼƬ���Ƶ�</td>
    </tr>
    <tr> 
      <td height="-2" bgcolor="#f9f9f9"><font color="#FF0000">*</font>ͼƬ���ͣ�</td>
      <td colspan="2" bgcolor="#f9f9f9"> 
        <%
        sql = "select * from BigClass where ClassType='product.asp' "
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "���������Ŀ��"
		else
		%>
        <select name="BigClassName" onChange="changelocation(document.addNEWS.BigClassName.options[document.addNEWS.BigClassName.selectedIndex].value)" size="1">
          <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
          <%
			dim selclass
		    selclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
          <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
          <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
        </select> <select name="SmallClassName">
          <option value="" selected>ѡ��С��</option>
          <%
			sql="select * from SmallClass where BigClassName='" & selclass & "'"
			rs.open sql,conn,1,1
			if not(rs.eof and rs.bof) then
			%>
          <option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
          <% rs.movenext
				do while not rs.eof%>
          <option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
          <%
			    	rs.movenext
				loop
			end if
	        rs.close
			%>
          <%
			ranNum=int(9*rnd)+10
			iddata=month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
			%>
        </select> </td>
    </tr>
     <tr> 
      <td height="-1" bgcolor="#f9f9f9"><font color="#FF0000">*</font>�ֶ�һ��</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductPrice5" type="text" id="ProductPrice5" size="30" value="0">
        �磺<b>��ʦ��飨����220��</b></td>
    </tr>

    <tr> 
      <td height="-1" bgcolor="#f9f9f9"><font color="#FF0000">*</font>�ֶζ���</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductType" type="text" id="ProductPrice" size="30" value="0">
        �磺ѧ��</td>
    </tr>
    <tr> 
      <td height="0" bgcolor="#f9f9f9"><font color="#FF0000">*</font>�ֶ�����</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductPrice1" type="text" id="ProductType" size="30" value="0"> 
      	�磺ѧλ</td>
    </tr>

    <tr> 
      <td bgcolor="#f9f9f9"><font color="#FF0000">*</font>�ֶ��ģ�</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductPrice2" type="text" id="ProductType0" size="30" value="0"> 
		�磺ְ��</td>
    </tr>
<tr> 
      <td bgcolor="#f9f9f9"><font color="#FF0000">*</font>�ֶ��壺</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductPrice3" type="text" id="ProductType0" size="30" value="0"> 
		�磺�����γ�</td>
    </tr>
<tr> 
      <td bgcolor="#f9f9f9"><font color="#FF0000">*</font>�ֶ�����</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductPrice4" type="text" id="ProductType0" size="30" value="0"> 
		�磺��Ƶ���ŵ�ַ</td>
    </tr>
<tr> 
      <td bgcolor="#f9f9f9">&nbsp;�ֶ��ߣ�</td>
      <td colspan="2" bgcolor="#f9f9f9">
		<input name="ProductPrice6" type="text" id="ProductType0" size="30" value="0"> 
		�磺����ר�ã�����д��� ���ԽС������ǰ</td>
    </tr>

    <tr> 
      <td height="20" bgcolor="#f9f9f9">����ͼƬ��ַ</td>
      <td width="191" bgcolor="#f9f9f9"> <input name="firstImageName" type="text" id="imgpath3" size="30"> 
      </td>
      <td width="480" bgcolor="#f9f9f9"> <iframe style="top:2px" id="UploadFiles" src="shangchuan_index.asp?picID=1" frameborder=0 scrolling=no width="320" height="25"></iframe>
		<b>��Ƶ�ϴ�Ҳ�ڴ˴���С��30M�����ϴ����������Ƶ���ŵ�ַ�����ֶ��������ϴ���Ƶ��ͼ</b></td>
    </tr>
    <tr> 
      <td height="25" valign="top" bgcolor="#f9f9f9"><font color="#FF0000">*</font>��ϸ���ܣ�</td>
      <td colspan="2" valign="top" bgcolor="#f9f9f9"> 
        <%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "FCKeditor/"
oFCKeditor.ToolbarSet = "bijie520"
oFCKeditor.Width = "100%"
oFCKeditor.Height = "400"
oFCKeditor.Value = Content
oFCKeditor.Create "d_content"
%>
      </td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#f9f9f9">�����ˣ�</td>
      <td colspan="2" bgcolor="#f9f9f9"> <input name="user" type="text" class="input" value="����Ա" size="30"></td>
    </tr>
    <tr> 
      <td height="25" bgcolor="#f9f9f9">��Ϊ��ҳ������</td>
      <td colspan="2" bgcolor="#f9f9f9"> <input type="radio" name="ok" value="true">
        �� 
        <input name="ok" type="radio" value="false" checked>
        �� ��</td>
    </tr>
    <tr> 
      <td height="30" colspan="3" align="center" bgcolor="#f9f9f9"> <input type="submit" name="Submit" value="�ύ" class="input">
        �� 
        <input type="reset" name="Submit2" value="����" class="input"> </td>
    </tr>
  </form>
</table>
</body>
</html>