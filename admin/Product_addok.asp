<!--#include file="adminconn.asp"-->
<%
Dim ProductName,ProductBigClass,ProductSmallClass,ProductPrice,ProductPrice1,ProductPrice2,ProductPrice3,ProductPrice4,ProductPrice5,ProductPrice6,ProductType,ProductImages,ProductUser,ok,sContent,i
ProductName=request("ProductName")
ProductBigClass=request("BigClassName")
ProductSmallClass=request("SmallClassName")
ProductPrice=request("ProductPrice")
ProductPrice1=request("ProductPrice1")
ProductPrice2=request("ProductPrice2")
ProductPrice3=request("ProductPrice3")
ProductPrice4=request("ProductPrice4")
ProductPrice5=request("ProductPrice5")
ProductPrice6=request("ProductPrice6")
ProductType=request("ProductType")
ProductImages=request("firstImageName")
ProductUser=request("user")
ok=request("ok")
	sContent = ""
	For i = 1 To Request.Form("d_content").Count
		sContent = sContent & Request.Form("d_content")(i)
	Next

'�ж����������Ƿ�����
if ProductName="" or ProductBigClass="" or ProductPrice1="" or ProductType="" or sContent="" then  
response.write"<div align=center><font color=red size=+3><br><br><strong>�������ˡ�</strong></font>"
response.write"<br><br><font color=green>�����������,û����ӳɹ���</font>"
response.write"<br><br>����<font color=red>*</font>�ĵط��Ƿ�ȫ����д��"
response.write"<br><br><a href=javascript:history.back()>���ؼ�����</a></div>"  
response.end 
end if

set rs=server.createobject("adodb.recordset")
sql="select * from Product where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("ProductName")=ProductName
rs("ProductPrice")=ProductPrice
rs("ProductPrice1")=ProductPrice1
rs("ProductPrice2")=ProductPrice2
rs("ProductPrice3")=ProductPrice3
rs("ProductPrice4")=ProductPrice4
rs("ProductPrice5")=ProductPrice5
rs("ProductPrice6")=ProductPrice6
rs("ProductBigClass")=ProductBigClass
rs("ProductSmallClass")=ProductSmallClass
rs("ProductType")=ProductType
rs("ProductImages")=ProductImages
rs("ProductUser")=ProductUser
if ok<>"" then rs("ok") = ok
rs("ProductContent")=sContent
rs.update
rs.close
set rs=nothing
conn.close  
set conn=nothing
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('�ύ�ɹ���');" & Chr(13)
		response.write "window.document.location.href='product_info.asp';"&Chr(13)
		response.write "</script>" & Chr(13)
Response.End
%>