<%
'--------���岿��------------------

'�Զ�����Ҫ���˵��ִ�,�� "|" �ָ�
Fy_In = "'|;|and|(|)|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare"
Kill_IP=True
WriteSql=True			
'----------------------------------


Fy_Inf = split(Fy_In,"|")
'--------POST����------------------
If Request.Form<>"" Then
	For Each Fy_Post In Request.Form
		For Fy_Xh=0 To Ubound(Fy_Inf)
			If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
				Response.Write"<script language=JavaScript>"
				Response.Write"alert(""��������"");"
				Response.Write"window.location='../index.asp'"
				Response.Write"</script>"
				Response.End
			End If
		Next
	Next
End If
'----------------------------------

'--------GET����-------------------
If Request.QueryString<>"" Then
	For Each Fy_Get In Request.QueryString
		For Fy_Xh=0 To Ubound(Fy_Inf)
			If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
				Response.Write"<script language=JavaScript>"
				Response.Write"alert(""��������"");"
				Response.Write"window.location='../index.asp'"
				Response.Write"</script>"
				Response.End
			End If
		Next
	Next
End If
%>