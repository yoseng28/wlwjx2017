<%
' ִ��ÿ��ֻ�账��һ�ε��¼�
Call BrandNewDay()

' ��ʼ�����ݿ�����

' ���ñ���
Dim sAction, sPosition
sAction = UCase(Trim(Request.QueryString("action")))
sPosition = "λ�ã�<a href='admin_default.asp'>��̨����</a> / "


' ********************************************
' ����Ϊҳ�湫��������
' ********************************************
' ============================================
' ���ÿҳ���õĶ�������
' ============================================
Sub Header()
	Response.Write "<html><head>"
	
	' ��� meta ���
	Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & _
		"<meta name='Author' content='Andy.m'>" 
	
	' �������
	Response.Write "<title>�ļ�����</title>"
	
    ' ���ÿҳ��ʹ�õĻ�����ʽ��
	Response.Write "<link rel='stylesheet' type='text/css' href='images/style.css'>"

	' ���ÿҳ��ʹ�õĻ����ͻ��˽ű�
	Response.Write "<script language='javaScript' SRC='Edit_private.js'></SCRIPT>"
	
	Response.Write "</head>"

	Response.Write "<body topmargin=0 leftmargin=0 bgcolor=#ffffff>"
	Response.Write "<a name=top></a>"
	
	' ���ҳ�涥��(Header1)


	' ҳ���������
	Response.Write "<table border=0 cellpadding=0 cellspacing=0 width=100% align=center>" & _
		"<tr valign=top>" 
	' ��̨����ģ��
	Call Block_Member()
	Response.Write "<td>" 


End Sub

' ============================================
' ���ÿҳ���õĵײ�����
' ============================================
Sub Footer()
	' �ͷ��������Ӷ���
	

	Response.Write "</td>"
	Response.Write "<td width=7></td><td width=1 class=rightbg></td>"
	Response.Write "</tr></table>"
	
	' �ײ�����

	


	Response.Write "</body></html>"
End Sub


' ��̨����ģ��
Sub Block_Member()

End Sub

' ===============================================
' ��ʼ��������
'	s_FieldName	: ���ص���������	
'	a_Name		: ��ֵ������
'	a_Value		: ��ֵֵ����
'	v_InitValue	: ��ʼֵ
'	s_Sql		: �����ݿ���ȡֵʱ,select name,value from table
'	s_AllName	: ��ֵ������,��:"ȫ��","����","Ĭ��"
' ===============================================
Function InitSelect(s_FieldName, a_Name, a_Value, v_InitValue, s_Sql, s_AllName)
	Dim i
	InitSelect = "<select name='" & s_FieldName & "' size=1>"
	If s_AllName <> "" Then
		InitSelect = InitSelect & "<option value=''>" & s_AllName & "</option>"
	End If
	If s_Sql <> "" Then
		oRs.Open s_Sql, oConn, 0, 1
		Do While Not oRs.Eof
			InitSelect = InitSelect & "<option value=""" & inHTML(oRs(1)) & """"
			If oRs(1) = v_InitValue Then
				InitSelect = InitSelect & " selected"
			End If
			InitSelect = InitSelect & ">" & outHTML(oRs(0)) & "</option>"
			oRs.MoveNext
		Loop
		oRs.Close
	Else
		For i = 0 To UBound(a_Name)
			InitSelect = InitSelect & "<option value=""" & inHTML(a_Value(i)) & """"
			If a_Value(i) = v_InitValue Then
				InitSelect = InitSelect & " selected"
			End If
			InitSelect = InitSelect & ">" & outHTML(a_Name(i)) & "</option>"
		Next
	End If
	InitSelect = InitSelect & "</select>"
End Function


%>