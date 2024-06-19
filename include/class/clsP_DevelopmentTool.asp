<%
'******************************************************************************
'���@�́FclsP_DevelopmentTool
'�T�@�v�Fform�Ŕ��ł���P_DevelopmentTool�e�[�u���p�̃f�[�^�������߂̃N���X
'�@�@�@�FCategoryCode���ɂ��̃N���X���쐬���Ďg�p����B
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_DevelopmentTool
	Public StaffCode
	Public CareerHistoryID()
	Public CategoryCode
	Public Code()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_DevelopmentTool �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize(vCategoryCode)
		Dim iCareerHistoryID
		Dim idx		: idx = 1
		Dim idx1
		Dim idx2
		Dim flag	: flag = False

		IsData = False
		MaxIndex = -1
		StaffCode = Request.Form("CONF_StaffCode")

		Err = ""

		'CONF_ �̖��O�� Lang, App, DB �Ɨ�����Ă��܂��Ă��鎖�ւ̏��u
		Select Case vCategoryCode
			Case "Lang": CategoryCode = "DevelopmentLanguage"
			Case "App": CategoryCode = "Application"
			Case "DB": CategoryCode = "Database"
			Case Else: CategoryCode = vCategoryCode
		End Select

		iCareerHistoryID = 0
		Do While True
			If ExistsForm("CONF_DevelopmentTool_" & vCategoryCode & idx) = False Then Exit Do

			If Request.Form("CONF_DevelopmentDetail" & idx) <> "" Then
				iCareerHistoryID = iCareerHistoryID + 1
				flag = True
			End If

			If flag = True And Request.Form("CONF_DevelopmentTool_" & vCategoryCode & idx) <> "" Then
				MaxIndex = MaxIndex + 1
				ReDim Preserve CareerHistoryID(MaxIndex)
				ReDim Preserve Code(MaxIndex)
				If flag = True Then
					CareerHistoryID(MaxIndex) = iCareerHistoryID
					Code(MaxIndex) = Split(Request.Form("CONF_DevelopmentTool_" & vCategoryCode & idx), ",")
				End If
			End If

			idx = idx + 1
			flag = False
		Loop

		For idx1 = 0 To MaxIndex
			For idx2 = LBound(Code(idx1)) To UBound(Code(idx1))
				Code(idx1)(idx2) = Trim(Code(idx1)(idx2))
				If Code(idx1)(idx2) <> "" And IsNumber(Code(idx1)(idx2), 3, False) = True Then
					IsData = True
				Else
					Err = Err & "Code(" & idx1 & ")(" & idx2 & ")" & vbCrLf
				End If
			Next
		Next
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_DevelopmentTool ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx
		Dim idxCode

		GetRegSQL = "EXEC sp_Del_P_DevelopmentTool '" & ChkSQLStr(vStaffCode) & "', '" & ChkSQLStr(CategoryCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			For idxCode = LBound(Code(idx)) To UBound(Code(idx))
				GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_DevelopmentTool" & _
					" '" & ChkSQLStr(vStaffCode) & "'" & _
					",'" & ChkSQLStr(CareerHistoryID(idx)) & "'" & _
					",''" & _
					",'" & ChkSQLStr(CategoryCode) & "'" & _
					",'" & ChkSQLStr(Code(idx)(idxCode)) & "'" & vbCrLf
			Next
		Next
	End Function
End Class
%>
