<%
'******************************************************************************
'���@�́FclsP_IntroductionComment
'�T�@�v�Fform�Ŕ��ł���P_IntroductionComment�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_IntroductionComment
	Public StaffCode
	Public Comment(16)
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_IntroductionComment �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		Dim sidx
		Dim idx

		IsData = False
		MaxIndex = UBound(Comment)
		StaffCode = Request.Form("CONF_StaffCode")

		For idx = 1 To UBound(Comment)
			If idx <= 9 Then
				sidx = "00" & idx
			Else
				sidx = "0" & idx
			End If

			Comment(idx) = Request.Form("CONF_IntroductionComment" & sidx)
			If Comment(idx) <> "" Then IsData = True
		Next
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_IntroductionComment ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx
		Dim sidx

		If IsData = False Then Exit Function

		GetRegSQL = ""
		For idx = 1 To MaxIndex
			If Comment(idx) <> "" Then
				If idx <= 9 Then
					sidx = "00" & idx
				Else
					sidx = "0" & idx
				End If

				GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_IntroductionComment" & _
					" '" & ChkSQLStr(vStaffCode) & "'" & _
					",'IntroductionComment'" & _
					",'" & ChkSQLStr(sidx) & "'" & _
					",'" & ChkSQLStr(Comment(idx)) & "'" & vbCrLf
			End If
		Next
	End Function
End Class
%>
