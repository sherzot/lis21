<%
'******************************************************************************
'���@�́FclsP_ResumeOther
'�T�@�v�Fform�Ŕ��ł���P_ResumeOther�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_ResumeOther
	Public StaffCode
	Public PrintFlag()
	Public Subject()
	Public WishMotive()
	Public CommuteTime()
	Public HopeColumn()
	Public IsData
	Public MaxIndex
	Public Err
	Public ErrStyle

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_ResumeOther �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If Request.Form("StaffCode") <> "" Then StaffCode = Request.Form("StaffCode")

		Do While True
			If ExistsForm("CONF_Subject" & idx) = False Then Exit Do

			If Request.Form("CONF_Subject" & idx) & Request.Form("CONF_WishMotive" & idx) & Request.Form("CONF_CommuteTime" & idx) & Request.Form("CONF_HopeColumn" & idx) <> "" Then
				MaxIndex = MaxIndex + 1

				ReDim Preserve Subject(MaxIndex) : Subject(MaxIndex) = Request.Form("CONF_Subject" & idx)
				ReDim Preserve WishMotive(MaxIndex) : WishMotive(MaxIndex) = Request.Form("CONF_WishMotive" & idx)
				ReDim Preserve CommuteTime(MaxIndex) : CommuteTime(MaxIndex) = Request.Form("CONF_CommuteTime" & idx)
				ReDim Preserve HopeColumn(MaxIndex) : HopeColumn(MaxIndex) = Request.Form("CONF_HopeColumn" & idx)

				If CommuteTime(MaxIndex) <> "" And IsNumber(CommuteTime(MaxIndex), 0, False) = False Then CommuteTime(MaxIndex) = "": Err = Err & "BranchCode" & vbCrLf
			End If
			idx = idx + 1
		Loop

		'�l�`�F�b�N
		Set ErrStyle = Server.CreateObject("scripting.dictionary")
		ErrStyle.CompareMode = 1

		For idx = 1 To MaxIndex
			'�^�C�g��
			If Subject(idx) <> "" And ChkLen(Subject(idx), 200) = False Then
				Call DicAdd(ErrStyle, "CONF_Subject" & idx, "style=""background-color:#ffff00;""")
				Err = Err & "�^�C�g���͔��p�P�����A�S�p�Q�����Ɛ����ĂQ�O�O�����܂łł��B<br>"
			End If

			'�u�]���@
			If WishMotive(idx) <> "" And ChkLen(WishMotive(idx), 2000) = False Then
				Call DicAdd(ErrStyle, "CONF_WishMotive" & idx, "style=""background-color:#ffff00;""")
				Err = Err & "�u�]���@�͔��p�P�����A�S�p�Q�����Ɛ����ĂQ�O�O�O�����܂łł��B<br>"
			End If

			'��]�ʋΎ���
			If CommuteTime(idx) <> "" And IsNumber(CommuteTime(idx), 0, True) = False Then
				Call DicAdd(ErrStyle, "CONF_CommuteTime" & idx, "style=""background-color:#ffff00;""")
				Err = Err & "��]�ʋΎ��Ԃ͔��p�����œ��͂��ĉ������B<br>"
			End If

			'�{�l��]
			If HopeColumn(idx) <> "" And ChkLen(HopeColumn(idx), 2000) = False Then
				Call DicAdd(ErrStyle, "CONF_HopeColumn" & idx, "style=""background-color:#ffff00;""")
				Err = Err & "�{�l��]�͔��p�P�����A�S�p�Q�����Ɛ����ĂQ�O�O�O�����܂łł��B<br>"
			End If
		Next

		'If Err = "" Then IsData = True
		IsData = True
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�FclsP_ResumeOther ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_ResumeOther '" & ChkSQLStr(vStaffCode) & "'"
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_ResumeOther" & _
				" '" & ChkSQLStr(vStaffCode) & "'" & _
				",''" & _
				",'" & ChkSQLStr(Subject(idx)) & "'" & _
				",'" & ChkSQLStr(WishMotive(idx)) & "'" & _
				",'" & ChkSQLStr(CommuteTime(idx)) & "'" & _
				",'" & ChkSQLStr(HopeColumn(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
