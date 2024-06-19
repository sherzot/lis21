<%
'******************************************************************************
'���@�́FclsP_ResumeStudent
'�T�@�v�Fform�Ŕ��ł���P_ResumeStudent�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/10/23
'�X�@�V�F
'******************************************************************************
Class clsP_ResumeStudent
	Public StaffCode
	Public Good
	Public Health
	Public Activity
	Public Specialty
	Public IsData
	Public MaxIndex
	Public Err
	Public ErrStyle

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_ResumeStudent �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/10/23
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1

		If Request.Form("CONF_StaffCode") <> "" Then StaffCode = Request.Form("CONF_StaffCode")
		If Request.Form("CONF_ResumeGood") <> "" Then IsData = True: Good = Request.Form("CONF_ResumeGood")
		If Request.Form("CONF_ResumeHealth") <> "" Then IsData = True: Health = Request.Form("CONF_ResumeHealth")
		If Request.Form("CONF_ResumeActivity") <> "" Then IsData = True: Activity = Request.Form("CONF_ResumeActivity")
		If Request.Form("CONF_ResumeSpecialty") <> "" Then IsData = True: Specialty = Request.Form("CONF_ResumeSpecialty")

		'�l�`�F�b�N
		Err = ""
		Set ErrStyle = Server.CreateObject("Scripting.Dictionary")
		ErrStyle.CompareMode = 1

		'���ӕ���E�Ȗ�
		If Good <> "" And ChkLen(Good, 500) = False Then
			Call DicAdd(ErrStyle, "CONF_ResumeGood", "style=""background-color:#ffff00;""")
			Err = Err & "���ӕ���E�Ȗڂ͔��p�P�����A�S�p�Q�����Ɛ����ĂT�O�O�����܂łł��B<br>"
		End If
		'���N���
		If Health <> "" And ChkLen(Health, 500) = False Then
			Call DicAdd(ErrStyle, "CONF_ResumeHealth", "style=""background-color:#ffff00;""")
			Err = Err & "���N��Ԃ͔��p�P�����A�S�p�Q�����Ɛ����ĂT�O�O�����܂łł��B<br>"
		End If
		'�N���u�����E��������
		If Activity <> "" And ChkLen(Activity, 500) = False Then
			Call DicAdd(ErrStyle, "CONF_ResumeActivity", "style=""background-color:#ffff00;""")
			Err = Err & "�N���u�����E���������͔��p�P�����A�S�p�Q�����Ɛ����ĂT�O�O�����܂łł��B<br>"
		End If
		'��E���Z
		If Specialty <> "" And ChkLen(Specialty, 500) = False Then
			Call DicAdd(ErrStyle, "CONF_ResumeSpecialty", "style=""background-color:#ffff00;""")
			Err = Err & "��E���Z�͔��p�P�����A�S�p�Q�����Ɛ����ĂT�O�O�����܂łł��B<br>"
		End If

		If ErrStyle.Count > 0 Then IsData = False
	End Sub

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_ResumeStudent ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/10/23
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If Good & Health & Activity & Specialty = "" Then
			GetRegSQL = "sp_Del_P_ResumeStudent '" & ChkSQLStr(vStaffCode) & "'"
			Exit Function
		End If

		If IsData = False Then Exit Function
		GetRegSQL = "up_Reg_P_ResumeStudent" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",'" & ChkSQLStr(Good) & "'" & _
			",'" & ChkSQLStr(Health) & "'" & _
			",'" & ChkSQLStr(Activity) & "'" & _
			",'" & ChkSQLStr(Specialty) & "'" & vbCrLf
	End Function
End Class
%>
