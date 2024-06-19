<%
'******************************************************************************
'���@�́FclsP_Info
'�T�@�v�Fform�Ŕ��ł���P_Info�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�쐬�ҁFLis Kokubo
'�쐬���F2006/04/05
'�X�@�V�F
'******************************************************************************
Class clsP_Info
	Public StaffCode
	Public Name
	Public Name_F
	Public SearchName
	Public SearchName_F
	Public OldName
	Public Birthday
	Public Sex
	Public MarriageFlag
	Public Post_U
	Public Post_L
	Public PrefectureCode
	Public City
	Public City_F
	Public Town
	Public Town_F
	Public Address
	Public Address_F
	Public LivingType
	Public HomeTelephoneNumber
	Public CountryTelephoneNumber
	Public PortableTelephoneNumber
	Public FaxNumber
	Public MailAddress
	Public PortableMailAddress
	Public UrgencyPost_U
	Public UrgencyPost_L
	Public UrgencyAddress
	Public UrgencyAddress_F
	Public UrgencyTelephoneNumber
	Public URL
	Public InfoSourceType
	Public InfoSourceDay
	Public InfoSourceOther
	Public DependentFlag
	Public DependentNumber
	Public SpouseFlag
	Public CurrentCompanyName
	Public CurrentCompanyName_F
	Public SocietyInsuranceIn
	Public SocietyInsuranceLoss
	Public EmployInsuranceIn
	Public EmployInsuranceLoss
	Public IsData
	Public MaxIndex
	Public Err
	Public ErrStyle

	'******************************************************************************
	'���@�́FInitialize
	'�T�@�v�FclsP_Info �N���X�̏������֐�
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Sub Initialize()
		IsData = False
		MaxIndex = -1
	End Sub

	Public Function ChkData()
		'�l�`�F�b�N
		Err = ""
		Set ErrStyle = Server.CreateObject("Scripting.Dictionary")
		ErrStyle.CompareMode = 1

		'���O
		If Name = "" Or ChkLen(Name, 100) = False Then
			ErrStyle("Name_1") = "background-color:#ffff00;"
			ErrStyle("Name_2") = "background-color:#ffff00;"
			Err = Err & "���O���s���ł��B<br>"
		End If

		'���O�J�i
		If Name_F = "" Or ChkLen(Name_F, 100) = False Then
			ErrStyle("Name_F_1") = "background-color:#ffff00;"
			ErrStyle("Name_F_2") = "background-color:#ffff00;"
			Err = Err & "���O�J�i���s���ł��B<br>"
		End If

		'�a����
		If Birthday <> "" And IsDay(Birthday) = False Then
			ErrStyle("Birthday") = "background-color:#ffff00;"
			Err = Err & "�a�����̓��t���s���ł��B<br>"
		End If

		'����
		If Sex <> "" And IsRE(Sex, "^[12]$", True) = False Then
			ErrStyle("Sex") = "background-color:#ffff00;"
			Err = Err & "���ʂ���͂��Ă��������B<br>"
		End If

		'
		If MarriageFlag <> "" And IsFlag(MarriageFlag) = False Then
			ErrStyle("MarriageFlag") = "background-color:#ffff00;"
			Err = Err & "�}�{���s���ł��B<br>"
		End If

		'�X�֔ԍ�
		If Post_U & Post_L <> "" And IsNumber(Post_U & Post_L, 7, False) = False Then
			ErrStyle("Post_U") = "background-color:#ffff00;"
			ErrStyle("Post_L") = "background-color:#ffff00;"
			Err = Err & "�X�֔ԍ����s���ł��B<br>"
		End If

		'�s���{���R�[�h
		If PrefectureCode <> "" And IsNumber(PrefectureCode, 3, False) = False Then
			ErrStyle("PrefectureCode") = "background-color:#ffff00;"
			Err = Err & "�Z���̓s���{�����s���ł��B<br>"
		End If

		'�����l���
		If LivingType <> "" And IsRE(LivingType, "^[1234]$", True) = False Then
			ErrStyle("PrefectureCode") = "background-color:#ffff00;"
			Err = Err & "�����l��ނ��s���ł��B<br>"
		End If

		'�d�b�ԍ�
		If HomeTelephoneNumber <> "" And IsNumber(HomeTelephoneNumber, 0, False) = False Then
			ErrStyle("HomeTelephoneNumber") = "background-color:#ffff00;"
			Err = Err & "�d�b�ԍ����s���ł��B<br>"
		End If

		'���Ɠd�b�ԍ�
		If CountryTelephoneNumber <> "" And IsNumber(CountryTelephoneNumber, 0, False) = False Then
			ErrStyle("CountryTelephoneNumber") = "background-color:#ffff00;"
			Err = Err & "���Ɠd�b�ԍ����s���ł��B<br>"
		End If

		'�g�єԍ�
		If PortableTelephoneNumber <> "" And IsNumber(PortableTelephoneNumber, 0, False) = False Then
			ErrStyle("PortableTelephoneNumber") = "background-color:#ffff00;"
			Err = Err & "�g�єԍ����s���ł��B<br>"
		End If

		'FAX
		If FaxNumber <> "" And IsNumber(FaxNumber, 0, False) = False Then
			ErrStyle("PortableTelephoneNumber") = "background-color:#ffff00;"
			Err = Err & "FAX�ԍ����s���ł��B<br>"
		End If

		'�ً}�A����X�֔ԍ�
		If UrgencyPost_U & UrgencyPost_L <> "" And IsNumber(UrgencyPost_U & UrgencyPost_L, 7, False) = False Then
			ErrStyle("UrgencyPost_U") = "background-color:#ffff00;"
			ErrStyle("UrgencyPost_L") = "background-color:#ffff00;"
			Err = Err & "�ً}�A����X�֔ԍ����s���ł��B<br>"
		End If

		'�ً}�A����Z��
		If UrgencyAddress <> "" And ChkLen(UrgencyAddress, 200) = False Then
			ErrStyle("UrgencyAddress") = "background-color:#ffff00;"
			Err = Err & "�ً}�A����Z���̕��������������܂��B�S�p�łP�O�O�����܂œ��͂ł��܂��B<br>"
		End If

		'�ً}�A����Z���J�i
		If UrgencyAddress_F <> "" And ChkLen(UrgencyAddress_F, 200) = False Then
			ErrStyle("UrgencyAddress_F") = "background-color:#ffff00;"
			Err = Err & "�ً}�A����Z���J�i�̕��������������܂��B�S�p�łP�O�O�����܂œ��͂ł��܂��B<br>"
		End If

		'�ً}�A����d�b�ԍ�
		If UrgencyTelephoneNumber <> "" And IsTel(UrgencyTelephoneNumber, 0) = False Then
			ErrStyle("UrgencyTelephoneNumber") = "background-color:#ffff00;"
			Err = Err & "�ً}�A����d�b�ԍ����s���ł��B<br>"
		End If

		'��񌹎��
		If InfoSourceType <> "" And IsNumber(InfoSourceType, 3, False) = False Then
			ErrStyle("InfoSourceType") = "background-color:#ffff00;"
			Err = Err & "��񌳎�ނ��s���ł��B<br>"
		End If

		'��񌹓��t
		If InfoSourceDay <> "" And IsDay(InfoSourceDay) = False Then
			ErrStyle("InfoSourceDay") = "background-color:#ffff00;"
			Err = Err & "��񌳓��t���s���ł��B<br>"
		End If

		'�}�{�t���O
		If DependentFlag <> "" And IsFlag(DependentFlag) = False Then
			ErrStyle("DependentFlag") = "background-color:#ffff00;"
			Err = Err & "�}�{���s���ł��B<br>"
		End If

		'�}�{�l��
		If DependentNumber <> "" And IsNumber(DependentNumber, 0, False) = False Then
			ErrStyle("DependentNumber") = "background-color:#ffff00;"
			Err = Err & "�}�{�l�����s���ł��B<br>"
		End If

		'�z��ҕ}�{�t���O
		If SpouseFlag <> "" And IsFlag(SpouseFlag) = False Then
			ErrStyle("SpouseFlag") = "background-color:#ffff00;"
			Err = Err & "�z��ҕ}�{���s���ł��B<br>"
		End If
	End Function

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_P_Info ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		If IsData = False Then Exit Function

		GetRegSQL = "up_Reg_P_Info '" & vStaffCode & "'" & _
			",'S'" & _
			",'" & Name & "'" & _
			",'" & Name_F & "'" & _
			",'" & SearchName & "'" & _
			",'" & SearchName_F & "'" & _
			",'" & OldName & "'" & _
			",'" & Birthday & "'" & _
			",'" & Sex & "'" & _
			",'" & MarriageFlag & "'" & _
			",'" & Post_U & "'" & _
			",'" & Post_L & "'" & _
			",'" & PrefectureCode & "'" & _
			",'" & City & "'" & _
			",'" & City_F & "'" & _
			",'" & Town & "'" & _
			",'" & Town_F & "'" & _
			",'" & Address & "'" & _
			",'" & Address_F & "'" & _
			",'" & LivingType & "'" & _
			",'" & HomeTelephoneNumber & "'" & _
			",'" & CountryTelephoneNumber & "'" & _
			",'" & PortableTelephoneNumber & "'" & _
			",'" & FaxNumber & "'" & _
			",'" & MailAddress & "'" & _
			",'" & PortableMailAddress & "'" & _
			",'" & UrgencyPost_U & "'" & _
			",'" & UrgencyPost_L & "'" & _
			",'" & UrgencyAddress & "'" & _
			",'" & UrgencyAddress_F & "'" & _
			",'" & UrgencyTelephoneNumber & "'" & _
			",'" & URL & "'" & _
			",'" & InfoSourceType & "'" & _
			",'" & InfoSourceDay & "'" & _
			",'" & InfoSourceOther & "'" & _
			",'" & DependentFlag & "'" & _
			",'" & DependentNumber & "'" & _
			",'" & SpouseFlag & "'" & _
			",'" & CurrentCompanyName & "'" & _
			",'" & CurrentCompanyName_F & "'" & _
			",'" & SocietyInsuranceIn & "'" & _
			",'" & SocietyInsuranceLoss & "'" & _
			",'" & EmployInsuranceIn & "'" & _
			",'" & EmployInsuranceLoss & "'" & vbCrLf
	End Function
End Class
%>
