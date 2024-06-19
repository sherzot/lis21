<%
'******************************************************************************
'�T�@�v�F��ƃe�[�u���Ƀf�[�^��Insert, Update���鎞��
'�@�@�@�Fform�Ŕ��ł����f�[�^���i�[���邽�߂̃N���X�Q
'���@�l�F���O�� commonfunc.asp ���C���N���[�h���Ă������ƁI
'�X�@�V�F2006/05/13 LIS K.Kokubo �쐬
'******************************************************************************
%>
<%
'******************************************************************************
'���@�́FclsCompanyInfo
'�T�@�v�Fform�Ŕ��ł���CompanyInfo�e�[�u���p�̃f�[�^�������߂̃N���X
'���@�l�F
'�X�@�V�F2006/03/24 LIS K.Kokubo �쐬
'�@�@�@�F2008/06/05 LIS K.Kokubo GetData�֐��쐬
'�@�@�@�F2008/06/05 LIS K.Kokubo ChkData�֐��쐬
'�@�@�@�F2008/08/14 LIS M.Hayashi �����t���O�̒ǉ��ƃt���b�N�X�ړ�
'�@�@�@�F2009/01/05 LIS K.Kokubo �����������l�ǉ�
'�@�@�@�F2010/01/06 LIS K.Kokubo ��Ђ̕��͋C�ǉ�
'******************************************************************************
Class clsCompanyInfo
	Public CompanyCode
	Public CompanyKbn
	Public CompanyName_K
	Public CompanyName_F
	Public EstablishYear
	Public IndustryType
	Public CapitalAmount
	Public ForeinCapital
	Public ListClass
	Public AllEmployeeNum
	Public ManEmployeeNum
	Public WomanEmployeeNum
	Public HomepageAddress
	Public Post_U
	Public Post_L
	Public PrefectureCode
	Public City_K
	Public City_F
	Public Town
	Public Address
	Public TelephoneNumber
	Public StationCode1
	Public StationName1
	Public CompanySyudan1_1
	Public WorkOrBus1
	Public CompanySyudan1_2
	Public WorkBusTime1
	Public StationCode2
	Public StationName2
	Public CompanySyudan2_1
	Public WorkOrBus2
	Public CompanySyudan2_2
	Public WorkBusTime2
	Public SocietyInsurance
	Public Sanatorium
	Public EnterprisePension
	Public WealthShape
	Public StockOption
	Public RetirementPay
	Public ResidencePay
	Public FamilyPay
	Public EmployeeDormitory
	Public CompanyHouse
	Public NewEmployeeTraining
	Public OverseasTraining
	Public OtherTraning
	'Public FlexTime	'2008/08/14 Lis�� DEL
	Public WelfareProgramRemark '2009/01/05 LIS K.Kokubo ADD
	Public BusinessContents
	Public CompanyPR
	Public Simebi
	Public ContactPersonName
	Public Tanto1Yakusyoku
	Public Tanto2Name
	Public Tanto2Yakusyoku
	Public MailAddr
	Public NewJobMail
	Public DemandPrefectureCode
	Public DemandCity_K
	Public DemandCity_F
	Public DemandTown
	Public DemandAddress
	Public DemandSectionName
	Public DemandPersonName
	Public Atmosphere
	Public IsData
	Public MaxIndex
	Public Err
	Public ErrStyle

	'******************************************************************************
	'�T�@�v�FclsCompanyInfo�N���X�̏������֐�
	'���@���F
	'�߂�l�F�~
	'���@�l�F
	'�X�@�V�F2006/03/24 LIS K.Kokubo �쐬
	'******************************************************************************
	Private Sub Class_Initialize()
		MaxIndex = -1
		IsData = False

		Err = ""

		Set ErrStyle = Server.CreateObject("Scripting.Dictionary")
		ErrStyle.CompareMode = 1
	End Sub

	'******************************************************************************
	'�T�@�v�F�f�[�^�擾
	'���@���F
	'�߂�l�F�~
	'���@�l�FPOST�f�[�^��ǂݎ��A�e�v���p�e�B�Ƀf�[�^��ݒ肷��
	'�X�@�V�F2008/06/05 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Function GetData()
		If GetForm("CONF_CompanyCode", 1) <> "" Then CompanyCode = GetForm("CONF_CompanyCode", 1)
		If GetForm("CONF_CompanyKbn", 1) <> "" Then CompanyKbn = GetForm("CONF_CompanyKbn", 1)
		If GetForm("CONF_CompanyName_K", 1) <> "" Then CompanyName_K = GetForm("CONF_CompanyName_K", 1)
		If GetForm("CONF_CompanyName_F", 1) <> "" Then CompanyName_F = GetForm("CONF_CompanyName_F", 1)
		If GetForm("CONF_EstablishYear", 1) <> "" Then EstablishYear = GetForm("CONF_EstablishYear", 1)
		If GetForm("CONF_IndustryType", 1) <> "" Then IndustryType = GetForm("CONF_IndustryType", 1)
		If GetForm("CONF_CapitalAmount", 1) <> "" Then CapitalAmount = GetForm("CONF_CapitalAmount", 1)
		If GetForm("CONF_ForeinCapital", 1) <> "" Then ForeinCapital = GetForm("CONF_ForeinCapital", 1)
		If GetForm("CONF_ListClass", 1) <> "" Then ListClass = GetForm("CONF_ListClass", 1)
		If GetForm("CONF_AllEmployeeNum", 1) <> "" Then AllEmployeeNum = GetForm("CONF_AllEmployeeNum", 1)
		If GetForm("CONF_ManEmployeeNum", 1) <> "" Then ManEmployeeNum = GetForm("CONF_ManEmployeeNum", 1)
		If GetForm("CONF_WomanEmployeeNum", 1) <> "" Then WomanEmployeeNum = GetForm("CONF_WomanEmployeeNum", 1)
		If GetForm("CONF_HomepageAddress", 1) <> "" Then HomepageAddress = GetForm("CONF_HomepageAddress", 1)
		If GetForm("CONF_Post_U", 1) <> "" Then Post_U = GetForm("CONF_Post_U", 1)
		If GetForm("CONF_Post_L", 1) <> "" Then Post_L = GetForm("CONF_Post_L", 1)
		If GetForm("CONF_PrefectureCode", 1) <> "" Then PrefectureCode = GetForm("CONF_PrefectureCode", 1)
		If GetForm("CONF_City_K", 1) <> "" Then City_K = GetForm("CONF_City_K", 1)
		If GetForm("CONF_City_F", 1) <> "" Then City_F = GetForm("CONF_City_F", 1)
		If GetForm("CONF_Town", 1) <> "" Then Town = GetForm("CONF_Town", 1)
		If GetForm("CONF_Address", 1) <> "" Then Address = GetForm("CONF_Address", 1)
		If GetForm("CONF_TelephoneNumber", 1) <> "" Then TelephoneNumber = GetForm("CONF_TelephoneNumber", 1)
		If GetForm("CONF_StationCode1", 1) <> "" Then StationCode1 = GetForm("CONF_StationCode1", 1)
		If GetForm("CONF_StationName1", 1) <> "" Then StationName1 = GetForm("CONF_StationName1", 1)
		If GetForm("CONF_CompanySyudan1_1", 1) <> "" Then CompanySyudan1_1 = GetForm("CONF_CompanySyudan1_1", 1)
		If GetForm("CONF_WorkOrBus1", 1) <> "" Then WorkOrBus1 = GetForm("CONF_WorkOrBus1", 1)
		If GetForm("CONF_CompanySyudan1_2", 1) <> "" Then CompanySyudan1_2 = GetForm("CONF_CompanySyudan1_2", 1)
		If GetForm("CONF_WorkBusTime1", 1) <> "" Then WorkBusTime1 = GetForm("CONF_WorkBusTime1", 1)
		If GetForm("CONF_StationCode2", 1) <> "" Then StationCode2 = GetForm("CONF_StationCode2", 1)
		If GetForm("CONF_StationName2", 1) <> "" Then StationName2 = GetForm("CONF_StationName2", 1)
		If GetForm("CONF_CompanySyudan2_1", 1) <> "" Then CompanySyudan2_1 = GetForm("CONF_CompanySyudan2_1", 1)
		If GetForm("CONF_WorkOrBus2", 1) <> "" Then WorkOrBus2 = GetForm("CONF_WorkOrBus2", 1)
		If GetForm("CONF_CompanySyudan2_2", 1) <> "" Then CompanySyudan2_2 = GetForm("CONF_CompanySyudan2_2", 1)
		If GetForm("CONF_WorkBusTime2", 1) <> "" Then WorkBusTime2 = GetForm("CONF_WorkBusTime2", 1)
		If GetForm("CONF_SocietyInsurance", 1) <> "" Then SocietyInsurance = GetForm("CONF_SocietyInsurance", 1)
		If GetForm("CONF_Sanatorium", 1) <> "" Then Sanatorium = GetForm("CONF_Sanatorium", 1)
		If GetForm("CONF_EnterprisePension", 1) <> "" Then EnterprisePension = GetForm("CONF_EnterprisePension", 1)
		If GetForm("CONF_WealthShape", 1) <> "" Then WealthShape = GetForm("CONF_WealthShape", 1)
		If GetForm("CONF_StockOption", 1) <> "" Then StockOption = GetForm("CONF_StockOption", 1)
		If GetForm("CONF_RetirementPay", 1) <> "" Then RetirementPay = GetForm("CONF_RetirementPay", 1)
		If GetForm("CONF_ResidencePay", 1) <> "" Then ResidencePay = GetForm("CONF_ResidencePay", 1)
		If GetForm("CONF_FamilyPay", 1) <> "" Then FamilyPay = GetForm("CONF_FamilyPay", 1)
		If GetForm("CONF_EmployeeDormitory", 1) <> "" Then EmployeeDormitory = GetForm("CONF_EmployeeDormitory", 1)
		If GetForm("CONF_CompanyHouse", 1) <> "" Then CompanyHouse = GetForm("CONF_CompanyHouse", 1)
		If GetForm("CONF_NewEmployeeTraining", 1) <> "" Then NewEmployeeTraining = GetForm("CONF_NewEmployeeTraining", 1)
		If GetForm("CONF_OverseasTraining", 1) <> "" Then OverseasTraining = GetForm("CONF_OverseasTraining", 1)
		If GetForm("CONF_OtherTraning", 1) <> "" Then OtherTraning = GetForm("CONF_OtherTraning", 1)
		'If GetForm("CONF_FlexTime", 1) <> "" Then FlexTime = GetForm("CONF_FlexTime", 1)	'08/08/14 Lis�� DEL
		If GetForm("CONF_WelfareProgramRemark", 1) <> "" Then WelfareProgramRemark = GetForm("CONF_WelfareProgramRemark", 1)
		If GetForm("CONF_BusinessContents", 1) <> "" Then BusinessContents = GetForm("CONF_BusinessContents", 1)
		If GetForm("CONF_CompanyPR", 1) <> "" Then CompanyPR = GetForm("CONF_CompanyPR", 1)
		If GetForm("CONF_Simebi", 1) <> "" Then Simebi = GetForm("CONF_Simebi", 1)
		If GetForm("CONF_ContactPersonName", 1) <> "" Then ContactPersonName = GetForm("CONF_ContactPersonName", 1)
		If GetForm("CONF_Tanto1Yakusyoku", 1) <> "" Then Tanto1Yakusyoku = GetForm("CONF_Tanto1Yakusyoku", 1)
		If GetForm("CONF_Tanto2Name", 1) <> "" Then Tanto2Name = GetForm("CONF_Tanto2Name", 1)
		If GetForm("CONF_Tanto2Yakusyoku", 1) <> "" Then Tanto2Yakusyoku = GetForm("CONF_Tanto2Yakusyoku", 1)
		If GetForm("CONF_MailAddr", 1) <> "" Then MailAddr = GetForm("CONF_MailAddr", 1)
		If GetForm("CONF_NewJobMail", 1) <> "" Then NewJobMail = GetForm("CONF_NewJobMail", 1)
		If GetForm("CONF_DemandPrefectureCode", 1) <> "" Then DemandPrefectureCode = GetForm("CONF_DemandPrefectureCode", 1)
		If GetForm("CONF_DemandCity_K", 1) <> "" Then DemandCity_K = GetForm("CONF_DemandCity_K", 1)
		If GetForm("CONF_DemandCity_F", 1) <> "" Then DemandCity_F = GetForm("CONF_DemandCity_F", 1)
		If GetForm("CONF_DemandTown", 1) <> "" Then DemandTown = GetForm("CONF_DemandTown", 1)
		If GetForm("CONF_DemandAddress", 1) <> "" Then DemandAddress = GetForm("CONF_DemandAddress", 1)
		If GetForm("CONF_DemandSectionName", 1) <> "" Then DemandSectionName = GetForm("CONF_DemandSectionName", 1)
		If GetForm("CONF_DemandPersonName", 1) <> "" Then DemandPersonName = GetForm("CONF_DemandPersonName", 1)
		If GetForm("CONF_Atmosphere", 1) <> "" Then Atmosphere = GetForm("CONF_Atmosphere", 1)
	End Function

	'******************************************************************************
	'�T�@�v�F�f�[�^�̐������`�F�b�N
	'���@���F
	'�߂�l�F�~
	'���@�l�F�G���[���e��Err�v���p�e�B�ɏ�������
	'�X�@�V�F2008/06/05 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Function ChkData()
		IsData = False

		'��Ɩ�
		If CompanyName_K = "" Or ChkLen(CompanyName_K, 100) = False Then
			Call DicAdd(ErrStyle, "CompanyName_K", "background-color:#ffff00;")
			Err = Err & "��Ɩ��͔��p�P�����A�S�p�Q�����Ɛ����ĂP�O�O�����܂łł��B<br>"
		End If
		'��Ɩ��J�i
		If CompanyName_F = "" Or ChkLen(CompanyName_F, 80) = False Then
			Call DicAdd(ErrStyle, "CompanyName_F", "background-color:#ffff00;")
			Err = Err & "��Ɩ��J�i�͔��p�P�����A�S�p�Q�����Ɛ����ĂW�O�����܂łł��B<br>"
		End If
		'���ߓ�
		If Simebi <> "" And ChkInt(Simebi) = True Then
			If CInt(Simebi) < 1 Or CInt(Simebi) > 31 Then
				Call DicAdd(ErrStyle, "Simebi", "background-color:#ffff00;")
				Err = Err & "���ߓ��ɔ��p�����Ő�����������͂��ĉ������B<br>"
			End If
		ElseIf Simebi <> "" And ChkInt(Simebi) = False Then
			Call DicAdd(ErrStyle, "Simebi", "background-color:#ffff00;")
			Err = Err & "���ߓ��ɔ��p�����Ő�����������͂��ĉ������B<br>"
		End If
		'�ݗ��N�x
		If EstablishYear <> "" And IsDate(EstablishYear & "/01/01") = False Then
			Call DicAdd(ErrStyle, "EstablishYear", "background-color:#ffff00;")
			Err = Err & "�ݗ��N�x�ɔ��p�����Ő������N����͂��ĉ������B<br>"
		End If
		'���{��
		If CapitalAmount <> "" And ChkLen(CapitalAmount, 40) = False Then
			Call DicAdd(ErrStyle, "CapitalAmount", "background-color:#ffff00;")
			Err = Err & "���{���͔��p�P�����A�S�p�Q�����Ɛ����ĂS�O�����܂łł��B<br>"
		End If
		'�O��
		If ForeinCapital <> "" And ChkLen(ForeinCapital, 12) = False Then
			Call DicAdd(ErrStyle, "ForeinCapital", "background-color:#ffff00;")
			Err = Err & "�O���͔��p�P�����A�S�p�Q�����Ɛ����ĂP�Q�����܂łł��B<br>"
		End If
		'����
		If ListClass <> "" And ChkLen(ListClass, 40) = False Then
			Call DicAdd(ErrStyle, "ListClass", "background-color:#ffff00;")
			Err = Err & "�����͔��p�P�����A�S�p�Q�����Ɛ����ĂS�O�����܂łł��B<br>"
		End If
		'�S�Ј���
		If AllEmployeeNum <> "" And IsRE(AllEmployeeNum, "^\d*$", False) = False Then
			Call DicAdd(ErrStyle, "AllEmployeeNum", "background-color:#ffff00;")
			Err = Err & "�S�Ј����͔��p�����łP�Q���܂łł��B<br>"
		End If
		'�j����
		If ManEmployeeNum <> "" And IsRE(ManEmployeeNum, "^\d*$", False) = False Then
			Call DicAdd(ErrStyle, "ManEmployeeNum", "background-color:#ffff00;")
			Err = Err & "�j���Ј����͔��p�����łP�Q���܂łł��B<br>"
		End If
		'������
		If WomanEmployeeNum <> "" And IsRE(WomanEmployeeNum, "^\d*$", False) = False Then
			Call DicAdd(ErrStyle, "WomanEmployeeNum", "background-color:#ffff00;")
			Err = Err & "�����Ј����͔��p�����łP�Q���܂łł��B<br>"
		End If
		'�z�[���y�[�W
		If WomanEmployeeNum <> "" And IsRE(WomanEmployeeNum, "^\d*$", False) = False Then
			Call DicAdd(ErrStyle, "WomanEmployeeNum", "background-color:#ffff00;")
			Err = Err & "�z�[���y�[�W�͔��p�P�����A�S�p�Q�����Ɛ����ĂP�O�O�����܂łł��B<br>"
		End If
		'�X�֔ԍ�
		If IsRE(Post_U & Post_L, "^\d\d\d\d\d\d\d$", False) = False Then
			Call DicAdd(ErrStyle, "Post_U", "background-color:#ffff00;")
			Call DicAdd(ErrStyle, "Post_L", "background-color:#ffff00;")
			Err = Err & "�������X�֔ԍ��𔼊p�����œ��͂��ĉ������B<br>"
		End If
		'�s���{��
		If IsRE(PrefectureCode, "^\d\d\d$", False) = False Then
			Call DicAdd(ErrStyle, "PrefectureCode", "background-color:#ffff00;")
			Err = Err & "�s���{����I�����ĉ������B<br>"
		End If
		'�s��S
		If City_K <> "" And ChkLen(City_K, 80) = False Then
			Call DicAdd(ErrStyle, "City_K", "background-color:#ffff00;")
			Err = Err & "�s��S�͔��p�P�����A�S�p�Q�����Ɛ����ĂW�O�����܂łł��B<br>"
		End If
		'�s��S�J�i
		If City_F <> "" And ChkLen(City_F, 80) = False Then
			Call DicAdd(ErrStyle, "City_F", "background-color:#ffff00;")
			Err = Err & "�s��S�J�i�͔��p�P�����A�S�p�Q�����Ɛ����ĂW�O�����܂łł��B<br>"
		End If
		'����
		If Town <> "" And ChkLen(Town, 80) = False Then
			Call DicAdd(ErrStyle, "Town", "background-color:#ffff00;")
			Err = Err & "�����͔��p�P�����A�S�p�Q�����Ɛ����ĂW�O�����܂łł��B<br>"
		End If
		'�Ԓn��
		If Address <> "" And ChkLen(Address, 80) = False Then
			Call DicAdd(ErrStyle, "Address", "background-color:#ffff00;")
			Err = Err & "�Ԓn���͔��p�P�����A�S�p�Q�����Ɛ����ĂW�O�����܂łł��B<br>"
		End If
		'�ړ���i�P
		If CompanySyudan1_1 <> "" And ChkLen(CompanySyudan1_1, 20) = False Then
			Call DicAdd(ErrStyle, "CompanySyudan1_1", "background-color:#ffff00;")
			Err = Err & "�ړ���i�P�͔��p�P�����A�S�p�Q�����Ɛ����ĂQ�O�����܂łł��B<br>"
		End If
		'�ړ���i�P�̎���
		If WorkOrBus1 <> "" And ChkLen(WorkOrBus1, 3) = False Then
			Call DicAdd(ErrStyle, "WorkOrBus1", "background-color:#ffff00;")
			Err = Err & "�ړ���i�P�̎��Ԃ͔��p�����łR���܂łł��B<br>"
		End If
		'�ړ���i�Q
		If CompanySyudan1_2 <> "" And ChkLen(CompanySyudan1_2, 20) = False Then
			Call DicAdd(ErrStyle, "CompanySyudan1_2", "background-color:#ffff00;")
			Err = Err & "�ړ���i�Q�͔��p�P�����A�S�p�Q�����Ɛ����ĂQ�O�����܂łł��B<br>"
		End If
		'�ړ���i�Q�̎���
		If WorkBusTime1 <> "" And ChkLen(WorkBusTime1, 3) = False Then
			Call DicAdd(ErrStyle, "WorkBusTime1", "background-color:#ffff00;")
			Err = Err & "�ړ���i�Q�̎��Ԃ͔��p�����łR���܂łł��B<br>"
		End If
		'�ړ���i�P
		If CompanySyudan2_1 <> "" And ChkLen(CompanySyudan2_1, 20) = False Then
			Call DicAdd(ErrStyle, "CompanySyudan2_1", "background-color:#ffff00;")
			Err = Err & "�ړ���i�P�͔��p�P�����A�S�p�Q�����Ɛ����ĂQ�O�����܂łł��B<br>"
		End If
		'�ړ���i�P�̎���
		If WorkOrBus2 <> "" And ChkLen(WorkOrBus2, 3) = False Then
			Call DicAdd(ErrStyle, "WorkOrBus2", "background-color:#ffff00;")
			Err = Err & "�ړ���i�P�̎��Ԃ͔��p�����łR���܂łł��B<br>"
		End If
		'�ړ���i�Q
		If CompanySyudan2_2 <> "" And ChkLen(CompanySyudan2_2, 20) = False Then
			Call DicAdd(ErrStyle, "CompanySyudan2_2", "background-color:#ffff00;")
			Err = Err & "�ړ���i�Q�͔��p�P�����A�S�p�Q�����Ɛ����ĂQ�O�����܂łł��B<br>"
		End If
		'�ړ���i�Q�̎���
		If WorkBusTime2 <> "" And ChkLen(WorkBusTime2, 3) = False Then
			Call DicAdd(ErrStyle, "WorkBusTime2", "background-color:#ffff00;")
			Err = Err & "�ړ���i�Q�̎��Ԃ͔��p�����łR���܂łł��B<br>"
		End If
		'�����������l
		If WelfareProgramRemark <> "" And ChkLen(WelfareProgramRemark, 100) = False Then
			Call DicAdd(ErrStyle, "WelfareProgramRemark", "background-color:#ffff00;")
			Err = Err & "�����������l�͔��p�P�����A�S�p�Q�����Ɛ����ĂP�O�O�����܂łł��B<br>"
		End If
		'���Ɠ��e
		If BusinessContents <> "" And ChkLen(BusinessContents, 1000) = False Then
			Call DicAdd(ErrStyle, "BusinessContents", "background-color:#ffff00;")
			Err = Err & "���Ɠ��e�͔��p�P�����A�S�p�Q�����Ɛ����ĂP�O�O�O�����܂łł��B<br>"
		End If
		'��Јē�
		If CompanyPR <> "" And ChkLen(CompanyPR, 1000) = False Then
			Call DicAdd(ErrStyle, "CompanyPR", "background-color:#ffff00;")
			Err = Err & "��Јē��͔��p�P�����A�S�p�Q�����Ɛ����ĂP�O�O�O�����܂łł��B<br>"
		End If
		'�����S���Җ��P
		If ContactPersonName <> "" And ChkLen(ContactPersonName, 40) = False Then
			Call DicAdd(ErrStyle, "ContactPersonName", "background-color:#ffff00;")
			Err = Err & "�����S���Җ��͔��p�P�����A�S�p�Q�����Ɛ����ĂS�O�����܂łł��B<br>"
		End If
		'�����S���Җ�E�P
		If Tanto1Yakusyoku <> "" And ChkLen(Tanto1Yakusyoku, 40) = False Then
			Call DicAdd(ErrStyle, "Tanto1Yakusyoku", "background-color:#ffff00;")
			Err = Err & "�����S���Җ�E�͔��p�P�����A�S�p�Q�����Ɛ����ĂS�O�����܂łł��B<br>"
		End If
		'�����S���Җ��Q
		If Tanto2Name <> "" And ChkLen(Tanto2Name, 40) = False Then
			Call DicAdd(ErrStyle, "Tanto2Name", "background-color:#ffff00;")
			Err = Err & "�����S���Җ��͔��p�P�����A�S�p�Q�����Ɛ����ĂS�O�����܂łł��B<br>"
		End If
		'�����S���Җ�E�Q
		If Tanto2Yakusyoku <> "" And ChkLen(Tanto2Yakusyoku, 40) = False Then
			Call DicAdd(ErrStyle, "Tanto2Yakusyoku", "background-color:#ffff00;")
			Err = Err & "�����S���Җ�E�͔��p�P�����A�S�p�Q�����Ɛ����ĂS�O�����܂łł��B<br>"
		End If
		'�d�b�ԍ�
		If IsTel(TelephoneNumber, "0") = False Then
			Call DicAdd(ErrStyle, "TelephoneNumber", "background-color:#ffff00;")
			Err = Err & "�������d�b�ԍ��𔼊p�����ƃn�C�t�� - �œ��͂��ĉ������B<br>"
		End If
		'���[���A�h���X
		If MailAddr <> "" And IsMailAddress(MailAddr) = False Then
			Call DicAdd(ErrStyle, "MailAddr", "background-color:#ffff00;")
			Err = Err & "���������[���A�h���X�𔼊p�A�T�O�����ȓ��œ��͂��ĉ������B<br>"
		End If
		'���������t��s��S
		If DemandCity_K <> "" And ChkLen(DemandCity_K, 80) = False Then
			Call DicAdd(ErrStyle, "DemandCity_K", "background-color:#ffff00;")
			Err = Err & "���������t��s��S�͔��p�P�����A�S�p�Q�����Ɛ����ĂW�O�����܂łł��B<br>"
		End If
		'���������t��s��S�J�i
		If DemandCity_F <> "" And ChkLen(DemandCity_F, 80) = False Then
			Call DicAdd(ErrStyle, "DemandCity_F", "background-color:#ffff00;")
			Err = Err & "���������t��s��S�J�i�͔��p�P�����A�S�p�Q�����Ɛ����ĂW�O�����܂łł��B<br>"
		End If
		'���������t�撬��
		If DemandTown <> "" And ChkLen(DemandTown, 80) = False Then
			Call DicAdd(ErrStyle, "DemandTown", "background-color:#ffff00;")
			Err = Err & "���������t�撬���͔��p�P�����A�S�p�Q�����Ɛ����ĂW�O�����܂łł��B<br>"
		End If
		'���������t��Ԓn��
		If DemandAddress <> "" And ChkLen(DemandAddress, 80) = False Then
			Call DicAdd(ErrStyle, "DemandAddress", "background-color:#ffff00;")
			Err = Err & "���������t��s��S�J�i�͔��p�P�����A�S�p�Q�����Ɛ����ĂW�O�����܂łł��B<br>"
		End If
		'���������t�敔��
		If DemandSectionName <> "" And ChkLen(DemandSectionName, 40) = False Then
			Call DicAdd(ErrStyle, "DemandSectionName", "background-color:#ffff00;")
			Err = Err & "���������t�敔���͔��p�P�����A�S�p�Q�����Ɛ����ĂS�O�����܂łł��B<br>"
		End If
		'���������t��S����
		If DemandPersonName <> "" And ChkLen(DemandPersonName, 40) = False Then
			Call DicAdd(ErrStyle, "DemandPersonName", "background-color:#ffff00;")
			Err = Err & "���������t��S���Җ��͔��p�P�����A�S�p�Q�����Ɛ����ĂS�O�����܂łł��B<br>"
		End If
		'��Ђ̕��͋C
		If Atmosphere <> "" And ChkLen(Atmosphere, 500) = False Then
			Call DicAdd(ErrStyle, "DemandPersonName", "background-color:#ffff00;")
			Err = Err & "��Ђ̕��͋C�͔��p�P�����A�S�p�Q�����Ɛ����ĂT�O�O�����܂łł��B<br>"
		End If

		If Err = "" Then IsData = True
	End Function

	'******************************************************************************
	'���@�́FGetRegSQL
	'�T�@�v�Fsp_Reg_CompanyInfo ���sSQL�擾
	'���@�l�F
	'�쐬�ҁFLis Kokubo
	'�쐬���F2006/03/24
	'�X�@�V�F
	'******************************************************************************
	Public Function GetRegSQL(vCompanyCode)
		GetRegSQL = ""
		GetRegSQL = GetRegSQL & "EXEC up_RegCompanyInfo_Navi"
		GetRegSQL = GetRegSQL & " '" & vCompanyCode & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanyKbn & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanyName_K & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanyName_F & "'"
		GetRegSQL = GetRegSQL & ",'" & EstablishYear & "'"
		GetRegSQL = GetRegSQL & ",'" & IndustryType & "'"
		GetRegSQL = GetRegSQL & ",'" & CapitalAmount & "'"
		GetRegSQL = GetRegSQL & ",'" & ForeinCapital & "'"
		GetRegSQL = GetRegSQL & ",'" & ListClass & "'"
		GetRegSQL = GetRegSQL & ",'" & AllEmployeeNum & "'"
		GetRegSQL = GetRegSQL & ",'" & ManEmployeeNum & "'"
		GetRegSQL = GetRegSQL & ",'" & WomanEmployeeNum & "'"
		GetRegSQL = GetRegSQL & ",'" & HomepageAddress & "'"
		GetRegSQL = GetRegSQL & ",'" & Post_U & "'"
		GetRegSQL = GetRegSQL & ",'" & Post_L & "'"
		GetRegSQL = GetRegSQL & ",'" & PrefectureCode & "'"
		GetRegSQL = GetRegSQL & ",'" & City_K & "'"
		GetRegSQL = GetRegSQL & ",'" & City_F & "'"
		GetRegSQL = GetRegSQL & ",'" & Town & "'"
		GetRegSQL = GetRegSQL & ",'" & Address & "'"
		GetRegSQL = GetRegSQL & ",'" & TelephoneNumber & "'"
		GetRegSQL = GetRegSQL & ",'" & StationCode1 & "'"
		GetRegSQL = GetRegSQL & ",'" & StationName1 & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanySyudan1_1 & "'"
		GetRegSQL = GetRegSQL & ",'" & WorkOrBus1 & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanySyudan1_2 & "'"
		GetRegSQL = GetRegSQL & ",'" & WorkBusTime1 & "'"
		GetRegSQL = GetRegSQL & ",'" & StationCode2 & "'"
		GetRegSQL = GetRegSQL & ",'" & StationName2 & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanySyudan2_1 & "'"
		GetRegSQL = GetRegSQL & ",'" & WorkOrBus2 & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanySyudan2_2 & "'"
		GetRegSQL = GetRegSQL & ",'" & WorkBusTime2 & "'"
		GetRegSQL = GetRegSQL & ",'" & SocietyInsurance & "'"
		GetRegSQL = GetRegSQL & ",'" & Sanatorium & "'"
		GetRegSQL = GetRegSQL & ",'" & EnterprisePension & "'"
		GetRegSQL = GetRegSQL & ",'" & WealthShape & "'"
		GetRegSQL = GetRegSQL & ",'" & StockOption & "'"
		GetRegSQL = GetRegSQL & ",'" & RetirementPay & "'"
		GetRegSQL = GetRegSQL & ",'" & ResidencePay & "'"
		GetRegSQL = GetRegSQL & ",'" & FamilyPay & "'"
		GetRegSQL = GetRegSQL & ",'" & EmployeeDormitory & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanyHouse & "'"
		GetRegSQL = GetRegSQL & ",'" & NewEmployeeTraining & "'"
		GetRegSQL = GetRegSQL & ",'" & OverseasTraining & "'"
		GetRegSQL = GetRegSQL & ",'" & OtherTraning & "'"
		'GetRegSQL = GetRegSQL & ",'" & FlexTime & "'" '08/08/14 Lis�� DEL
		GetRegSQL = GetRegSQL & ",'" & WelfareProgramRemark & "'"
		GetRegSQL = GetRegSQL & ",'" & BusinessContents & "'"
		GetRegSQL = GetRegSQL & ",'" & CompanyPR & "'"
		GetRegSQL = GetRegSQL & ",'" & Simebi & "'"
		GetRegSQL = GetRegSQL & ",'" & ContactPersonName & "'"
		GetRegSQL = GetRegSQL & ",'" & Tanto1Yakusyoku & "'"
		GetRegSQL = GetRegSQL & ",'" & Tanto2Name & "'"
		GetRegSQL = GetRegSQL & ",'" & Tanto2Yakusyoku & "'"
		GetRegSQL = GetRegSQL & ",'" & MailAddr & "'"
		GetRegSQL = GetRegSQL & ",'" & NewJobMail & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandPrefectureCode & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandCity_K & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandCity_F & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandTown & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandAddress & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandSectionName & "'"
		GetRegSQL = GetRegSQL & ",'" & DemandPersonName & "'"
		GetRegSQL = GetRegSQL & ",'" & Atmosphere & "'"
	End Function
End Class
%>
