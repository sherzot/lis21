<%
'******************************************************************************
'�T�@�v�F����������ێ�����N���X
'�ց@���F��Public
'�@�@�@�FGetSearchParam				�F���d���ڍ׌����y�[�W�֓n��GET�p�����[�^�𐶐����Ď擾
'�@�@�@�FGetSQLOrderSearchDetail	�F���l�[�ڍ׌����r�p�k���擾
'�@�@�@�FGetHtmlSearchCondition		�F���l�[�ڍ׌��������o�͂g�s�l�k���擾
'�@�@�@�F
'�@�@�@�F��Private
'�@�@�@�FClass_Initialize			�F�R���X�g���N�^
'�@�@�@�FSetNames					�F�R�[�h�ɑΉ��������̂������o�ϐ��ɐݒ�
'�@�@�@�FChkData					�F�����o�ϐ��̐��������`�F�b�N���Ē���
'�@�@�@�F
'���@�l�F������ �ڍ׌����p�p�����[�^ �i�A�h�z�b�N�Ȃr�p�k�����j
'�@�@�@�Fsotf	�F�Г��O�Č������t���O
'�@�@�@�Fsnewf	�F�V���t���O�i�P�T�Ԉȓ��Ɍf�ڂ̂��������́j
'�@�@�@�Fsjtbig1�F��]�E��啪�ނP
'�@�@�@�Fsjt1	�F��]�E��P
'�@�@�@�Fsjtbig2�F��]�E��啪�ނQ
'�@�@�@�Fsjt2	�F��]�E��Q
'�@�@�@�Fsrc	�F��]����
'�@�@�@�Fssc	�F��]�w
'�@�@�@�Fspc	�F��]�s���{��
'�@�@�@�Fsct	�F��]�s��S
'�@�@�@�Fswt1	�F��]�Ζ��`�ԂP
'�@�@�@�Fswt2	�F��]�Ζ��`�ԂQ
'�@�@�@�Fswt3	�F��]�Ζ��`�ԂR
'�@�@�@�Fsit	�F��]�Ǝ�(�J���}��؂� [XX,XX,XX])
'�@�@�@�Fssp01	�F�����i���o�����}�j
'�@�@�@�Fssp02	�F�����i��w���������j
'�@�@�@�Fssp03	�F�����i�h���j�����ݖ��g�p
'�@�@�@�Fssp04	�F�����i�t�h�^�[���j
'�@�@�@�Fssp05	�F�����i�x���P�Q�O���ȏ�j
'�@�@�@�Fssp06	�F�����i�t���b�N�X�^�C���j
'�@�@�@�Fssp07	�F�����i�w�߁j
'�@�@�@�Fssp08	�F�����i�։��E�����j
'�@�@�@�Fssp09	�F�����i�V�z�r���E�I�t�B�X�j
'�@�@�@�Fssp10	�F�����i���w�r���i�����h�}�[�N�j�j
'�@�@�@�Fssp11	�F�����i���m�x�[�V�����r���E�I�t�B�X�j
'�@�@�@�Fssp12	�F�����i�f�U�C�i�[�Y�r���E�I�t�B�X�j
'�@�@�@�Fssp13	�F�����i�Ј��H���j
'�@�@�@�Fssp14	�F�����i�c��10h�ȓ��j
'�@�@�@�Fssp15	�F�����i�Y�x�E��x���т���j
'�@�@�@�Fssp16	�F�����i�������R�j
'�@�@�@�Fssp17	�F�����i�q��ă}�}���}�j
'�@�@�@�Fssp18	�F�����i18���܂łɑގЁj
'�@�@�@�Fssp19	�F�����i1��6���Ԉȓ��J���j
'�@�@�@�Fssp20	�F�����i��Q�Ҋ��}�j
'�@�@�@�Fssp21	�F�����i�Z���p�S�z�⏕����j
'�@�@�@�Fssp22	�F�����i�Z���p�ꕔ�⏕����j
'�@�@�@�Fssp23	�F�����i�H���E�d���t���Č��j
'�@�@�@�Fssp24	�F�����i�H���⏕���x����j
'�@�@�@�Fssp25	�F�����i���C������x����j
'�@�@�@�Fssp26	�F�����i�N�Ƌ@�ޕ⏕���x����j
'�@�@�@�Fssp27	�F�����i�����q�E�ᗘ�q�⏕���x����j
'�@�@�@�Fssp28	�F�����i�y�n�E�X�ܓ��񋟐��x����j
'�@�@�@�Fssp29	�F�����i�A�E���j�������x����j
'�@�@�@�Fssp30	�F�����i���Ј��o�p���x����j
'�@�@�@�Fssp31	�F�����i�Еۊ����j
'�@�@�@�Fsnewf	�F�V���t���O
'�@�@�@�Fsppf	�F������
'�@�@�@�Fsyimin	�F�N������
'�@�@�@�Fsyimax	�F�N�����
'�@�@�@�Fsmimin	�F��������
'�@�@�@�Fsmimax	�F�������
'�@�@�@�Fsdimin	�F��������
'�@�@�@�Fsdimax	�F�������
'�@�@�@�Fshimin	�F��������
'�@�@�@�Fshimax	�F�������
'�@�@�@�Fswsh	�F�A�ƊJ�n���ԁi���j
'�@�@�@�Fswsm	�F�A�ƊJ�n���ԁi���j
'�@�@�@�Fsweh	�F�A�ƏI�����ԁi���j
'�@�@�@�Fswem	�F�A�ƏI�����ԁi���j
'�@�@�@�Fswht	�F�T�x���
'�@�@�@�Fsage	�F�N��
'�@�@�@�Fsat	�F�_�����
'�@�@�@�Fslg1	�F���i�啪��
'�@�@�@�Fslc1	�F���i������
'�@�@�@�Fsl1	�F���i������
'�@�@�@�Fsos1	�F�n�r
'�@�@�@�Fsap1	�F�A�v���P�[�V����
'�@�@�@�Fsdl1	�F�J������
'�@�@�@�Fsdb1	�F�f�[�^�x�[�X
'�@�@�@�Fskw	�F�������[�h
'�@�@�@�Fskwflg	�F�������[�h�t���O [1]OR [2]AND
'�@�@�@�Fsst	�F�����r�b�g������(000000)
'�@�@�@�Fpoc	�F���R�[�h�i�Ώۊ�Ƃ̋��l�[�ꗗ�p�j
'�@�@�@�Fsoc	�F���R�[�h�i�����j
'�@�@�@�Fsocs	�F���R�[�hCSV
'�@�@�@�Fslocc	�F�Г��Č��̑Ώۊ�ƃR�[�h
'�@�@�@�Fsnewfkouko	�F�L���V���t���O
'�@�@�@�F
'�@�@�@�F������ �J���^�������p�p�����[�^ (�X�g�A�h up_SearchOrder ���p)
'�@�@�@�Fjt		�F�E��啪�ރR�[�h
'�@�@�@�Fjt2	�F�E��R�[�h
'�@�@�@�Fac		�F�G���A�R�[�h ��2012/02/28 LIS K.Kokubo �폜
'�@�@�@�Fac2	�F�s���{���R�[�h
'�@�@�@�Fwt		�F�Ζ��`�ԃR�[�h
'�@�@�@�Fkw		�F�L�[���[�h
'�@�@�@�F
'�@�@�@�F������ ���c�[���p
'�@�@�@�Fboc	�F�O��\�����R�[�h
'�@�@�@�F
'�@�@�@�F������ �g�p���@
'�@�@�@�FDim oSOC
'�@�@�@�FDim sSQL
'�@�@�@�FSet oSOC = New clsSearchOrderCondition	'�������ꂽ���_�Ńp�����[�^�Ƃo�n�r�s�f�[�^����r�p�k����������Ă���
'�@�@�@�FoSOC.Top = 100	'SELECT��ŏ����ݒ�
'�@�@�@�FsSQL = oSOC.GetSQLOrderSearchDetail()	'�r�p�k���擾
'�@�@�@�F
'���@���F2007/04/05 LIS K.Kokubo �쐬
'�@�@�@�F2007/10/10 LIS K.Kokubo ���c�[���p�ϐ��ǉ�
'�@�@�@�F2007/10/31 LIS K.Kokubo TOP ??? �p�ϐ��ǉ�
'�@�@�@�F2008/01/15 LIS K.Kokubo �p�����[�^���N�G����
'�@�@�@�F2008/03/26 LIS K.Kokubo �o�^�������ǉ�
'�@�@�@�F2008/08/14 LIS M.Hayashi �����t���O�ǉ��ƃt���b�N�X�ړ�
'�@�@�@�F2009/11/17 LIS K.Kokubo ���^�⎞�Ԃ̕ϐ��ɑS�p�������������ꍇ�A���p�����ɕϊ�
'�@�@�@�F2010/10/08 LIS K.Kokubo �Г��Č��̑Ώۊ�ƃR�[�h�ǉ�
'�@�@�@�F2012/02/28 LIS K.Kokubo �G���A�R�[�h�̌������p��p�~
'�@�@�@�F2012/03/12 LIS K.Kokubo �N����p�~�����ƔN�����ǉ�
'******************************************************************************
Class clsSearchOrderCondition
	'�������������o�ϐ�
	Public Top						'SELECT�Ŏ擾���錏�� (SELECT TOP �� * FROM �`)
	PUblic SearchDetailFlag			'�ڍ׌����t���O
	Public OrderTypeFlag			'�Г��O�Č������t���O
	Public NewFlag					'�V���t���O
	Public JobTypeBigCode1			'��]�E��啪�ނP
	Public JobTypeCode1				'��]�E��P
	Public JobTypeBigCode2			'��]�E��啪�ނQ
	Public JobTypeCode2				'��]�E��Q
	Public JobTypeBigCode3			'��]�E��啪�ނR
	Public JobTypeCode3				'��]�E��R
	Public RailwayLineCode			'��]����(�J���}��؂�)
	Public StationCode				'��]�w(�J���}��؂�)
	Public PrefectureCode			'��]�s���{��(�J���}��؂�)
	Public City						'��]�s��S(�J���}��؂�)
	Public WorkingTypeCode1			'��]�Ζ��`�ԂP
	Public WorkingTypeCode2			'��]�Ζ��`�ԂQ
	Public WorkingTypeCode3			'��]�Ζ��`�ԂR
	Public IndustryTypeCode			'��]�Ǝ�(�J���}��؂� [XX,XX,XX])
	Public PercentagePayFlag		'������
	Public YearlyIncomeMin			'�N������
	Public YearlyIncomeMax			'�N�����
	Public MonthlyIncomeMin			'��������
	Public MonthlyIncomeMax			'�������
	Public DailyIncomeMin			'��������
	Public DailyIncomeMax			'�������
	Public HourlyIncomeMin			'��������
	Public HourlyIncomeMax			'�������
	Public WorkStartHour			'�A�ƊJ�n���ԁi���j
	Public WorkStartMinute			'�A�ƊJ�n���ԁi���j
	Public WorkEndHour				'�A�ƏI�����ԁi���j
	Public WorkEndMinute			'�A�ƏI�����ԁi���j
	Public WeeklyHolidayType		'�T�x���
	'Public Age						'�N��
	Public SchoolTypeCode			'���ƔN�����i�w���j
	Public GraduateYear				'���ƔN�����i���ƔN�j
	Public AgreementTerm			'�_�����
	Public LicenseCount				'���i����
	Public LicenseGroupCode			'���i�啪��
	Public LicenseCategoryCode		'���i������
	Public LicenseCode				'���i������
	Public OSCode					'�n�r�iCSV�j
	Public ApplicationCode			'�A�v���P�[�V�����iCSV�j
	Public DevelopmentLanguageCode	'�J������iCSV�j
	Public DatabaseCode			'�f�[�^�x�[�X�iCSV�j
	Public Keyword					'�������[�h
	Public KeywordFlag				'�������[�h�t���O [1]OR [2]AND
	Public PictureOrderCode			'���R�[�h�i�Ώۊ�Ƃ̋��l�[�ꗗ�p�j
	Public OrderCode				'���R�[�h�i�����j CSV��
	Public Specialty
	Public InexperiencedPersonFlag	'�����i���o�����}�j
	Public UtilizeLanguageFlag		'�����i��w���������j
	Public TempFlag					'�����i�h���j
	Public UITurnFlag				'�����i�t�h�^�[�����}�j
	Public ManyHolidayFlag			'�����i�x���P�Q�O���ȏ�j
	Public FlexFlag					'�����i�t���b�N�X�j
	Public NearStationFlag			'�����i�w�߁j
	Public NoSmokingFlag			'�����i�։��E�����j
	Public NewlyBuiltFlag			'�����i�V�z�j
	Public LandmarkFlag				'�����i���w�j
	Public RenovationFlag			'�����i���m�x�[�V�����j
	Public DesignersFlag			'�����i�f�U�C�i�[�Y�j
	Public CompanyCafeteriaFlag		'�����i�Ј��H���j
	Public ShortOvertimeFlag		'�����i�Z���Ԏc�Ɓj
	Public MaternityFlag			'�����i�Y�x��x���т���j
	Public DressFreeFlag			'�����i�������R�j
	Public MammyFlag				'�����i�}�}���}�j
	Public FixedTimeFlag			'�����i18���܂łɑގЁj
	Public ShortTimeFlag			'�����i�Z���ԘJ���j
	Public HandicappedFlag			'�����i��Q�Ҋ��}�j
	Public RentAllFlag				'�����i�Z���p�S�z�⏕����j
	Public RentPartFlag				'�����i�Z���p�ꕔ�⏕����j
	Public MealsFlag				'�����i�H���E�d���t���Č��j
	Public MealsAssistanceFlag		'�����i�H���⏕���x����j
	Public TrainingCostFlag			'�����i���C������x����j
	Public EntrepreneurCostFlag		'�����i�N�Ƌ@�ޕ⏕���x����j
	Public MoneyFlag				'�����i�����q�E�ᗘ�q�⏕���x����j
	Public LandShopFlag				'�����i�y�n�E�X�ܓ��񋟐��x����j
	Public FindJobFestiveFlag		'�����i�A�E���j�������x����j
	Public AppointmentFlag			'�����i���Ј��o�p���x����j
	Public SocietyInsuranceFlag		'�����i�Еۊ����j
	Public RegistDay				'�o�^��
	Public LISOrderCompanyCode		'�Г��Č��̑Ώۊ�ƃR�[�h
    Public NewKoukokuFlag          	'�L���V���t���O
    Public FeatureFlag          	'���W�����t���O

	'�J���^����������
	Public JT	'�E��啪�ރR�[�h
	Public JT2	'�E��R�[�h
	Public AC2	'�s���{���R�[�h
	Public WT	'�Ζ��`�ԃR�[�h
	Public KW	'�L�[���[�h

	'�s�n�o�̎ʐ^����
	Public POC

	'��������
	Public PC	'�s���{���R�[�h
	Public RC	'�����R�[�h
	Public SC	'�w�R�[�h

	'���W
	Public SP	'���W�R�[�h

	'���c�[��
	Public BOC	'�O��\�����̍ŐV���R�[�h

	'�R�[�h�Ή�����
	Public JobTypeBigName1	'��]�E��啪�ޖ��̂P
	Public JobTypeName1	'��]�E�햼�̂P
	Public JobTypeBigName2	'��]�E��啪�ޖ��̂Q
	Public JobTypeName2	'��]�E�햼�̂Q
	Public JobTypeBigName3	'��]�E��啪�ޖ��̂R
	Public JobTypeName3	'��]�E�햼�̂R
	Public RailwayLineName	'��]��������
	Public StationName
	Public AreaName
	Public PrefectureName
	Public WorkingTypeName1
	Public WorkingTypeName2
	Public WorkingTypeName3
	Public IndustryTypeName	'�Ǝ햼�z��
	Public WeeklyHolidayTypeName
	Public OSName
	Public ApplicationName
	Public DevelopmentLanguageName
	Public DatabaseName
	Public SchoolTypeName
	Public LicenseGroupName
	Public LicenseCategoryName
	Public LicenseName

	'���̑������o�ϐ�
	Public HtmlOrderSearch	'���������o�͂g�s�l�k��
	Public SQLOrderSearch	'�����r�p�k
	Public SQLWriteLog		'���O�������݂r�p�k

	'******************************************************************************
	'�T�@�v�F�R���X�g���N�^
	'�쐬�ҁFLis K.Kokubo
	'�쐬���F2007/04/04 Lis K.Kokubo
	'�X�@�V�F
	'���@�l�F
	'******************************************************************************
	Private Sub Class_Initialize()
		LicenseCount = 0

		'�p�����[�^���猟���������擾
		Call ReadParam()

		'�f�[�^�������`�F�b�N
		Call ChkData()

		'�R�[�h�Ή����̎擾
		Call SetNames()

		'���l�[����SQL����
		SQLOrderSearch = GetSQLOrderSearchDetail()

		'���O��������SQL����
		SQLWriteLog = GetSQLWriteLog()

		'���l�[���������o�͂g�s�l�k��
		HtmlOrderSearch = GetHtmlSearchCondition()

		'Response.Write SQLOrderSearch
	End Sub

	'******************************************************************************
	'�T�@�v�FGET�f�[�^�̓ǂݍ���
	'���@���F
	'���@�l�F
	'���@���F2007/04/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Private Sub ReadParam()
		Dim idx

		If GetForm("sdf", 2) <> "" Then SearchDetailFlag = GetForm("sdf", 2)
		If GetForm("sotf", 2) <> "" Then OrderTypeFlag = GetForm("sotf", 2)
		If GetForm("snewf", 2) <> "" Then NewFlag = GetForm("snewf", 2)
		If GetForm("sjtbig1", 2) <> "" Then JobTypeBigCode1 = GetForm("sjtbig1", 2)
		If GetForm("sjt1", 2) <> "" Then JobTypeCode1 = GetForm("sjt1", 2)
		If GetForm("sjtbig2", 2) <> "" Then JobTypeBigCode2 = GetForm("sjtbig2", 2)
		If GetForm("sjt2", 2) <> "" Then JobTypeCode2 = GetForm("sjt2", 2)
		If GetForm("sjtbig3", 2) <> "" Then JobTypeBigCode3 = GetForm("sjtbig3", 2)
		If GetForm("sjt3", 2) <> "" Then JobTypeCode3 = GetForm("sjt3", 2)
		If GetForm("src", 2) <> "" Then RailwayLineCode = Replace(GetForm("src", 2)," ","")
		If GetForm("ssc", 2) <> "" Then StationCode = Replace(GetForm("ssc", 2)," ","")
		If GetForm("spc", 2) <> "" Then PrefectureCode = Replace(GetForm("spc", 2)," ","")
		If GetForm("sct", 2) <> "" Then City = GetForm("sct", 2)
		If GetForm("swt1", 2) <> "" Then WorkingTypeCode1 = GetForm("swt1", 2)
		If GetForm("swt2", 2) <> "" Then WorkingTypeCode2 = GetForm("swt2", 2)
		If GetForm("swt3", 2) <> "" Then WorkingTypeCode3 = GetForm("swt3", 2)
		If GetForm("sit", 2) <> "" Then IndustryTypeCode = GetForm("sit", 2)
		If GetForm("ssp01", 2) <> "" Then InexperiencedPersonFlag = GetForm("ssp01", 2)
		If GetForm("ssp02", 2) <> "" Then UtilizeLanguageFlag = GetForm("ssp02", 2)
		If GetForm("ssp03", 2) <> "" Then TempFlag = GetForm("ssp03", 2)
		If GetForm("ssp04", 2) <> "" Then UITurnFlag = GetForm("ssp04", 2)
		If GetForm("ssp05", 2) <> "" Then ManyHolidayFlag = GetForm("ssp05", 2)
		If GetForm("ssp06", 2) <> "" Then FlexFlag = GetForm("ssp06", 2)
		If GetForm("ssp07", 2) <> "" Then NearStationFlag = GetForm("ssp07", 2)
		If GetForm("ssp08", 2) <> "" Then NoSmokingFlag = GetForm("ssp08", 2)
		If GetForm("ssp09", 2) <> "" Then NewlyBuiltFlag = GetForm("ssp09", 2)
		If GetForm("ssp10", 2) <> "" Then LandmarkFlag = GetForm("ssp10", 2)
		If GetForm("ssp11", 2) <> "" Then RenovationFlag = GetForm("ssp11", 2)
		If GetForm("ssp12", 2) <> "" Then DesignersFlag = GetForm("ssp12", 2)
		If GetForm("ssp13", 2) <> "" Then CompanyCafeteriaFlag = GetForm("ssp13", 2)
		If GetForm("ssp14", 2) <> "" Then ShortOvertimeFlag = GetForm("ssp14", 2)
		If GetForm("ssp15", 2) <> "" Then MaternityFlag = GetForm("ssp15", 2)
		If GetForm("ssp16", 2) <> "" Then DressFreeFlag = GetForm("ssp16", 2)
		If GetForm("ssp17", 2) <> "" Then MammyFlag = GetForm("ssp17", 2)
		If GetForm("ssp18", 2) <> "" Then FixedTimeFlag = GetForm("ssp18", 2)
		If GetForm("ssp19", 2) <> "" Then ShortTimeFlag = GetForm("ssp19", 2)
		If GetForm("ssp20", 2) <> "" Then HandicappedFlag = GetForm("ssp20", 2)
		If GetForm("ssp21", 2) <> "" Then RentAllFlag = GetForm("ssp21", 2)
		If GetForm("ssp22", 2) <> "" Then RentPartFlag = GetForm("ssp22", 2)
		If GetForm("ssp23", 2) <> "" Then MealsFlag = GetForm("ssp23", 2)
		If GetForm("ssp24", 2) <> "" Then MealsAssistanceFlag = GetForm("ssp24", 2)
		If GetForm("ssp25", 2) <> "" Then TrainingCostFlag = GetForm("ssp25", 2)
		If GetForm("ssp26", 2) <> "" Then EntrepreneurCostFlag = GetForm("ssp26", 2)
		If GetForm("ssp27", 2) <> "" Then MoneyFlag = GetForm("ssp27", 2)
		If GetForm("ssp28", 2) <> "" Then LandShopFlag = GetForm("ssp28", 2)
		If GetForm("ssp29", 2) <> "" Then FindJobFestiveFlag = GetForm("ssp29", 2)
		If GetForm("ssp30", 2) <> "" Then AppointmentFlag = GetForm("ssp30", 2)
		If GetForm("ssp31", 2) <> "" Then SocietyInsuranceFlag = GetForm("ssp31", 2)
		If GetForm("sppf", 2) <> "" Then PercentagePayFlag = GetForm("sppf", 2)
		If GetForm("syimin", 2) <> "" Then YearlyIncomeMin = Replace(Replace(GetForm("syimin", 2),",",""),"��","0000")
		If GetForm("syimax", 2) <> "" Then YearlyIncomeMax = Replace(Replace(GetForm("syimax", 2),",",""),"��","0000")
		If GetForm("smimin", 2) <> "" Then MonthlyIncomeMin = GetForm("smimin", 2)
		If GetForm("smimax", 2) <> "" Then MonthlyIncomeMax = GetForm("smimax", 2)
		If GetForm("sdimin", 2) <> "" Then DailyIncomeMin = GetForm("sdimin", 2)
		If GetForm("sdimax", 2) <> "" Then DailyIncomeMax = GetForm("sdimax", 2)
		If GetForm("shimin", 2) <> "" Then HourlyIncomeMin = GetForm("shimin", 2)
		If GetForm("shimax", 2) <> "" Then HourlyIncomeMax = GetForm("shimax", 2)
		If GetForm("swsh", 2) <> "" Then WorkStartHour = GetForm("swsh", 2)
		If GetForm("swsm", 2) <> "" Then WorkStartMinute = GetForm("swsm", 2)
		If GetForm("sweh", 2) <> "" Then WorkEndHour = GetForm("sweh", 2)
		If GetForm("swem", 2) <> "" Then WorkEndMinute = GetForm("swem", 2)
		If GetForm("swht", 2) <> "" Then WeeklyHolidayType = GetForm("swht", 2)
		'If GetForm("sage", 2) <> "" Then Age = GetForm("sage", 2)
		If GetForm("sstc", 2) <> "" Then SchoolTypeCode = GetForm("sstc", 2)
		If GetForm("sgy", 2) <> "" Then GraduateYear = GetForm("sgy", 2)
		If GetForm("sat", 2) <> "" Then AgreementTerm = GetForm("sat", 2)
		If GetForm("slocc",2) <> "" Then LISOrderCompanyCode = GetForm("slocc",2)
        If GetForm("snewfkouko", 2) <> "" Then NewKoukokuFlag = GetForm("snewfkouko", 2)
        If GetForm("FeatureFlag", 2) <> "" Then FeatureFlag = GetForm("FeatureFlag", 2)
        

		'<���i>
		idx = 0
		Do While (IsEmpty(Request.Querystring("slg"&idx+1)) = False Or IsEmpty(Request.Querystring("slc"&idx+1)) = False Or IsEmpty(Request.Querystring("sl"&idx+1)) = False)
			idx = idx + 1
		Loop
		LicenseCount = idx
		ReDim LicenseGroupCode(LicenseCount)
		ReDim LicenseCategoryCode(LicenseCount)
		ReDim LicenseCode(LicenseCount)
		ReDim LicenseGroupName(LicenseCount)
		ReDim LicenseCategoryName(LicenseCount)
		ReDim LicenseName(LicenseCount)
		For idx = 0 To LicenseCount - 1
			LicenseGroupCode(idx) = GetForm("slg"&idx+1, 2)
			LicenseCategoryCode(idx) = GetForm("slc"&idx+1, 2)
			LicenseCode(idx) = Right(GetForm("sl"&idx+1, 2),2)
		Next
		'</���i>

		If GetForm("sos", 2) <> "" Then OSCode = Replace(GetForm("sos", 2)," ","")
		If GetForm("sap", 2) <> "" Then ApplicationCode = Replace(GetForm("sap", 2)," ","")
		If GetForm("sdl", 2) <> "" Then DevelopmentLanguageCode = Replace(GetForm("sdl", 2)," ","")
		If GetForm("sdb", 2) <> "" Then DatabaseCode = Replace(GetForm("sdb", 2)," ","")
		If GetForm("skw", 2) <> "" Then Keyword = GetForm("skw", 2)
		If GetForm("skwflg", 2) <> "" Then KeywordFlag = GetForm("skwflg", 2)
		If GetForm("sst", 2) <> "" Then Specialty = GetForm("sst", 2)
		If GetForm("poc", 2) <> "" Then PictureOrderCode = GetForm("poc", 2)
		If GetForm("soc", 2) <> "" Then OrderCode = GetForm("soc", 2)
		If GetForm("srd", 2) <> "" Then RegistDay = GetForm("srd", 2)

		If IsRE(GetForm("sst", 2), "^[01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01][01]$", True) = True Then
			If Mid(GetForm("sst", 2), 1, 1) = "1" Then InexperiencedPersonFlag = "1"
			If Mid(GetForm("sst", 2), 2, 1) = "1" Then UtilizeLanguageFlag = "1"
			If Mid(GetForm("sst", 2), 3, 1) = "1" Then TempFlag = "1"
			If Mid(GetForm("sst", 2), 4, 1) = "1" Then UITurnFlag = "1"
			If Mid(GetForm("sst", 2), 5, 1) = "1" Then ManyHolidayFlag = "1"
			If Mid(GetForm("sst", 2), 6, 1) = "1" Then FlexFlag = "1"
			If Mid(GetForm("sst", 2), 7, 1) = "1" Then NearStationFlag = "1"
			If Mid(GetForm("sst", 2), 8, 1) = "1" Then NoSmokingFlag = "1"
			If Mid(GetForm("sst", 2), 9, 1) = "1" Then NewlyBuiltFlag = "1"
			If Mid(GetForm("sst", 2), 10, 1) = "1" Then LandmarkFlag = "1"
			If Mid(GetForm("sst", 2), 11, 1) = "1" Then RenovationFlag = "1"
			If Mid(GetForm("sst", 2), 12, 1) = "1" Then DesignersFlag = "1"
			If Mid(GetForm("sst", 2), 13, 1) = "1" Then CompanyCafeteriaFlag = "1"
			If Mid(GetForm("sst", 2), 14, 1) = "1" Then ShortOvertimeFlag = "1"
			If Mid(GetForm("sst", 2), 15, 1) = "1" Then MaternityFlag = "1"
			If Mid(GetForm("sst", 2), 16, 1) = "1" Then DressFreeFlag = "1"
			If Mid(GetForm("sst", 2), 17, 1) = "1" Then MammyFlag = "1"
			If Mid(GetForm("sst", 2), 18, 1) = "1" Then FixedTimeFlag = "1"
			If Mid(GetForm("sst", 2), 19, 1) = "1" Then ShortTimeFlag = "1"
			If Mid(GetForm("sst", 2), 20, 1) = "1" Then HandicappedFlag = "1"
			If Mid(GetForm("sst", 2), 21, 1) = "1" Then RentAllFlag = "1"
			If Mid(GetForm("sst", 2), 22, 1) = "1" Then RentPartFlag = "1"
			If Mid(GetForm("sst", 2), 23, 1) = "1" Then MealsFlag = "1"
			If Mid(GetForm("sst", 2), 24, 1) = "1" Then MealsAssistanceFlag = "1"
			If Mid(GetForm("sst", 2), 25, 1) = "1" Then TrainingCostFlag = "1"
			If Mid(GetForm("sst", 2), 26, 1) = "1" Then EntrepreneurCostFlag = "1"
			If Mid(GetForm("sst", 2), 27, 1) = "1" Then MoneyFlag = "1"
			If Mid(GetForm("sst", 2), 28, 1) = "1" Then LandShopFlag = "1"
			If Mid(GetForm("sst", 2), 29, 1) = "1" Then FindJobFestiveFlag = "1"
			If Mid(GetForm("sst", 2), 30, 1) = "1" Then AppointmentFlag = "1"
			If Mid(GetForm("sst", 2), 31, 1) = "1" Then SocietyInsuranceFlag = "1"
		End If

		'TOP:�E��啪��
		If GetForm("jt", 2) <> "" Then JobTypeBigCode1 = GetForm("jt", 2)
		'TOP:�E�포����
		If GetForm("jt2", 2) <> "" Then JobTypeCode1 = GetForm("jt2", 2)
		'TOP:�Ζ��`�ԃR�[�h
		If GetForm("wt", 2) <> "" Then WorkingTypeCode1 = GetForm("wt", 2)
		'TOP:�L�[���[�h
		If GetForm("kw", 2) <> "" Then Keyword = GetForm("kw", 2)
		'���W
		If GetForm("sp", 2) <> "" Then SP = GetForm("sp", 2)

		'���������i�p�����[�^�j
		If GetForm("pc", 2) <> "" Then PC = GetForm("pc", 2)
		If GetForm("rc", 2) <> "" Then RC = GetForm("rc", 2)
		If GetForm("sc", 2) <> "" Then SC = GetForm("sc", 2)

		'���c�[��
		BOC = GetForm("boc", 2)
		If BOC <> "" Then SearchDetailFlag = "1"
	End Sub

	'******************************************************************************
	'�T�@�v�F�p�����[�^���ƃ����o�ϐ���R�t���Ēl��ݒ肷��
	'���@���FvKey	�F
	'�@�@�@�FvValue	�F
	'�@�@�@�FvFlag	�F
	'���@�l�F
	'�X�@�V�F2010/11/06 LIS K.Kokubo
	'******************************************************************************
	Private Sub SetData_ParamPart(ByVal vKey, ByVal vValue)
		If Len(vValue) = 0 Then vValue = GetForm(vKey, 2)

		Select Case vKey
			Case "sdf": SearchDetailFlag = vValue
			Case "sotf": OrderTypeFlag = vValue
			Case "snewf": NewFlag = vValue
			Case "sjtbig1": JobTypeBigCode1 = vValue
			Case "sjt1": JobTypeCode1 = vValue
			Case "sjtbig2": JobTypeBigCode2 = vValue
			Case "sjt2": JobTypeCode2 = vValue
			Case "sjtbig3": JobTypeBigCode3 = vValue
			Case "sjt3": JobTypeCode3 = vValue
			Case "src": RailwayLineCode = Replace(vValue," ","")
			Case "ssc": StationCode = Replace(vValue," ","")
			Case "spc": PrefectureCode = Replace(vValue," ","")
			Case "sct": City = vValue
			Case "swt1": WorkingTypeCode1 = vValue
			Case "swt2": WorkingTypeCode2 = vValue
			Case "swt3": WorkingTypeCode3 = vValue
			Case "sit": IndustryTypeCode = vValue
			Case "ssp01": InexperiencedPersonFlag = vValue
			Case "ssp02": UtilizeLanguageFlag = vValue
			Case "ssp03": TempFlag = vValue
			Case "ssp04": UITurnFlag = vValue
			Case "ssp05": ManyHolidayFlag = vValue
			Case "ssp06": FlexFlag = vValue
			Case "ssp07": NearStationFlag = vValue
			Case "ssp08": NoSmokingFlag = vValue
			Case "ssp09": NewlyBuiltFlag = vValue
			Case "ssp10": LandmarkFlag = vValue
			Case "ssp11": RenovationFlag = vValue
			Case "ssp12": DesignersFlag = vValue
			Case "ssp13": CompanyCafeteriaFlag = vValue
			Case "ssp14": ShortOvertimeFlag = vValue
			Case "ssp15": MaternityFlag = vValue
			Case "ssp16": DressFreeFlag = vValue
			Case "ssp17": MammyFlag = vValue
			Case "ssp18": FixedTimeFlag = vValue
			Case "ssp19": ShortTimeFlag = vValue
			Case "ssp20": HandicappedFlag = vValue
			Case "ssp21": RentAllFlag = vValue
			Case "ssp22": RentPartFlag = vValue
			Case "ssp23": MealsFlag = vValue
			Case "ssp24": MealsAssistanceFlag = vValue
			Case "ssp25": TrainingCostFlag = vValue
			Case "ssp26": EntrepreneurCostFlag = vValue
			Case "ssp27": MoneyFlag = vValue
			Case "ssp28": LandShopFlag = vValue
			Case "ssp29": FindJobFestiveFlag = vValue
			Case "ssp30": AppointmentFlag = vValue
			Case "ssp31": SocietyInsuranceFlag = vValue
			Case "sppf": PercentagePayFlag = vValue
			Case "syimin": YearlyIncomeMin = vValue
			Case "syimax": YearlyIncomeMax = vValue
			Case "smimin": MonthlyIncomeMin = vValue
			Case "smimax": MonthlyIncomeMax = vValue
			Case "sdimin": DailyIncomeMin = vValue
			Case "sdimax": DailyIncomeMax = vValue
			Case "shimin": HourlyIncomeMin = vValue
			Case "shimax": HourlyIncomeMax = vValue
			Case "swsh": WorkStartHour = vValue
			Case "swsm": WorkStartMinute = vValue
			Case "sweh": WorkEndHour = vValue
			Case "swem": WorkEndMinute = vValue
			Case "swht": WeeklyHolidayType = vValue
			Case "sage": Age = vValue
			Case "sstc": SchoolTypeCode = vValue
			Case "sgy": GraduateYear = vValue
			Case "sat": AgreementTerm = vValue
			Case "slocc": LISOrderCompanyCode = vValue
			Case "sos": OSCode = vValue
			Case "sap": ApplicationCode = vValue
			Case "sdl": DevelopmentLanguageCode = vValue
			Case "sdb": DatabaseCode = vValue
			Case "skw": Keyword = vValue
			Case "skwflg": KeywordFlag = vValue
			Case "sst": Specialty = vValue
			Case "poc": PictureOrderCode = vValue
			Case "soc": OrderCode = vValue
			Case "srd": RegistDay = vValue
			Case "snewfkouko": NewKoukokuFlag = vValue
            Case "FeatureFlag": FeatureFlag = vValue
		End Select
	End Sub

	'******************************************************************************
	'�T�@�v�F�p�����[�^�����񂩂烁���o�ϐ��̐ݒ�
	'���@�l�F
	'���@���F2010/11/06 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Function SetData_Param(ByVal vParam)
		Dim idx
		Dim a1,a2

		If Len(vParam) = 0 Then Exit Function
		If Len(vParam) > 1 And Left(vParam,1) = "?" Then vParam = Mid(vParam, 2)

		If InStr(vParam,"&amp;") > 0 Then
			a1 = Split(vParam,"&amp;")
		Else
			a1 = Split(vParam,"&")
		End If

		For idx= LBound(a1) To UBound(a1)
			a2 = Split(a1(idx),"=")
			If UBound(a2) = 1 Then
				Call SetData_ParamPart(a2(0),a2(1))
			End If
		Next

		'<URL�G���R�[�h����Ă��镶������f�R�[�h>
		If City <> "" Then City = getURLDecode(HopeCity1,"sjis")
		If City <> "" Then City = getURLDecode(HopeCity2,"sjis")
		If KeyWord <> "" Then KeyWord = getURLDecode(KeyWord,"sjis")
		'</URL�G���R�[�h����Ă��镶������f�R�[�h>

		Call SetNames()
	End Function

	'******************************************************************************
	'�T�@�v�F�R�[�h�ɑΉ��������̂��擾����
	'���@���F
	'���@�l�F
	'���@���F2007/04/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Private Sub SetNames()
		Dim sSQL,oRS,flgQE,sError
		Dim idx,aValue,sXML

		'��]�E��P
		If IsRE(JobTypeBigCode1, "^\d\d$", True) = True Then
			'�啪��
			sSQL = "sp_GetListJobTypeBig '" & JobTypeBigCode1 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				JobTypeBigName1 = ChkStr(oRS.Collect("BigClassName"))
			End If
			Call RSClose(oRS)

			'������
			If IsRE(JobTypeCode1, "^\d\d\d\d\d\d\d$", True) = True Then
				sSQL = "sp_GetListJobType '" & Left(JobTypeCode1, 2) & "', '" & Mid(JobTypeCode1, 3, 2) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					JobTypeName1 = ChkStr(oRS.Collect("MiddleClassName"))
				End If
				Call RSClose(oRS)
			End If
		Else
			'�����ނ̂�
			If IsRE(JobTypeCode1, "^\d\d\d\d\d\d\d$", True) = True Then
				sSQL = "sp_GetListJobType '" & Left(JobTypeCode1, 2) & "', '" & Mid(JobTypeCode1, 3, 2) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					JobTypeBigName1 = ChkStr(oRS.Collect("BigClassName"))
					JobTypeName1 = ChkStr(oRS.Collect("MiddleClassName"))
				End If
				Call RSClose(oRS)
			End If
		End If

		'��]�E��Q
		If IsRE(JobTypeBigCode2, "^\d\d$", True) = True Then
			'�啪��
			sSQL = "sp_GetListJobTypeBig '" & JobTypeBigCode2 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				JobTypeBigName2 = ChkStr(oRS.Collect("BigClassName"))
			End If
			Call RSClose(oRS)

			'������
			If IsRE(JobTypeCode2, "^\d\d\d\d\d\d\d$", True) = True Then
				sSQL = "sp_GetListJobType '" & Left(JobTypeCode2, 2) & "', '" & Mid(JobTypeCode2, 3, 2) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					JobTypeName2 = ChkStr(oRS.Collect("MiddleClassName"))
				End If
				Call RSClose(oRS)
			End If
		Else
			'�����ނ̂�
			If IsRE(JobTypeCode2, "^\d\d\d\d\d\d\d$", True) = True Then
				sSQL = "sp_GetListJobType '" & Left(JobTypeCode2, 2) & "', '" & Mid(JobTypeCode2, 3, 2) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					JobTypeBigName2 = ChkStr(oRS.Collect("BigClassName"))
					JobTypeName2 = ChkStr(oRS.Collect("MiddleClassName"))
				End If
				Call RSClose(oRS)
			End If
		End If

		'��]�E��R
		If IsRE(JobTypeBigCode3, "^\d\d$", True) = True Then
			'�啪��
			sSQL = "sp_GetListJobTypeBig '" & JobTypeBigCode3 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				JobTypeBigName3 = ChkStr(oRS.Collect("BigClassName"))
			End If
			Call RSClose(oRS)

			'������
			If IsRE(JobTypeCode3, "^\d\d\d\d\d\d\d$", True) = True Then
				sSQL = "sp_GetListJobType '" & Left(JobTypeCode3, 2) & "', '" & Mid(JobTypeCode3, 3, 2) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					JobTypeName3 = ChkStr(oRS.Collect("MiddleClassName"))
				End If
				Call RSClose(oRS)
			End If
		Else
			'�����ނ̂�
			If IsRE(JobTypeCode3, "^\d\d\d\d\d\d\d$", True) = True Then
				sSQL = "sp_GetListJobType '" & Left(JobTypeCode3, 2) & "', '" & Mid(JobTypeCode3, 3, 2) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					JobTypeBigName3 = ChkStr(oRS.Collect("BigClassName"))
					JobTypeName3 = ChkStr(oRS.Collect("MiddleClassName"))
				End If
				Call RSClose(oRS)
			End If
		End If

		'��]����
		If RailwayLineCode <> "" Then
			aValue = Split(Replace(RailwayLineCode, " ", ""), ",")

			sXML = ""
			For idx = 0 To UBound(aValue)
				sXML = sXML & "<railwayline><railwaylinecode>" & aValue(idx) & "</railwaylinecode></railwayline>"
			Next
			sXML = "<root>" & sXML & "</root>"

			sSQL = "EXEC up_DtlRailwayLine_XML '" & sXML & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Do While GetRSState(oRS) = True
				If RailwayLineName <> "" Then RailwayLineName = RailwayLineName & ","
				RailwayLineName = RailwayLineName & ChkStr(oRS.Collect("RailwayLineName2"))

				oRS.MoveNext
			Loop
			Call RSClose(oRS)
		End If

		'��]�w
		If StationCode <> "" Then
			aValue = Split(Replace(StationCode, " ", ""), ",")

			sXML = ""
			For idx = 0 To UBound(aValue)
				sXML = sXML & "<station><stationcode>" & aValue(idx) & "</stationcode></station>"
			Next
			sXML = "<root>" & sXML & "</root>"

			sSQL = "EXEC up_DtlStation_XML '" & sXML & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Do While GetRSState(oRS) = True
				If StationName <> "" Then StationName = StationName & ","
				StationName = StationName & ChkStr(oRS.Collect("StationName"))

				oRS.MoveNext
			Loop
			Call RSClose(oRS)
		End If

		'�s���{��
		If PrefectureCode <> "" Then
			aValue = Split(Replace(PrefectureCode, " ", ""), ",")

			sXML = ""
			For idx = 0 To UBound(aValue)
				sXML = sXML & "<prefecture><prefecturecode>" & aValue(idx) & "</prefecturecode></prefecture>"
			Next
			sXML = "<root>" & sXML & "</root>"

			sSQL = "EXEC up_DtlPrefecture_XML '" & sXML & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Do While GetRSState(oRS) = True
				If PrefectureName <> "" Then PrefectureName = PrefectureName & ","
				PrefectureName = PrefectureName & ChkStr(oRS.Collect("PrefectureName"))

				oRS.MoveNext
			Loop
			Call RSClose(oRS)
		End If

		'�Ζ��`�ԂP
		If WorkingTypeCode1 <> "" Then
			WorkingTypeName1 = GetDetail("WorkingType", WorkingTypeCode1)
		End If

		'�Ζ��`�ԂQ
		If WorkingTypeCode2 <> "" Then
			WorkingTypeName2 = GetDetail("WorkingType", WorkingTypeCode2)
		End If

		'�Ζ��`�ԂR
		If WorkingTypeCode3 <> "" Then
			WorkingTypeName3 = GetDetail("WorkingType", WorkingTypeCode3)
		End If

		'�Ǝ�
		If IndustryTypeCode <> "" Then
			aValue = Split(Replace(IndustryTypeCode, " ", ""), ",")

			IndustryTypeName = ""
			For idx = 0 To UBound(aValue)
				If IndustryTypeName <> "" Then IndustryTypeName = IndustryTypeName & ","
				IndustryTypeName = IndustryTypeName & GetDetail("IndustryType", aValue(idx))
			Next
		End If

		'�T�x���
		If WeeklyHolidayType <> "" Then
			WeeklyHolidayTypeName = GetDetail("WeeklyHolidayType", WeeklyHolidayType)
		End If

		'�n�r
		If OSCode <> "" Then
			aValue = Split(Replace(OSCode, " ", ""), ",")
			For idx = 0 To UBound(aValue)
				If OSName <> "" Then OSName = OSName & ","
				OSName = OSName & GetDetail("OS", aValue(idx))
			Next
		End If

		'�A�v���P�[�V����
		If ApplicationCode <> "" Then
			aValue = Split(Replace(ApplicationCode, " ", ""), ",")
			For idx = 0 To UBound(aValue)
				If ApplicationName <> "" Then ApplicationName = ApplicationName & ","
				ApplicationName = ApplicationName & GetDetail("Application", aValue(idx))
			Next
		End If

		'�J������
		If DevelopmentLanguageCode <> "" Then
			aValue = Split(Replace(DevelopmentLanguageCode, " ", ""), ",")
			For idx = 0 To UBound(aValue)
				If DevelopmentLanguageName <> "" Then DevelopmentLanguageName = DevelopmentLanguageName & ","
				DevelopmentLanguageName = DevelopmentLanguageName & GetDetail("DevelopmentLanguage", aValue(idx))
			Next
		End If

		'�f�[�^�x�[�X
		If DatabaseCode <> "" Then
			aValue = Split(Replace(DatabaseCode, " ", ""), ",")
			For idx = 0 To UBound(aValue)
				If DatabaseName <> "" Then DatabaseName = DatabaseName & ","
				DatabaseName = DatabaseName & GetDetail("Database", aValue(idx))
			Next
		End If

		'�ŏI�w��
		If SchoolTypeCode <> "" Then
			'sSQL = "EXEC up_DtlSchoolType '" & SchoolType & "';"
			SchoolTypeName = GetDetail("SchoolType", SchoolTypeCode)
		End If

		'���i
		For idx = 0 To LicenseCount - 1
			If LicenseGroupCode(idx) <> "" Then
				'�啪��
				sSQL = "sp_GetListLicenseGroup '" & LicenseGroupCode(idx) & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					LicenseGroupName(idx) = ChkStr(oRS.Collect("GroupName"))
				End If
				Call RSClose(oRS)

				'������
				If IsRE(LicenseCategoryCode(idx), "^\d\d\d$", True) = True Then
					sSQL = "sp_GetListLicenseCategory '" & LicenseGroupCode(idx) & "', '" & LicenseCategoryCode(idx) & "'"
					flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
					If GetRSState(oRS) = True Then
						LicenseCategoryName(idx) = ChkStr(oRS.Collect("CategoryName"))
					End If
					Call RSClose(oRS)

					'������
					If IsRE(LicenseCode(idx), "^\d\d$", True) = True Then
						sSQL = "sp_GetListLicenseCode '" & LicenseGroupCode(idx) & "', '" & LicenseCategoryCode(idx) & "', '" & LicenseCode(idx) & "'"
						flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
						If GetRSState(oRS) = True Then
							LicenseName(idx) = ChkStr(oRS.Collect("Name"))
						End If
						Call RSClose(oRS)
					End If
				End If
			End If
		Next
	End Sub

	'******************************************************************************
	'�T�@�v�F�S�p�����𔼊p�����ɕϊ�
	'���@���F
	'���@�l�F
	'���@���F2009/11/18 LIS K.Kokubo �쐬
	'******************************************************************************
	Private Function ChgZenNum(ByVal vNum)
		Dim sChg
		ChgZenNum = ""

		sChg = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(CStr(vNum),"�O","0"),"�P","1"),"�Q","2"),"�R","3"),"�S","4"),"�T","5"),"�U","6"),"�V","7"),"�W","8"),"�X","9")
		If IsRE(sChg,0,False) = False Then Exit Function
		ChgZenNum = sChg
	End Function

	'******************************************************************************
	'�T�@�v�F�f�[�^�̐��������`�F�b�N
	'���@���F
	'���@�l�F
	'���@���F2007/04/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Private Sub ChkData()
		Dim aValue
		Dim idx
		Dim tmp

		'�������ƃ`�F�b�J�[�Ή�
		If IsRE(JobTypeCode1, "^\d\d$", True) = True Then
			JobTypeBigCode1 = JobTypeCode1
			JobTypeCode1 = ""
		End If

		'��]�Ǝ�J���}��؂�
		IndustryTypeCode = Replace(IndustryTypeCode, " ", "")

		'�A�ƊJ�n����
		If WorkStartHour <> "" Then
			If WorkStartMinute = "" Then WorkStartMinute = "00"
		ElseIf WorkStartMinute <> "" Then
			WorkStartMinute = ""
		End If

		'�A�ƏI������
		If WorkEndHour <> "" Then
			If WorkEndMinute = "" Then WorkEndMinute = "00"
		ElseIf WorkEndMinute <> "" Then
			WorkEndMinute = ""
		End If

		If JobTypeCode1 <> "" Then JobTypeBigCode1 = Left(JobTypeCode1, 2)
		If JobTypeCode2 <> "" Then JobTypeBigCode2 = Left(JobTypeCode2, 2)
		If JobTypeCode3 <> "" Then JobTypeBigCode3 = Left(JobTypeCode3, 2)

		'�N��
		'Age = ChgZenNum(Replace(Age,",",""))

		'���ƔN
		If IsNumber(GraduateYear,0,False) Then
			If CInt(GraduateYear) < 1900 Or CInt(GraduateYear) > 2099 Then GraduateYear = ""
		End If

		'���^
		YearlyIncomeMin = Replace(ChgZenNum(Replace(YearlyIncomeMin,",","")),"��","0000")
		YearlyIncomeMax = Replace(ChgZenNum(Replace(YearlyIncomeMax,",","")),"��","0000")
		MonthlyIncomeMin = ChgZenNum(Replace(MonthlyIncomeMin,",",""))
		MonthlyIncomeMax = ChgZenNum(Replace(MonthlyIncomeMax,",",""))
		DailyIncomeMin = ChgZenNum(Replace(DailyIncomeMin,",",""))
		DailyIncomeMax = ChgZenNum(Replace(DailyIncomeMax,",",""))
		HourlyIncomeMin = ChgZenNum(Replace(HourlyIncomeMin,",",""))
		HourlyIncomeMax = ChgZenNum(Replace(HourlyIncomeMax,",",""))
		WorkStartHour = ChgZenNum(WorkStartHour)
		WorkEndHour = ChgZenNum(WorkEndHour)

		'���R�[�hCSV
		If InStr(OrderCode, ",") > 0 Then
			tmp = ""
			aValue = Split(OrderCode, ",")
			For idx = 0 To UBound(aValue)
				If tmp <> "" Then tmp = tmp & ","
				tmp = tmp & aValue(idx)
			Next
			OrderCode = tmp
		End If

		'�����r�b�g������
		Specialty = ""
		If InexperiencedPersonFlag & UtilizeLanguageFlag & TempFlag & UITurnFlag & ManyHolidayFlag & FlexFlag & _
		NearStationFlag & NoSmokingFlag & NewlyBuiltFlag & LandmarkFlag & RenovationFlag & DesignersFlag & _
		CompanyCafeteriaFlag & ShortOvertimeFlag & MaternityFlag & DressFreeFlag & MammyFlag & FixedTimeFlag & _
		ShortTimeFlag & HandicappedFlag & RentAllFlag & RentPartFlag & MealsFlag & MealsAssistanceFlag & _
		TrainingCostFlag & EntrepreneurCostFlag & MoneyFlag & LandShopFlag & FindJobFestiveFlag & AppointmentFlag & SocietyInsuranceFlag <> "" Then
			If InexperiencedPersonFlag <> "" Then: Specialty = Specialty & InexperiencedPersonFlag: Else: Specialty = Specialty & "0": End If
			If UtilizeLanguageFlag <> "" Then: Specialty = Specialty & UtilizeLanguageFlag: Else: Specialty = Specialty & "0": End If
			If TempFlag <> "" Then: Specialty = Specialty & TempFlag: Else: Specialty = Specialty & "0": End If
			If UITurnFlag <> "" Then: Specialty = Specialty & UITurnFlag: Else: Specialty = Specialty & "0": End If
			If ManyHolidayFlag <> "" Then: Specialty = Specialty & ManyHolidayFlag: Else: Specialty = Specialty & "0": End If
			If FlexFlag <> "" Then: Specialty = Specialty & FlexFlag: Else: Specialty = Specialty & "0": End If
			If NearStationFlag <> "" Then: Specialty = Specialty & NearStationFlag: Else: Specialty = Specialty & "0": End If
			If NoSmokingFlag <> "" Then: Specialty = Specialty & NoSmokingFlag: Else: Specialty = Specialty & "0": End If
			If NewlyBuiltFlag <> "" Then: Specialty = Specialty & NewlyBuiltFlag: Else: Specialty = Specialty & "0": End If
			If LandmarkFlag <> "" Then: Specialty = Specialty & LandmarkFlag: Else: Specialty = Specialty & "0": End If
			If RenovationFlag <> "" Then: Specialty = Specialty & RenovationFlag: Else: Specialty = Specialty & "0": End If
			If DesignersFlag <> "" Then: Specialty = Specialty & DesignersFlag: Else: Specialty = Specialty & "0": End If
			If CompanyCafeteriaFlag <> "" Then: Specialty = Specialty & CompanyCafeteriaFlag: Else: Specialty = Specialty & "0": End If
			If ShortOvertimeFlag <> "" Then: Specialty = Specialty & ShortOvertimeFlag: Else: Specialty = Specialty & "0": End If
			If MaternityFlag <> "" Then: Specialty = Specialty & MaternityFlag: Else: Specialty = Specialty & "0": End If
			If DressFreeFlag <> "" Then: Specialty = Specialty & DressFreeFlag: Else: Specialty = Specialty & "0": End If
			If MammyFlag <> "" Then: Specialty = Specialty & MammyFlag: Else: Specialty = Specialty & "0": End If
			If FixedTimeFlag <> "" Then: Specialty = Specialty & FixedTimeFlag: Else: Specialty = Specialty & "0": End If
			If ShortTimeFlag <> "" Then: Specialty = Specialty & ShortTimeFlag: Else: Specialty = Specialty & "0": End If
			If HandicappedFlag <> "" Then: Specialty = Specialty & HandicappedFlag: Else: Specialty = Specialty & "0": End If
			If RentAllFlag <> "" Then: Specialty = Specialty & RentAllFlag: Else: Specialty = Specialty & "0": End If
			If RentPartFlag <> "" Then: Specialty = Specialty & RentPartFlag: Else: Specialty = Specialty & "0": End If
			If MealsFlag <> "" Then: Specialty = Specialty & MealsFlag: Else: Specialty = Specialty & "0": End If
			If MealsAssistanceFlag <> "" Then: Specialty = Specialty & MealsAssistanceFlag: Else: Specialty = Specialty & "0": End If
			If TrainingCostFlag <> "" Then: Specialty = Specialty & TrainingCostFlag: Else: Specialty = Specialty & "0": End If
			If EntrepreneurCostFlag <> "" Then: Specialty = Specialty & EntrepreneurCostFlag: Else: Specialty = Specialty & "0": End If
			If MoneyFlag <> "" Then: Specialty = Specialty & MoneyFlag: Else: Specialty = Specialty & "0": End If
			If LandShopFlag <> "" Then: Specialty = Specialty & LandShopFlag: Else: Specialty = Specialty & "0": End If
			If FindJobFestiveFlag <> "" Then: Specialty = Specialty & FindJobFestiveFlag: Else: Specialty = Specialty & "0": End If
			If AppointmentFlag <> "" Then: Specialty = Specialty & AppointmentFlag: Else: Specialty = Specialty & "0": End If
			If SocietyInsuranceFlag <> "" Then: Specialty = Specialty & SocietyInsuranceFlag: Else: Specialty = Specialty & "0": End If
		End If

		If JT <> "" Then JobTypeCode1 = JT
		If JT2 <> "" Then JobTypeCode1 = JT2
		If WT <> "" Then WorkingTypeCode1 = WT
		If KW <> "" Then Keyword = KW

		If PC <> "" Then PrefectureCode = PC
		If RC <> "" Then RailwayLineCode1 = RC
		If SC <> "" Then StationCode = SC
	End Sub

	'******************************************************************************
	'�T�@�v�F���d���ڍ׌����y�[�W�֓n��GET�p�����[�^�𐶐����Ď擾�B
	'���@���F
	'���@�l�F������
	'�@�@�@�F�p�����[�^���܂�URL�́AIE�̐�����2048�����܂łł���̂ŁA����ɍ��킹��B
	'���@���F2007/04/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Function GetSearchParam()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sError

		Dim sParam
		Dim idx

		GetSearchParam = ""

		If SearchDetailFlag <> "" Then sParam = sParam & "&sdf=" & SearchDetailFlag
		If OrderTypeFlag <> "" Then sParam = sParam & "&sotf=" & OrderTypeFlag
		If NewFlag <> "" Then sParam = sParam & "&snewf=" & NewFlag
		If JobTypeBigCode1 <> "" Then sParam = sParam & "&sjtbig1=" & JobTypeBigCode1
		If JobTypeCode1 <> "" Then sParam = sParam & "&sjt1=" & JobTypeCode1
		If JobTypeBigCode2 <> "" Then sParam = sParam & "&sjtbig2=" & JobTypeBigCode2
		If JobTypeCode2 <> "" Then sParam = sParam & "&sjt2=" & JobTypeCode2
		If JobTypeBigCode3 <> "" Then sParam = sParam & "&sjtbig3=" & JobTypeBigCode3
		If JobTypeCode3 <> "" Then sParam = sParam & "&sjt3=" & JobTypeCode3
		If RailwayLineCode <> "" Then sParam = sParam & "&src=" & RailwayLineCode
		If StationCode <> "" Then sParam = sParam & "&ssc=" & StationCode
		If PrefectureCode <> "" Then sParam = sParam & "&spc=" & PrefectureCode
		If City <> "" Then sParam = sParam & "&sct=" & Server.URLEncode(City)
		If WorkingTypeCode1 <> "" Then sParam = sParam & "&swt1=" & WorkingTypeCode1
		If WorkingTypeCode2 <> "" Then sParam = sParam & "&swt2=" & WorkingTypeCode2
		If WorkingTypeCode3 <> "" Then sParam = sParam & "&swt3=" & WorkingTypeCode3
		If IndustryTypeCode <> "" Then sParam = sParam & "&sit=" & IndustryTypeCode
		If InexperiencedPersonFlag <> "" Then sParam = sParam & "&ssp01=" & InexperiencedPersonFlag
		If UtilizeLanguageFlag <> "" Then sParam = sParam & "&ssp02=" & UtilizeLanguageFlag
		If TempFlag <> "" Then sParam = sParam & "&ssp03=" & TempFlag
		If UITurnFlag <> "" Then sParam = sParam & "&ssp04=" & UITurnFlag
		If ManyHolidayFlag <> "" Then sParam = sParam & "&ssp05=" & ManyHolidayFlag
		If FlexFlag <> "" Then sParam = sParam & "&ssp06=" & FlexFlag
		If NearStationFlag <> "" Then sParam = sParam & "&ssp07=" & NearStationFlag
		If NoSmokingFlag <> "" Then sParam = sParam & "&ssp08=" & NoSmokingFlag
		If NewlyBuiltFlag <> "" Then sParam = sParam & "&ssp09=" & NewlyBuiltFlag
		If LandmarkFlag <> "" Then sParam = sParam & "&ssp10=" & LandmarkFlag
		If RenovationFlag <> "" Then sParam = sParam & "&ssp11=" & RenovationFlag
		If DesignersFlag <> "" Then sParam = sParam & "&ssp12=" & DesignersFlag
		If CompanyCafeteriaFlag <> "" Then sParam = sParam & "&ssp13=" & CompanyCafeteriaFlag
		If ShortOvertimeFlag <> "" Then sParam = sParam & "&ssp14=" & ShortOvertimeFlag
		If MaternityFlag <> "" Then sParam = sParam & "&ssp15=" & MaternityFlag
		If DressFreeFlag <> "" Then sParam = sParam & "&ssp16=" & DressFreeFlag
		If MammyFlag <> "" Then sParam = sParam & "&ssp17=" & MammyFlag
		If FixedTimeFlag <> "" Then sParam = sParam & "&ssp18=" & FixedTimeFlag
		If ShortTimeFlag <> "" Then sParam = sParam & "&ssp19=" & ShortTimeFlag
		If HandicappedFlag <> "" Then sParam = sParam & "&ssp20=" & HandicappedFlag
		If RentAllFlag <> "" Then sParam = sParam & "&ssp21=" & RentAllFlag
		If RentPartFlag <> "" Then sParam = sParam & "&ssp22=" & RentPartFlag
		If MealsFlag <> "" Then sParam = sParam & "&ssp23=" & MealsFlag
		If MealsAssistanceFlag <> "" Then sParam = sParam & "&ssp24=" & MealsAssistanceFlag
		If TrainingCostFlag <> "" Then sParam = sParam & "&ssp25=" & TrainingCostFlag
		If EntrepreneurCostFlag <> "" Then sParam = sParam & "&ssp26=" & EntrepreneurCostFlag
		If MoneyFlag <> "" Then sParam = sParam & "&ssp27=" & MoneyFlag
		If LandShopFlag <> "" Then sParam = sParam & "&ssp28=" & LandShopFlag
		If FindJobFestiveFlag <> "" Then sParam = sParam & "&ssp29=" & FindJobFestiveFlag
		If AppointmentFlag <> "" Then sParam = sParam & "&ssp30=" & AppointmentFlag
		If SocietyInsuranceFlag <> "" Then sParam = sParam & "&ssp31=" & SocietyInsuranceFlag
		If PercentagePayFlag <> "" Then sParam = sParam & "&sppf=" & PercentagePayFlag
		If YearlyIncomeMin <> "" Then sParam = sParam & "&syimin=" & YearlyIncomeMin
		If YearlyIncomeMax <> "" Then sParam = sParam & "&syimax=" & YearlyIncomeMax
		If MonthlyIncomeMin <> "" Then sParam = sParam & "&smimin=" & MonthlyIncomeMin
		If MonthlyIncomeMax <> "" Then sParam = sParam & "&smimax=" & MonthlyIncomeMax
		If DailyIncomeMin <> "" Then sParam = sParam & "&sdimin=" & DailyIncomeMin
		If DailyIncomeMax <> "" Then sParam = sParam & "&sdimax=" & DailyIncomeMax
		If HourlyIncomeMin <> "" Then sParam = sParam & "&shimin=" & HourlyIncomeMin
		If HourlyIncomeMax <> "" Then sParam = sParam & "&shimax=" & HourlyIncomeMax
		If WorkStartHour <> "" Then sParam = sParam & "&swsh=" & WorkStartHour
		If WorkStartMinute <> "" Then sParam = sParam & "&swsm=" & WorkStartMinute
		If WorkEndHour <> "" Then sParam = sParam & "&sweh=" & WorkEndHour
		If WorkEndMinute <> "" Then sParam = sParam & "&swem=" & WorkEndMinute
		If WeeklyHolidayType <> "" Then sParam = sParam & "&swht=" & WeeklyHolidayType
		'If Age <> "" Then sParam = sParam & "&sage=" & Age
		If SchoolTypeCode <> "" Then sParam = sParam & "&sstc=" & SchoolTypeCode
		If GraduateYear <> "" Then sParam = sParam & "&sgy=" & GraduateYear
		If AgreementTerm <> "" Then sParam = sParam & "&sat=" & AgreementTerm
		If NewKoukokuFlag <> "" Then sParam = sParam & "&snewfkouko=" & NewKoukokuFlag
        If FeatureFlag <> "" Then sParam = sParam & "&FeatureFlag=" & FeatureFlag

		For idx = 0 To LicenseCount - 1
			If LicenseGroupCode(idx) <> "" Then
				sParam = sParam & "&slg"&idx+1 & "=" & LicenseGroupCode(idx)
				sParam = sParam & "&slc"&idx+1 & "=" & LicenseCategoryCode(idx)
				sParam = sParam & "&sl"&idx+1 & "=" & LicenseCode(idx)
			End If
		Next

		If OSCode <> "" Then sParam = sParam & "&sos=" & OSCode
		If ApplicationCode <> "" Then sParam = sParam & "&sap=" & ApplicationCode
		If DevelopmentLanguageCode <> "" Then sParam = sParam & "&sdl=" & DevelopmentLanguageCode
		If DatabaseCode <> "" Then sParam = sParam & "&sdb=" & DatabaseCode
		If Keyword <> "" Then sParam = sParam & "&skw=" & Server.URLEncode(Keyword)
		If KeywordFlag <> "" Then sParam = sParam & "&skwflg=" & KeywordFlag
		If PictureOrderCode <> "" Then sParam = sParam & "&poc=" & PictureOrderCode
		If OrderCode <> "" Then sParam = sParam & "&soc=" & OrderCode
		If Specialty <> "" Then sParam = sParam & "&sst=" & Specialty
		If SP <> "" Then sParam = sParam & "&sp=" & SP
		If RegistDay <> "" Then sParam = sParam & "&srd=" & RegistDay
		If LISOrderCompanyCode <> "" Then sParam = sParam & "&slocc=" & LISOrderCompanyCode

		If sParam <> "" Then
			'����&���H�ɕϊ�
			sParam = "?" & Mid(sParam, 2)

			'�h�d�̎d�l�̓p�����[�^�̏�����Q�O�S�W�o�C�g
			sParam = Left(sParam, 2048)
		End If

		GetSearchParam = Replace(sParam, "&", "&amp;")
	End Function

	'******************************************************************************
	'�T�@�v�F���l�[�ڍ׌����r�p�k���擾
	'���@���F
	'���@�l�F
	'���@���F2007/04/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Function GetSQLOrderSearchDetail()
		Dim sSQL

		Dim sJoin
		Dim sWhere
		Dim sDeclare
		Dim sParams
		Dim iParamNo
		Dim iParamNo2
		Dim sFrom
		Dim sTemp
		Dim sTemp2
		Dim sTemp3
		Dim aValue
		Dim idx
		Dim sSearchCondition

		sJoin = ""
		sWhere = ""
		sDeclare = ""
		sParams = ""

		'�f�[�^�������`�F�b�N
		Call ChkData()

		'******************************************************************************
		'�Г��O�Č������t���O start
		'------------------------------------------------------------------------------
		If OrderTypeFlag <> "" Then
			If OrderTypeFlag = "0" Then
				'��ʋ��l�L��
				If sWhere <> "" Then sWhere = sWhere & "AND "
				sWhere = sWhere & "VWOC.OrderType = '0'" & vbCrLf
			ElseIf OrderTypeFlag = "1" Then
				'�Г��Č�
				If sWhere <> "" Then sWhere = sWhere & "AND "
				sWhere = sWhere & "VWOC.OrderType > '0'" & vbCrLf
			End If
		End If
		'------------------------------------------------------------------------------
		'�Г��O�Č������t���O end
		'******************************************************************************

		'******************************************************************************
		'�V���t���O start
		'------------------------------------------------------------------------------
		If NewFlag = "1" Then
			If sWhere <> "" Then sWhere = sWhere & "AND "
			sWhere = sWhere & "CONVERT(VARCHAR(8), VWOC.RegistDay, 112) >= DATEADD(DAY,  -9, CONVERT(DATETIME, CONVERT(VARCHAR(8), GETDATE(), 112))) "
			'���C�Z���X�̌f�ڊJ�n�����l������o�[�W�����̓R�����g�A�E�g
			'sWhere = sWhere & "(CASE WHEN VWOC.OrderType = '0' AND VWOC.RegistDay < VWOC.RiyoFromDate THEN VWOC.RiyoFromDate ELSE CONVERT(DATETIME, CONVERT(VARCHAR(8), VWOC.RegistDay, 112)) END) >= DATEADD(DAY, -6, CONVERT(DATETIME, CONVERT(VARCHAR(8), GETDATE(), 112))) "
		End If
		'------------------------------------------------------------------------------
		'�Г��O�Č������t���O end
		'******************************************************************************
        '******************************************************************************
		'�V���t���O start
		'------------------------------------------------------------------------------
		If NewKoukokuFlag = "1" Then
			If sWhere <> "" Then sWhere = sWhere & "AND "
			sWhere = sWhere & "VWOC.OrderType = '0' AND (CONVERT(VARCHAR(8), VWOC.RegistDay, 112) >= DATEADD(DAY,  -30, CONVERT(DATETIME, CONVERT(VARCHAR(8), GETDATE(), 112)))) "
		End If
		'------------------------------------------------------------------------------
		'�V���t���O end
		'******************************************************************************

		'******************************************************************************
		'�E�� start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		iParamNo = 0
		If JobTypeBigCode1 & JobTypeCode1 & JobTypeBigCode2 & JobTypeCode2 & JobTypeBigCode3 & JobTypeCode3 <> "" Then
			If JobTypeBigCode1 & JobTypeCode1 <> "" Then
				sTemp = JobTypeBigCode1
				If JobTypeCode1 <> "" Then sTemp = JobTypeCode1

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "A.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " + '%' "

				iParamNo = iParamNo + 1
			End If

			If JobTypeBigCode2 & JobTypeCode2 <> "" Then
				sTemp = JobTypeBigCode2
				If JobTypeCode2 <> "" Then sTemp = JobTypeCode2

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "A.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " + '%' "

				iParamNo = iParamNo + 1
			End If

			If JobTypeBigCode3 & JobTypeCode3 <> "" Then
				sTemp = JobTypeBigCode3
				If JobTypeCode3 <> "" Then sTemp = JobTypeCode3

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "A.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " + '%' "

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT A.OrderCode "
			sJoin = sJoin & "FROM C_JobType AS A WITH(NOLOCK) "
			sJoin = sJoin & "WHERE (" & RTrim(sTemp2) & ") "
			sJoin = sJoin & ") AS CJT ON VWOC.OrderCode = CJT.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'�E�� end
		'******************************************************************************

		'******************************************************************************
		'���� start
		'------------------------------------------------------------------------------
		sTemp = ""

		If RailwayLineCode <> "" Then
			aValue = Split(Replace(RailwayLineCode, " ", ""), ",")
			For iParamNo = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vRailwayLineCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vRailwayLineCode" & iParamNo & " = N'" & aValue(iParamNo) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vRailwayLineCode" & iParamNo
			Next

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT A.OrderCode "
			sJoin = sJoin & "FROM C_NearbyStation AS A WITH(NOLOCK) "
			sJoin = sJoin & "INNER JOIN StationStop AS B WITH(NOLOCK) "
			sJoin = sJoin & "ON A.StationCode = B.StationCode "
			sJoin = sJoin & "INNER JOIN B_RailwayLine AS C WITH(NOLOCK) "
			sJoin = sJoin & "ON B.RailwayLineCode = C.RailwayLineCode "
			sJoin = sJoin & "AND C.RailwayLineCode IN (" & RTrim(sTemp) & ")"
			sJoin = sJoin & ") AS CRL "
			sJoin = sJoin & "ON VWOC.OrderCode = CRL.OrderCode" & vbCrLf
		End If

		'------------------------------------------------------------------------------
		'���� end
		'******************************************************************************

		'******************************************************************************
		'�w start
		'------------------------------------------------------------------------------
		sTemp = ""

		If StationCode <> "" Then
			aValue = Split(Replace(StationCode, " ", ""), ",")
			For iParamNo = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vStationCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vStationCode" & iParamNo & " = N'" & aValue(iParamNo) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vStationCode" & iParamNo
			Next

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT A.OrderCode "
			sJoin = sJoin & "FROM C_NearbyStation AS A WITH(NOLOCK) "
			sJoin = sJoin & "WHERE A.StationCode IN (" & sTemp & ")"
			sJoin = sJoin & ") AS CNS "
			sJoin = sJoin & "ON VWOC.OrderCode = CNS.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'�w end
		'******************************************************************************

		'******************************************************************************
		'��]�Ζ��n start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""

		If PrefectureCode <> "" Or City <> "" Then
			If PrefectureCode <> "" Then
				aValue = Split(Replace(PrefectureCode, " ", ""), ",")
				For iParamNo = LBound(aValue) To UBound(aValue)
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vPrefectureCode" & iParamNo & " VARCHAR(3)"
					sParams = sParams & ",@vPrefectureCode" & iParamNo & " = N'" & aValue(iParamNo) & "'"

					If sTemp <> "" Then sTemp = sTemp & ","
					sTemp = sTemp & "@vPrefectureCode" & iParamNo
				Next

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "AND "
				sTemp2 = sTemp2 & "A.PrefectureCode IN (" & sTemp & ") "
			End If

			If City <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vCity VARCHAR(200)"
				sParams = sParams & ",@vCity = N'" & City & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "AND "
				sTemp2 = sTemp2 & "A.City LIKE '%' + @vCity + '%' "
			End If

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT A.OrderCode "
			sJoin = sJoin & "FROM C_WorkingPlace AS A WITH(NOLOCK) "
			sJoin = sJoin & "WHERE " & RTrim(sTemp2)
			sJoin = sJoin & ") AS CWP "
			sJoin = sJoin & "ON VWOC.OrderCode = CWP.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'��]�Ζ��n end
		'******************************************************************************

		'******************************************************************************
		'��]�Ζ��`�� start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If WorkingTypeCode1 & WorkingTypeCode2 & WorkingTypeCode3 <> "" Then
			If WorkingTypeCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkingTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vWorkingTypeCode" & iParamNo & " = N'" & WorkingTypeCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vWorkingTypeCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			If WorkingTypeCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkingTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vWorkingTypeCode" & iParamNo & " = N'" & WorkingTypeCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vWorkingTypeCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			If WorkingTypeCode3 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkingTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vWorkingTypeCode" & iParamNo & " = N'" & WorkingTypeCode3 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vWorkingTypeCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT A.OrderCode "
			sJoin = sJoin & "FROM C_WorkingType AS A WITH(NOLOCK) "
			sJoin = sJoin & "WHERE A.WorkingTypeCode IN (" & RTrim(sTemp) & ") "
			sJoin = sJoin & ") AS CWT "
			sJoin = sJoin & "ON VWOC.OrderCode = CWT.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'��]�Ζ��`�� end
		'******************************************************************************

		'******************************************************************************
		'��]�Ǝ� start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If IndustryTypeCode <> "" Then
			aValue = Split(Replace(IndustryTypeCode, " ", ""), ",")
			For iParamNo = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vIndustryTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vIndustryTypeCode" & iParamNo & " = N'" & aValue(iParamNo) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vIndustryTypeCode" & iParamNo
			Next

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT A.CompanyCode "
			sJoin = sJoin & "FROM CompanyInfo AS A WITH(NOLOCK) "
			sJoin = sJoin & "WHERE A.IndustryType IN (" & RTrim(sTemp) & ") "
			sJoin = sJoin & ") AS CIDST "
			sJoin = sJoin & "ON VWOC.CompanyCode = CIDST.CompanyCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'��]�Ǝ� end
		'******************************************************************************

		'******************************************************************************
		'���� start
		'------------------------------------------------------------------------------
		'���o�����}�A��w���������AUI�^�[���A�x���P�Q�O���ȏ�
		sTemp = ""

		If InexperiencedPersonFlag = "1" Or UtilizeLanguageFlag = "1" Or UITurnFlag = "1" Or ManyHolidayFlag = "1" Or _
		FlexFlag = "1" Or NearStationFlag = "1" Or NoSmokingFlag = "1" Or NewlyBuiltFlag = "1" Or LandmarkFlag = "1" Or _
		RenovationFlag = "1" Or DesignersFlag = "1" Or CompanyCafeteriaFlag = "1" Or ShortOvertimeFlag = "1" Or MaternityFlag = "1" Or _
		DressFreeFlag = "1" Or MammyFlag = "1" Or FixedTimeFlag = "1" Or ShortTimeFlag = "1" Or HandicappedFlag = "1" Or RentAllFlag = "1" Or _
		RentPartFlag = "1" Or MealsFlag = "1" Or MealsAssistanceFlag = "1" Or TrainingCostFlag = "1" Or EntrepreneurCostFlag = "1" Or _
		MoneyFlag = "1" Or LandShopFlag = "1" Or FindJobFestiveFlag = "1" Or AppointmentFlag = "1" Or SocietyInsuranceFlag = "1" Then
			'���o�����}
			If InexperiencedPersonFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.InexperiencedPersonFlag = '1' "
			End If

			'��w��������
			If UtilizeLanguageFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.UtilizeLanguageFlag = '1' "
			End If

			'UI�^�[��
			If UITurnFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.UITurnFlag = '1' "
			End If

			'�x���P�Q�O���ȏ�
			If ManyHolidayFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.ManyHolidayFlag = '1' "
			End If
			'�t���b�N�X�^�C��
			If FlexFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.FlexTimeFlag = '1' "
			End If
			'�w��
			If NearStationFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.NearStationFlag = '1' "
			End If
			'�։��E����
			If NoSmokingFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.NoSmokingFlag = '1' "
			End If
			'�V�z�r���E�I�t�B�X
			If NewlyBuiltFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.NewlyBuiltFlag = '1' "
			End If
			'���w�r��
			If LandmarkFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.LandmarkFlag = '1' "
			End If
			'���m�x�[�V����
			If RenovationFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.RenovationFlag = '1' "
			End If
			'�f�U�C�i�[�Y
			If DesignersFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.DesignersFlag = '1' "
			End If
			'�Ј��H��
			If CompanyCafeteriaFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.CompanyCafeteriaFlag = '1' "
			End If
			'�Z���Ԏc��
			If ShortOvertimeFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.ShortOvertimeFlag = '1' "
			End If
			'�Y�x�E��x
			If MaternityFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.MaternityFlag = '1' "
			End If
			'�������R
			If DressFreeFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.DressFreeFlag = '1' "
			End If
			'�}�}���}
			If MammyFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.MammyFlag = '1' "
			End If
			'18���܂łɑގ�
			If FixedTimeFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.FixedTimeFlag = '1' "
			End If
			'�Z���ԘJ��
			If ShortTimeFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.ShortTimeFlag = '1' "
			End If
			'��Q�Ҋ��}
			If HandicappedFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.HandicappedFlag = '1' "
			End If
			'�Z���p�S�z�⏕����
			If RentAllFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.RentAllFlag = '1' "
			End If
			'�Z���p�ꕔ�⏕����
			If RentPartFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.RentPartFlag = '1' "
			End If
			'�H���E�d���t���Č�
			If MealsFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.MealsFlag = '1' "
			End If
			'�H���⏕���x����
			If MealsAssistanceFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.MealsAssistanceFlag = '1' "
			End If
			'���C������x����
			If TrainingCostFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.TrainingCostFlag = '1' "
			End If
			'�N�Ƌ@�ޕ⏕���x����
			If EntrepreneurCostFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.EntrepreneurCostFlag = '1' "
			End If
			'�����q�E�ᗘ�q�⏕���x����
			If MoneyFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.MoneyFlag = '1' "
			End If
			'�y�n�E�X�ܓ��񋟐��x����
			If LandShopFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.LandShopFlag = '1' "
			End If
			'�A�E���j�������x����
			If FindJobFestiveFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.FindJobFestiveFlag = '1' "
			End If
			'���Ј��o�p���x����
			If AppointmentFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.AppointmentFlag = '1' "
			End If
			'�Еۊ���
			If SocietyInsuranceFlag = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "EXISTS(SELECT * FROM C_Info AS A WHERE CSP.OrderCode = A.OrderCode AND EXISTS(SELECT * FROM CompanyInfo AS B WHERE A.CompanyCode = B.CompanyCode AND B.SocietyInsurance = 'ON') AND EXISTS(SELECT * FROM C_WorkingType AS C WHERE A.OrderCode = C.OrderCode AND C.WorkingTypeCode <= '005') AND NOT EXISTS(SELECT * FROM C_WorkingType AS D WHERE A.OrderCode = D.OrderCode AND D.WorkingTypeCode IN ('006','007'))) "
			End If

			sJoin = sJoin & "INNER JOIN C_SupplementInfo AS CSP WITH(NOLOCK) ON VWOC.OrderCode = CSP.OrderCode AND " & RTrim(sTemp) & vbCrLf
		End If

'		'�h��
'		If TempFlag = "1" Then
'			If InStr(sJoin, "INNER JOIN C_WorkingType AS CWT") = 0 Then sJoin = sJoin & "INNER JOIN C_WorkingType AS CWT ON CI.OrderCode = CWT.OrderCode "
'			If sWhere <> "" Then sWhere = sWhere & "AND "
'			sWhere = sWhere & "CWT.WorkingTypeCode IN ('001', '004') " & vbCrLf
'		End If
		'------------------------------------------------------------------------------
		'���� end
		'******************************************************************************

		'******************************************************************************
		'���^ start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		If YearlyIncomeMin & YearlyIncomeMax & MonthlyIncomeMin & MonthlyIncomeMax & DailyIncomeMin & DailyIncomeMax & HourlyIncomeMin & HourlyIncomeMax & PercentagePayFlag <> "" Then
			'<�N��>
			If YearlyIncomeMin & YearlyIncomeMax <> "" Then
				If YearlyIncomeMin <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vYearlyIncomeMin INT"
					sParams = sParams & ",@vYearlyIncomeMin = " & YearlyIncomeMin
				End If

				If YearlyIncomeMax <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vYearlyIncomeMax INT"
					sParams = sParams & ",@vYearlyIncomeMax = " & YearlyIncomeMax
				End If

				If sTemp <> "" Then sTemp = sTemp & "OR "
				If YearlyIncomeMin <> "" And YearlyIncomeMax <> "" Then
					'�N������,��������̓��͂�����ꍇ
					sTemp = sTemp & "((COALESCE(A.YearlyIncomeMin, 0) > 0 AND (A.YearlyIncomeMin BETWEEN @vYearlyIncomeMin AND @vYearlyIncomeMax)) OR (COALESCE(A.YearlyIncomeMax, 0) > 0 AND (A.YearlyIncomeMax BETWEEN @vYearlyIncomeMin AND @vYearlyIncomeMax))) "
				ElseIf YearlyIncomeMin <> "" Then
					'�N�������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "((COALESCE(A.YearlyIncomeMin, 0) > 0 AND A.YearlyIncomeMin >= @vYearlyIncomeMin) OR (COALESCE(A.YearlyIncomeMax, 0) > 0 AND A.YearlyIncomeMax >= @vYearlyIncomeMin)) "
				ElseIf YearlyIncomeMax <> "" Then
					'�N������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "((COALESCE(A.YearlyIncomeMin, 0) > 0 AND A.YearlyIncomeMin <= @vYearlyIncomeMax) OR (COALESCE(A.YearlyIncomeMin, 0) = 0 AND COALESCE(A.YearlyIncomeMax, 0) > 0 AND A.YearlyIncomeMax <= @vYearlyIncomeMax)) "
				End If
			End If
			'</�N��>

			'<����>
			If MonthlyIncomeMin & MonthlyIncomeMax <> "" Then
				If MonthlyIncomeMin <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vMonthlyIncomeMin INT"
					sParams = sParams & ",@vMonthlyIncomeMin = " & MonthlyIncomeMin
				End If

				If MonthlyIncomeMax <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vMonthlyIncomeMax INT"
					sParams = sParams & ",@vMonthlyIncomeMax = " & MonthlyIncomeMax
				End If

				If sTemp <> "" Then sTemp = sTemp & "OR "
				If MonthlyIncomeMin <> "" And MonthlyIncomeMax <> "" Then
					'��������,��������̓��͂�����ꍇ
					sTemp = sTemp & "((COALESCE(A.MonthlyIncomeMin, 0) > 0 AND (A.MonthlyIncomeMin BETWEEN @vMonthlyIncomeMin AND @vMonthlyIncomeMax)) OR (COALESCE(A.MonthlyIncomeMax, 0) > 0 AND (A.MonthlyIncomeMax BETWEEN @vMonthlyIncomeMin AND @vMonthlyIncomeMax))) "
				ElseIf MonthlyIncomeMin <> "" Then
					'���������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "((COALESCE(A.MonthlyIncomeMin, 0) > 0 AND A.MonthlyIncomeMin >= @vMonthlyIncomeMin) OR (COALESCE(A.MonthlyIncomeMax, 0) > 0 AND A.MonthlyIncomeMax >= @vMonthlyIncomeMin)) "
				ElseIf MonthlyIncomeMax <> "" Then
					'��������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "((COALESCE(A.MonthlyIncomeMin, 0) > 0 AND A.MonthlyIncomeMin <= @vMonthlyIncomeMax) OR (COALESCE(A.MonthlyIncomeMin, 0) = 0 AND COALESCE(A.MonthlyIncomeMax, 0) > 0 AND A.MonthlyIncomeMax <= @vMonthlyIncomeMax)) "
				End If
			End If
			'</����>

			'<����>
			If DailyIncomeMin & DailyIncomeMax <> "" Then
				If DailyIncomeMin <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDailyIncomeMin INT"
					sParams = sParams & ",@vDailyIncomeMin = " & DailyIncomeMin
				End If

				If DailyIncomeMax <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDailyIncomeMax INT"
					sParams = sParams & ",@vDailyIncomeMax = " & DailyIncomeMax
				End If

				If sTemp <> "" Then sTemp = sTemp & "OR "
				If DailyIncomeMin <> "" And DailyIncomeMax <> "" Then
					'��������,��������̓��͂�����ꍇ
					sTemp = sTemp & "((COALESCE(A.DailyIncomeMin, 0) > 0 AND (A.DailyIncomeMin BETWEEN @vDailyIncomeMin AND @vDailyIncomeMax)) OR (COALESCE(A.DailyIncomeMax, 0) > 0 AND (A.DailyIncomeMax BETWEEN @vDailyIncomeMin AND @vDailyIncomeMax))) "
				ElseIf DailyIncomeMin <> "" Then
					'���������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "((COALESCE(A.DailyIncomeMin, 0) > 0 AND A.DailyIncomeMin >= @vDailyIncomeMin) OR (COALESCE(A.DailyIncomeMax, 0) > 0 AND A.DailyIncomeMax >= @vDailyIncomeMin)) "
				ElseIf DailyIncomeMax <> "" Then
					'��������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "((COALESCE(A.DailyIncomeMin, 0) > 0 AND A.DailyIncomeMin <= @vDailyIncomeMax) OR (COALESCE(A.DailyIncomeMin, 0) = 0 AND COALESCE(A.DailyIncomeMax, 0) > 0 AND A.DailyIncomeMax <= @vDailyIncomeMax)) "
				End If
			End If
			'</����>

			'<����>
			If HourlyIncomeMin & HourlyIncomeMax <> "" Then
				If HourlyIncomeMin <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vHourlyIncomeMin INT"
					sParams = sParams & ",@vHourlyIncomeMin = " & HourlyIncomeMin
				End If

				If HourlyIncomeMax <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vHourlyIncomeMax INT"
					sParams = sParams & ",@vHourlyIncomeMax = " & HourlyIncomeMax
				End If

				If sTemp <> "" Then sTemp = sTemp & "OR "
				If HourlyIncomeMin <> "" And HourlyIncomeMax <> "" Then
					'��������,��������̓��͂�����ꍇ
					sTemp = sTemp & "((COALESCE(A.HourlyIncomeMin, 0) > 0 AND (A.HourlyIncomeMin BETWEEN @vHourlyIncomeMin AND @vHourlyIncomeMax)) OR (COALESCE(A.HourlyIncomeMax, 0) > 0 AND (A.HourlyIncomeMax BETWEEN @vHourlyIncomeMin AND @vHourlyIncomeMax))) "
				ElseIf HourlyIncomeMin <> "" Then
					'���������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "((COALESCE(A.HourlyIncomeMin, 0) > 0 AND A.HourlyIncomeMin >= @vHourlyIncomeMin) OR (COALESCE(A.HourlyIncomeMax, 0) > 0 AND A.HourlyIncomeMax >= @vHourlyIncomeMin)) "
				ElseIf HourlyIncomeMax <> "" Then
					'��������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "((COALESCE(A.HourlyIncomeMin, 0) > 0 AND A.HourlyIncomeMin <= @vHourlyIncomeMax) OR (COALESCE(A.HourlyIncomeMin, 0) = 0 AND COALESCE(A.HourlyIncomeMax, 0) > 0 AND A.HourlyIncomeMax <= @vHourlyIncomeMax)) "
				End If
			End If
			'</����>

			If sTemp <> "" Then sTemp = "(" & sTemp & ") "

			'������
			If PercentagePayFlag <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vPercentagePayFlag VARCHAR(1)"
				sParams = sParams & ",@vPercentagePayFlag = N'" & PercentagePayFlag & "'"

				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = "A.PercentagePayFlag = @vPercentagePayFlag "
			End If

			sJoin = sJoin & "INNER JOIN (SELECT A.OrderCode FROM C_Info AS A WHERE " & RTrim(sTemp) & ") AS CSLY ON VWOC.OrderCode = CSLY.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'���^ end
		'******************************************************************************

		'******************************************************************************
		'�Ζ��J�n�E�I������ start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		If WorkStartHour & WorkEndHour <> "" Then
			If WorkStartHour <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkStartHour VARCHAR(2) "
				sParams = sParams & ",@vWorkStartHour = N'" & WorkStartHour & "'"

				If WorkStartMinute = "" Then WorkStartMinute = "00"

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkStartMinute VARCHAR(2) "
				sParams = sParams & ",@vWorkStartMinute = N'" & WorkStartMinute & "'"

				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "A.WorkStartTime >= @vWorkStartHour + @vWorkStartMinute "
			End If

			If WorkEndHour <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkEndHour VARCHAR(2) "
				sParams = sParams & ",@vWorkEndHour = N'" & WorkEndHour & "'"

				If WorkEndMinute = "" Then WorkEndMinute = "00"

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkEndMinute VARCHAR(2) "
				sParams = sParams & ",@vWorkEndMinute = N'" & WorkEndMinute & "'"

				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "A.WorkEndTime <= @vWorkEndHour + @vWorkEndMinute "
			End If

			If WorkStartHour <> "" And WorkEndHour <> "" Then
				If WorkStartHour < WorkEndHour Then
					'�Ζ��J�n���� < �Ζ��I�����Ԃ̏ꍇ�A��Ԃ̋Ɩ����Ԃ������悤�ɂ���
					sTemp2 = "AND A.WorkStartTime < A.WorkEndTime "
				End If
			End If

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_WorkingCondition AS A WITH(NOLOCK) WHERE " & sTemp & RTrim(sTemp2) & ") AS CWTM ON VWOC.OrderCode = CWTM.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'�Ζ��J�n�E�I������ end
		'******************************************************************************

		'******************************************************************************
		'�T�x start
		'------------------------------------------------------------------------------
		sTemp = ""
		If WeeklyHolidayType <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vWeeklyHolidayType VARCHAR(3) "
			sParams = sParams & ",@vWeeklyHolidayType = N'" & WeeklyHolidayType & "'"

			sTemp = sTemp & "CWHT.WeeklyHolidayType = @vWeeklyHolidayType "

			sJoin = sJoin & "INNER JOIN C_Info AS CWHT WITH(NOLOCK) ON VWOC.OrderCode = CWHT.OrderCode AND " & RTrim(sTemp) & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'�T�x end
		'******************************************************************************

		'******************************************************************************
		'�N�� start
		'------------------------------------------------------------------------------
		'sTemp = ""
		'If Age <> "" Then
		'	If sDeclare <> "" Then sDeclare = sDeclare & ","
		'	sDeclare = sDeclare & "@vAge INT "
		'	sParams = sParams & ",@vAge = " & Age

		'	sTemp = sTemp & "(@vAge BETWEEN ISNULL(CAGE.AgeMin, 0) AND ISNULL(CAGE.AgeMax, 255)) "

		'	sJoin = sJoin & "INNER JOIN C_Info AS CAGE WITH(NOLOCK) ON VWOC.OrderCode = CAGE.OrderCode AND " & RTrim(sTemp) & vbCrLf
		'End If
		'------------------------------------------------------------------------------
		' �N�� end
		'******************************************************************************

		'******************************************************************************
		'���ƔN���� start
		'------------------------------------------------------------------------------
		'sTemp = ""
		If CStr(GraduateYear) <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vGraduateYear SMALLINT "
			sParams = sParams & ",@vGraduateYear = " & GraduateYear

			sTemp = sTemp & "(@vGraduateYear BETWEEN A.YearMin AND A.YearMax) "

			sJoin = sJoin & "INNER JOIN (SELECT A.OrderCode FROM C_GraduateYear AS A WITH(NOLOCK) WHERE " & RTrim(sTemp) & ") AS CGY ON VWOC.OrderCode = CGY.OrderCode " & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'���ƔN���� end
		'******************************************************************************

		'******************************************************************************
		'�_����� start
		'------------------------------------------------------------------------------
		sTemp = ""
		If IsRE(AgreementTerm, "^[123]$", True) = True Then
			If AgreementTerm = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT OrderCode FROM C_Temp WITH(NOLOCK) WHERE WorkPeriod <= 1 UNION SELECT OrderCode FROM C_Undertake WITH(NOLOCK) WHERE WorkPeriod <= 1 UNION SELECT OrderCode FROM C_TTP WITH(NOLOCK) WHERE WorkPeriod <= 1) AS CAT ON VWOC.OrderCode = CAT.OrderCode" & vbCrLf
			ElseIf AgreementTerm = "2" Then
				sJoin = sJoin & "INNER JOIN (SELECT OrderCode FROM C_Temp WITH(NOLOCK) WHERE WorkPeriod <= 2 UNION SELECT OrderCode FROM C_Undertake WITH(NOLOCK) WHERE WorkPeriod <= 2 UNION SELECT OrderCode FROM C_TTP WITH(NOLOCK) WHERE WorkPeriod <= 2) AS CAT ON VWOC.OrderCode = CAT.OrderCode" & vbCrLf
			ElseIf AgreementTerm = "3" Then
				sJoin = sJoin & "INNER JOIN (SELECT OrderCode FROM C_Temp WITH(NOLOCK) WHERE WorkPeriod > 3 UNION SELECT OrderCode FROM C_Undertake WITH(NOLOCK) WHERE WorkPeriod > 3 UNION SELECT OrderCode FROM C_TTP WITH(NOLOCK) WHERE WorkPeriod > 3) AS CAT ON VWOC.OrderCode = CAT.OrderCode" & vbCrLf
			End If
		End If
		'------------------------------------------------------------------------------
		'�_����� end
		'******************************************************************************

		'******************************************************************************
		'�ۗL���i start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		iParamNo = 0
		If LicenseCount > 0 Then
			For idx = 0 To LicenseCount - 1
				sTemp = ""

				If LicenseGroupCode(idx) <> "" Then
					'�啪��
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vLicenseGroupCode" & iParamNo & " VARCHAR(2)"
					sParams = sParams & ",@vLicenseGroupCode" & iParamNo & " = N'" & LicenseGroupCode(idx) & "'"

					If sTemp <> "" Then sTemp = sTemp & "AND "
					sTemp = sTemp & "A.GroupCode = @vLicenseGroupCode" & iParamNo & " "

					'������
					If LicenseCategoryCode(idx) <> "" Then
						If sDeclare <> "" Then sDeclare = sDeclare & ","
						sDeclare = sDeclare & "@vLicenseCategoryCode" & iParamNo & " VARCHAR(3)"
						sParams = sParams & ",@vLicenseCategoryCode" & iParamNo & " = N'" & LicenseCategoryCode(idx) & "'"

						If sTemp <> "" Then sTemp = sTemp & "AND "
						sTemp = sTemp & "A.CategoryCode = @vLicenseCategoryCode" & iParamNo & " "
					End If

					'������
					If LicenseCode(idx) <> "" Then
						If sDeclare <> "" Then sDeclare = sDeclare & ","
						sDeclare = sDeclare & "@vLicenseCode" & iParamNo & " VARCHAR(2)"
						sParams = sParams & ",@vLicenseCode" & iParamNo & " = N'" & LicenseCode(idx) & "'"

						If sTemp <> "" Then sTemp = sTemp & "AND "
						sTemp = sTemp & "A.Code = @vLicenseCode" & iParamNo & " "
					End If

					If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
					sTemp2 = sTemp2 & "(" & Trim(sTemp) & ") "

					iParamNo = iParamNo + 1
				End If
			Next

			If sTemp2 <> "" Then sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_License AS A WITH(NOLOCK) WHERE " & RTrim(sTemp2) & ") AS CL ON VWOC.OrderCode = CL.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'�ۗL���i end
		'******************************************************************************

		'******************************************************************************
		'�X�L�� start
		'------------------------------------------------------------------------------
		iParamNo2 = 1
		'OS
		sTemp = ""
		iParamNo = 1
		If OSCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo2 & " VARCHAR(20)"
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo2 & " = N'OS'"

			aValue = Split(Replace(OSCode, " ", ""), ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSkillCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vSkillCode" & iParamNo & " = N'" & aValue(idx) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vSkillCode" & iParamNo

				iParamNo = iParamNo + 1
			Next

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_Skill AS A WITH(NOLOCK) WHERE A.CategoryCode = @vSkillCategoryCode" & iParamNo2 & " AND A.Code IN (" & Trim(sTemp) & ")) AS CSKL" & iParamNo2 & " ON VWOC.OrderCode = CSKL" & iParamNo2 & ".OrderCode" & vbCrLf
			iParamNo2 = iParamNo2 + 1
		End If

		'�A�v���P�[�V����
		sTemp = ""
		If ApplicationCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo2 & " VARCHAR(20)"
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo2 & " = N'Application'"

			aValue = Split(Replace(ApplicationCode, " ", ""), ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSkillCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vSkillCode" & iParamNo & " = N'" & aValue(idx) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vSkillCode" & iParamNo

				iParamNo = iParamNo + 1
			Next

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_Skill AS A WITH(NOLOCK) WHERE A.CategoryCode = @vSkillCategoryCode" & iParamNo2 & " AND A.Code IN (" & Trim(sTemp) & ")) AS CSKL" & iParamNo2 & " ON VWOC.OrderCode = CSKL" & iParamNo2 & ".OrderCode" & vbCrLf
			iParamNo2 = iParamNo2 + 1
		End If

		'�J������
		sTemp = ""
		If DevelopmentLanguageCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo2 & " VARCHAR(20)"
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo2 & " = N'DevelopmentLanguage'"

			aValue = Split(Replace(DevelopmentLanguageCode, " ", ""), ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSkillCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vSkillCode" & iParamNo & " = N'" & aValue(idx) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vSkillCode" & iParamNo

				iParamNo = iParamNo + 1
			Next

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_Skill AS A WITH(NOLOCK) WHERE A.CategoryCode = @vSkillCategoryCode" & iParamNo2 & " AND A.Code IN (" & Trim(sTemp) & ")) AS CSKL" & iParamNo2 & " ON VWOC.OrderCode = CSKL" & iParamNo2 & ".OrderCode" & vbCrLf
			iParamNo2 = iParamNo2 + 1
		End If

		'�f�[�^�x�[�X
		sTemp = ""
		If DatabaseCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo2 & " VARCHAR(20)"
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo2 & " = N'Database'"

			aValue = Split(Replace(DatabaseCode, " ", ""), ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSkillCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vSkillCode" & iParamNo & " = N'" & aValue(idx) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vSkillCode" & iParamNo

				iParamNo = iParamNo + 1
			Next

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.OrderCode FROM C_Skill AS A WITH(NOLOCK) WHERE A.CategoryCode = @vSkillCategoryCode" & iParamNo2 & " AND A.Code IN (" & Trim(sTemp) & ")) AS CSKL" & iParamNo2 & " ON VWOC.OrderCode = CSKL" & iParamNo2 & ".OrderCode" & vbCrLf
			iParamNo2 = iParamNo2 + 1
		End If
		'------------------------------------------------------------------------------
		'�X�L�� end
		'******************************************************************************

		'******************************************************************************
		'�L�[���[�h start
		'------------------------------------------------------------------------------
		sTemp = ""
		If Keyword <> "" Then
			aValue = Split(Replace(Replace(Replace(Keyword, "(", "�i"), ")", "�j"), "�@", " "), " ")
			For idx = LBound(aValue) To UBound(aValue)
				If sTemp <> "" Then
					If KeywordFlag = "1" Then
						sTemp = sTemp & " OR "
					ElseIf KeywordFlag = "2" Then
						sTemp = sTemp & " AND "
					Else
						sTemp = sTemp & " AND "
					End If
				End If
				sTemp = sTemp & "FORMSOF(THESAURUS, " & aValue(idx) & "*)"
			Next
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vKeyword VARCHAR(400)"
			sParams = sParams & ",@vKeyword = N'" & sTemp & "'"

			sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY A.OrderCode) AS Num, A.OrderCode FROM C_FullTextNavi AS A WITH(NOLOCK) WHERE CONTAINS(A.Text, @vKeyword)) AS CFTN ON VWOC.OrderCode = CFTN.OrderCode" & vbCrLf
            'sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY A.OrderCode) AS Num, A.OrderCode FROM C_FullTextNavi AS A WITH(NOLOCK) left join (SELECT A.OrderCode From C_info as A INNER JOIN CompanyInfo as B on A.CompanyCode = B.CompanyCode WHERE (b.CompanyName_K like @vKeyword OR b.CompanyName_F like @vKeyword)) as B ON A.OrderCode = B.OrderCode WHERE CONTAINS(A.Text, @vKeyword)) AS CFTN ON VWOC.OrderCode = CFTN.OrderCode" & vbCrLf
		
        End If
		'------------------------------------------------------------------------------
		'�L�[���[�h end
		'******************************************************************************

		'******************************************************************************
		'�Ώۊ�Ƃ̋��l�[�ꗗ�p���R�[�h start
		'------------------------------------------------------------------------------
		If PictureOrderCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vPictureOrderCode VARCHAR(8) "
			sParams = sParams & ",@vPictureOrderCode = N'" & PictureOrderCode & "'"

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.CompanyCode FROM C_Info AS A WITH(NOLOCK) WHERE OrderCode = @vPictureOrderCode) AS CPOC ON VWOC.CompanyCode = CPOC.CompanyCode AND VWOC.OrderType = '0'" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'�Ώۊ�Ƃ̋��l�[�ꗗ�p���R�[�h end
		'******************************************************************************

		'******************************************************************************
		'�o�^�� start
		'------------------------------------------------------------------------------
		If RegistDay <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vRegistDay VARCHAR(8) "
			sParams = sParams & ",@vRegistDay = N'" & RegistDay & "'"

			sJoin = sJoin & "INNER JOIN (SELECT A.OrderCode FROM C_Info AS A WITH(NOLOCK) WHERE RegistDay >= CONVERT(DATETIME, @vRegistDay)) AS CRD ON VWOC.OrderCode = CRD.OrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'�o�^�� end
		'******************************************************************************

		'******************************************************************************
		'�O��\�����̍ŐV���R�[�h start
		'------------------------------------------------------------------------------
		If BOC <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vBeforeOrderCode VARCHAR(8) "
			sParams = sParams & ",@vBeforeOrderCode = N'" & BOC & "'"

			If sWhere <> "" Then sWhere = sWhere & "AND "
			sWhere = sWhere & "VWOC.OrderCode > @vBeforeOrderCode" & vbCrLf
		End If
		'------------------------------------------------------------------------------
		'�O��\�����̍ŐV���R�[�h end
		'******************************************************************************

		'******************************************************************************
		'���R�[�hCSV start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If OrderCode <> "" Then
			aValue = Split(Replace(OrderCode, " ", ""), ",")
			For iParamNo = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOrderCode" & iParamNo & " CHAR(8)"
				sParams = sParams & ",@vOrderCode" & iParamNo & " = N'" & aValue(iParamNo) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vOrderCode" & iParamNo
			Next

			If sWhere <> "" Then sWhere = sWhere & "AND "
			If UBound(aValue) = 0 Then
				sWhere = sWhere & "VWOC.OrderCode = " & sTemp & vbCrLf
			Else
				sWhere = sWhere & "VWOC.OrderCode IN (" & sTemp & ")" & vbCrLf
			End If
		End If
		'------------------------------------------------------------------------------
		'���R�[�hCSV end
		'******************************************************************************

		'<�Г��Č��̑Ώۊ��>
		If LISOrderCompanyCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vLISOrderCompanyCode VARCHAR(8) "
			sParams = sParams & ",@vLISOrderCompanyCode = N'" & LISOrderCompanyCode & "'"

			If sWhere <> "" Then sWhere = sWhere & "AND "
			sWhere = sWhere & "VWOC.CompanyCode = @vLISOrderCompanyCode" & vbCrLf
		End If
		'</�Г��Č��̑Ώۊ��>

		If CStr(Top) <> "" Then Top = "TOP " & Top & vbCrLf
		sSQL = ""
		sSQL = sSQL & "SELECT " & Top & "VWOC.OrderCode "
		sSQL = sSQL & ",VWOC.SortNum "
		sSQL = sSQL & ",VWOC.RegistDay ,VWOC.UpdateDay" & vbCrLf
		sSQL = sSQL & "FROM vw_OrderCode_PlusOld AS VWOC WITH(NOLOCK)" & vbCrLf
		sSQL = sSQL & sJoin
		If sWhere <> "" Then sSQL = sSQL & "WHERE " & sWhere
		sSQL = sSQL & "ORDER BY VWOC.SortNum ASC, VWOC.UpdateDay DESC"

        If FeatureFlag <> "" Then
            sSQL = ""
            sSQL = sSQL & "SELECT  " & vbCrLf
            sSQL = sSQL & "VWOC.OrderCode  " & vbCrLf
            sSQL = sSQL & ",VWOC.SortNum  " & vbCrLf
            sSQL = sSQL & ",VWOC.RegistDay  " & vbCrLf
            sSQL = sSQL & ",VWOC.UpdateDay  " & vbCrLf
            sSQL = sSQL & "FROM vw_OrderCode_PlusOld AS VWOC WITH(NOLOCK)  " & vbCrLf
            sSQL = sSQL & "INNER JOIN ( " & vbCrLf
            sSQL = sSQL & "SELECT DISTINCT A.OrderCode  " & vbCrLf
            sSQL = sSQL & "FROM C_JobType AS A WITH(NOLOCK)  " & vbCrLf
            If FeatureFlag = "1" Then
                sSQL = sSQL & "WHERE (A.JobTypeCode IN ('1302000','1325000','1326000','1308000','1319000','1312000','1318000','1311000','1399000'))) AS CJT ON VWOC.OrderCode = CJT.OrderCode  " & vbCrLf
            End If
            sSQL = sSQL & "ORDER BY VWOC.SortNum ASC, VWOC.UpdateDay DESC " & vbCrLf
        End If


		GetSQLOrderSearchDetail = ""
		GetSQLOrderSearchDetail = GetSQLOrderSearchDetail & "/*�i�r�E���l�[�ڍ׌���*/" & vbCrLf
		GetSQLOrderSearchDetail = GetSQLOrderSearchDetail & "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
		GetSQLOrderSearchDetail = GetSQLOrderSearchDetail & "EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
		If sDeclare <> "" Then GetSQLOrderSearchDetail = GetSQLOrderSearchDetail & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams

		If sSearchCondition <> "" Then
			sSearchCondition = "<table class=""pattern1"" border=""0"" style=""width:600px;""><thead><tr><th colspan=""2"" style=""width:588px;"">��������</th></tr></thead><tbody>" & sSearchCondition & "</tbody></table>"
		Else
			sSearchCondition = "�Ȃ�"
		End If
'Response.Write GetSQLOrderSearchDetail
	End Function

	'******************************************************************************
	'�T�@�v�F���l�̃L�[���[�h�����k�n�f�������݂r�p�k���擾
	'���@���F
	'���@�l�F
	'���@���F2012/02/21 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Function GetSQLWriteLog()
		Dim sSQL,sSN,sKW,sSiteType

		sSN = Request.ServerVariables("SERVER_NAME")
		If InStr(sSN,"www.shigotonavi.co.jp") + InStr(sSN,"www-b1.shigotonavi.co.jp") > 0 Then
			sSiteType = "1"
		ElseIf InStr(sSN,"m.shigotonavi.jp") + InStr(sSN,"m-b1.shigotonavi.jp") > 0 Then
			sSiteType = "2"
		ElseIf InStr(sSN,"www.a-rirekisyo.jp") + InStr(sSN,"www-b1.a-rirekisyo.jp") > 0 Then
			sSiteType = "3"
		End If

		sKW = KW
		If sKW = "" Then sKW = Keyword

		If sKW > "" Then
			sSQL = sSQL & "EXEC up_RegLOG_SearchOrderKeyword '" & G_USERID & "'"
			sSQL = sSQL & ",'" & ChkSQLStr(Request.ServerVariables("REMOTE_ADDR")) & "'"
			sSQL = sSQL & ",'" & ChkSQLStr(Session.SessionID) & "'"
			sSQL = sSQL & ",'" & sSiteType & "'"
			sSQL = sSQL & ",'" & sKW & "';"
		End If

		GetSQLWriteLog = sSQL
	End Function

	'******************************************************************************
	'�T�@�v�F���l�[�ڍ׌��������o�͂g�s�l�k���擾
	'���@���F
	'���@�l�F
	'���@���F2007/04/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Function GetHtmlSearchCondition()
		Dim sTemp
		Dim sTemp2
		Dim idx

		If SearchDetailFlag = "" Then Exit Function

		GetHtmlSearchCondition = ""

		'�Г��O�Č������t���O
		sTemp = ""
		If OrderTypeFlag <> "" Then
			If OrderTypeFlag = "0" Then
				sTemp = "��ʋ��l���"
			ElseIf OrderTypeFlag = "1" Then
				sTemp = "���X�̏Љ�E�h�����"
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("��ʁ^���X�敪",sTemp)
		End If

		'�V���t���O
		sTemp = ""
		If NewFlag = "1" or NewKoukokuFlag = "1" Then
			sTemp = "�V�����"
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�V���敪",sTemp)
		End If

		'�E��
		sTemp2 = ""
		If JobTypeBigCode1 & JobTypeCode1 & JobTypeBigCode2 & JobTypeCode2 & JobTypeBigCode3 & JobTypeCode3 <> "" Then
			sTemp = ""
			If JobTypeBigCode1 & JobTypeCode1 <> "" Then
				sTemp = sTemp & JobTypeName1
				If sTemp = "" And JobTypeBigName1 <> "" Then sTemp = sTemp & JobTypeBigName1

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "�@"
				sTemp2 = sTemp2 & sTemp
			End If

			sTemp = ""
			If JobTypeBigCode2 & JobTypeCode2 <> "" Then
				sTemp = sTemp & JobTypeName2
				If sTemp = "" And JobTypeBigName2 <> "" Then sTemp = sTemp & JobTypeBigName2

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "�@"
				sTemp2 = sTemp2 & sTemp
			End If

			sTemp = ""
			If JobTypeBigCode3 & JobTypeCode3 <> "" Then
				sTemp = sTemp & JobTypeName3
				If sTemp = "" And JobTypeBigName3 <> "" Then sTemp = sTemp & JobTypeBigName3

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "�@"
				sTemp2 = sTemp2 & sTemp
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�E��",sTemp2)
		End If

		'�Ζ��n
		sTemp = ""
		If PrefectureCode & City & RailwayLineCode & RailwayLineCode <> "" Then
			'�G���A
			sTemp = sTemp & AreaName

			'�s���{��
			If PrefectureName <> "" Then
				sTemp = sTemp & "�@"
				sTemp = sTemp & PrefectureName
			End If

			'�s��S
			If City <> "" Then
				sTemp = sTemp & "�@"
				sTemp = sTemp & City
			End If

			'����
			If RailwayLineCode <> "" Then
				sTemp = sTemp & "�@"
				sTemp = sTemp & RailwayLineName
			End If

			'�w
			If StationCode <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "�@"
				sTemp = sTemp & StationName & "�w"
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�Ζ��n",sTemp)
		End If

		'�Ζ��`��
		sTemp = ""
		If WorkingTypeCode1 & WorkingTypeCode2 & WorkingTypeCode3 <> "" Then
			If WorkingTypeCode1 <> "" Then sTemp = sTemp & WorkingTypeName1
			If WorkingTypeCode2 <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "�@"
				sTemp = sTemp & WorkingTypeName2
			End If
			If WorkingTypeCode3 <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "�@"
				sTemp = sTemp & WorkingTypeName3
			End If
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�Ζ��`��",sTemp)
		End If

		'�Ǝ�
		sTemp = ""
		If IndustryTypeCode <> "" Then
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�Ǝ�",IndustryTypeName)
		End If

		'���^
		sTemp = ""
		If YearlyIncomeMin & YearlyIncomeMax & MonthlyIncomeMin & MonthlyIncomeMax & DailyIncomeMin & DailyIncomeMax & HourlyIncomeMin & HourlyIncomeMax & PercentagePayFlag <> "" Then
			If PercentagePayFlag = "1" Then
				sTemp = sTemp & "����������"
			ElseIf PercentagePayFlag = "0" Then
				sTemp = sTemp & "�������Ȃ�"
			End If
			If YearlyIncomeMin & YearlyIncomeMax <> "" Then
				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "�N���F" & YearlyIncomeMin & "�`" & YearlyIncomeMax
			End If
			If MonthlyIncomeMin & YearlyIncomeMax <> "" Then
				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "�����F" & MonthlyIncomeMin & "�`" & MonthlyIncomeMax
			End If
			If DailyIncomeMin & DailyIncomeMax <> "" Then
				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "�����F" & DailyIncomeMin & "�`" & DailyIncomeMax
			End If
			If HourlyIncomeMin & HourlyIncomeMax <> "" Then
				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "�����F" & HourlyIncomeMin & "�`" & HourlyIncomeMax
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("���^",sTemp)
		End If

		'����
		sTemp = ""
		If InexperiencedPersonFlag & UtilizeLanguageFlag & TempFlag & UITurnFlag & ManyHolidayFlag & FlexFlag & _
		NearStationFlag & NoSmokingFlag & NewlyBuiltFlag & LandmarkFlag & RenovationFlag & DesignersFlag & _
		CompanyCafeteriaFlag & ShortOvertimeFlag & MaternityFlag & DressFreeFlag & MammyFlag & FixedTimeFlag & _
		ShortTimeFlag & HandicappedFlag & RentAllFlag & RentPartFlag & MealsFlag & MealsAssistanceFlag & _
		TrainingCostFlag & EntrepreneurCostFlag & MoneyFlag & LandShopFlag & FindJobFestiveFlag & AppointmentFlag & SocietyInsuranceFlag <> "" Then
			If InexperiencedPersonFlag = "1" Then sTemp = sTemp & "�u���o���҂n�j�v"
			If UtilizeLanguageFlag = "1" Then sTemp = sTemp & "�u��w���������v"
			If TempFlag = "1" Then sTemp = sTemp & "�u�h���v"
			If UITurnFlag = "1" Then sTemp = sTemp & "�u�t�h�^�[�����}�v"
			If ManyHolidayFlag = "1" Then sTemp = sTemp & "�u�x���P�Q�O���ȏ�v"
			If FlexFlag = "1" Then sTemp = sTemp & "�u�t���b�N�X�v"
			If NearStationFlag = "1" Then sTemp = sTemp & "�u�w��(�k��5���ȓ�)�v"
			If NoSmokingFlag = "1" Then sTemp = sTemp & "�u�։��E�����v"
			If NewlyBuiltFlag = "1" Then sTemp = sTemp & "�u�V�z�r���E�I�t�B�X(5�N�ȓ�)�v"
			If LandmarkFlag = "1" Then sTemp = sTemp & "�u���w(15�K�ȏ�)�r���v"
			If RenovationFlag = "1" Then sTemp = sTemp & "�u���m�x�[�V�����r���E�I�t�B�X(5�N�ȓ�)�v"
			If DesignersFlag = "1" Then sTemp = sTemp & "�u�f�U�C�i�[�Y�r���E�I�t�B�X�v"
			If CompanyCafeteriaFlag = "1" Then sTemp = sTemp & "�u�Ј��H���v"
			If ShortOvertimeFlag = "1" Then sTemp = sTemp & "�u�c��10h/���ȓ��v"
			If MaternityFlag = "1" Then sTemp = sTemp & "�u�Y�x�E��x���т���v"
			If DressFreeFlag = "1" Then sTemp = sTemp & "�u�������R�v"
			If MammyFlag = "1" Then sTemp = sTemp & "�u�q��ă}�}���}�v"
			If FixedTimeFlag = "1" Then sTemp = sTemp & "�u18���܂łɑގЁv"
			If ShortTimeFlag = "1" Then sTemp = sTemp & "�u1��6���Ԉȓ��J���v"
			If HandicappedFlag = "1" Then sTemp = sTemp & "�u��Q�Ҋ��}�v"
			If RentAllFlag = "1" Then sTemp = sTemp & "�u�Z���p�S�z�⏕����v"
			If RentPartFlag = "1" Then sTemp = sTemp & "�u�Z���p�ꕔ�⏕����v"
			If MealsFlag = "1" Then sTemp = sTemp & "�u�H���E�d���t���Č��v"
			If MealsAssistanceFlag = "1" Then sTemp = sTemp & "�u�H���⏕���x����v"
			If TrainingCostFlag = "1" Then sTemp = sTemp & "�u���C������x����v"
			If EntrepreneurCostFlag = "1" Then sTemp = sTemp & "�u�N�Ƌ@�ޕ⏕���x����v"
			If MoneyFlag = "1" Then sTemp = sTemp & "�u�����q�E�ᗘ�q�⏕���x����v"
			If LandShopFlag = "1" Then sTemp = sTemp & "�u�y�n�E�X�ܓ��񋟐��x����v"
			If FindJobFestiveFlag = "1" Then sTemp = sTemp & "�u�A�E���j�������x����v"
			If AppointmentFlag = "1" Then sTemp = sTemp & "�u���Ј��o�p���x����v"
			If SocietyInsuranceFlag = "1" Then sTemp = sTemp & "�u�Еۊ����v"

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("����",sTemp)
		End If

		'�A�Ǝ���
		sTemp = ""
		If WorkStartHour & WorkStartMinute & WorkEndHour & WorkEndMinute <> "" Then
			If WorkStartHour & WorkStartMinute <> "" Then sTemp = sTemp & "�A�ƊJ�n���ԁF" & WorkStartHour & ":" & WorkStartMinute & "&nbsp;�ȍ~"
			If WorkEndHour & WorkEndMinute <> "" And sTemp <> "" Then sTemp = sTemp & ","
			If WorkEndHour & WorkEndMinute <> "" Then sTemp = sTemp & "�A�ƏI�����ԁF" & WorkEndHour & ":" & WorkEndMinute & "&nbsp;�ȑO"

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�A�Ǝ���",sTemp)
		End If

		'�T�x���
		sTemp = ""
		If WeeklyHolidayType <> "" Then
			sTemp = sTemp & WeeklyHolidayTypeName
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�T�x���",sTemp)
		End If

		'�N��
		'sTemp = ""
		'If Age <> "" Then
		'	sTemp = sTemp & Age & "�΂��܂�"
		'	GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�N��",sTemp)
		'End If

		'���ƔN
		sTemp = ""
		If SchoolTypeName & GraduateYear <> "" Then
			sTemp = SchoolTypeName & "�@" & GraduateYear & "�N��"
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�w��",sTemp)
		End If

		'�_�����
		sTemp = ""
		If AgreementTerm <> "" Then
			If AgreementTerm = "1" Then
				sTemp = "�`�P����"
			ElseIf AgreementTerm = "2" Then
				sTemp = "�`�Q����"
			ElseIf AgreementTerm = "3" Then
				sTemp = "�R�����ȏ�"
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�_�����",sTemp)
		End If

		'���i
		sTemp = ""
		If LicenseCount > 0 Then
			For idx = 0 To LicenseCount - 1
				If sTemp <> "" Then sTemp = sTemp & ","

				If LicenseName(idx) <> "" Then
					sTemp = sTemp & LicenseName(idx)
				ElseIf LicenseCategoryName(idx) <> "" Then
					sTemp = sTemp & LicenseCategoryName(idx)
				ElseIf LicenseGroupName(idx) <> "" Then
					sTemp = sTemp & LicenseGroupName(idx)
				End If
			Next

			If sTemp <> "" Then GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("���i",sTemp)
		End If

		'�n�r
		sTemp = ""
		If OSName <> "" Then
			sTemp = sTemp & OSName
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�n�r",sTemp)
		End If

		'�A�v���P�[�V����
		sTemp = ""
		If ApplicationName <> "" Then
			sTemp = sTemp & ApplicationName
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�A�v���P�[�V����",sTemp)
		End If

		'�J������
		sTemp = ""
		If DevelopmentLanguageName <> "" Then
			sTemp = sTemp & DevelopmentLanguageName
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�J������",sTemp)
		End If

		'�f�[�^�x�[�X
		sTemp = ""
		If DatabaseName <> "" Then
			sTemp = sTemp & DatabaseName
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�f�[�^�x�[�X",sTemp)
		End If

		'�L�[���[�h
		sTemp = ""
		If Keyword <> "" Then
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("�L�[���[�h",Keyword)
		End If

		'���R�[�h�i�����j
		If OrderCode <> "" Then
			GetHtmlSearchCondition = GetHtmlSearchCondition & GetHtmlSearchConditionTable("���R�[�h",OrderCode)
		End If

		If GetHtmlSearchCondition <> "" Then
			'GetHtmlSearchCondition = "<table class=""pattern1"" border=""0"" style=""width:600px;""><colgroup><col style=""width:138px;""><col style=""width:439px;""></colgroup><thead><tr><th colspan=""2"" style=""width:588px;"">��������</th></tr></thead><tbody>" & GetHtmlSearchCondition & "</tbody></table>"
			GetHtmlSearchCondition = "<div class=""description"">" & GetHtmlSearchCondition & "</div>"
		End If

	End Function

	Private Function GetHtmlSearchConditionTable(ByVal vKey, ByVal vValue)
		'GetHtmlSearchConditionTable = "<tr><th>" & vKey & "</th><td>" & vValue & "</td></tr>"
		GetHtmlSearchConditionTable = "�y"&vKey&"�z&nbsp;" & vValue & "<br>"
	End Function
End Class
%>
