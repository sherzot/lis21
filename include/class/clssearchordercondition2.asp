<%
'******************************************************************************
'�T�@�v�F����������ێ�����N���X
'�ց@���F��Public
'�@�@�@�FGetSearchParam				�F���d���ڍ׌����y�[�W�֓n��GET�p�����[�^�𐶐����Ď擾
'�@�@�@�FDspConditionHidden			�F���d���ڍ׌����̏���hidden���o�͂���
'�@�@�@�FGetSQLOrderSearchDetail	�F���l�[�ڍ׌����r�p�k���擾
'�@�@�@�FGetSQLWriteLog				�F���l�[�����k�n�f�������݂r�p�k���擾
'�@�@�@�FGetHtmlSearchCondition		�F���l�[�ڍ׌��������o�͂g�s�l�k���擾
'�@�@�@�F
'�@�@�@�F��Private
'�@�@�@�FClass_Initialize			�F�R���X�g���N�^
'�@�@�@�FSetNames					�F�R�[�h�ɑΉ��������̂������o�ϐ��ɐݒ�
'�@�@�@�FChkSQLType					�F�J���^���������ڍ׌������𔻒f����flgEasySearch��ݒ�
'�@�@�@�FChkData					�F�����o�ϐ��̐��������`�F�b�N���Ē���
'�@�@�@�F
'���@�l�F������ �ڍ׌����p�p�����[�^ �i�A�h�z�b�N�Ȃr�p�k�����j
'�@�@�@�Fsjtbig1�F��]�E��啪�ނP
'�@�@�@�Fsjt1	�F��]�E��P
'�@�@�@�Fsjtbig2�F��]�E��啪�ނQ
'�@�@�@�Fsjt2	�F��]�E��Q
'�@�@�@�Fsrc1	�F��]�����P
'�@�@�@�Fsrc2	�F��]�����Q
'�@�@�@�Fssc1	�F��]�w�P
'�@�@�@�Fssc2	�F��]�w�Q
'�@�@�@�Fsac1	�F��]�G���A�P
'�@�@�@�Fspc1	�F��]�s���{���P
'�@�@�@�Fsct1	�F��]�s��S�P
'�@�@�@�Fsac2	�F��]�G���A�Q
'�@�@�@�Fspc2	�F��]�s���{���Q
'�@�@�@�Fsct2	�F��]�s��S�Q
'�@�@�@�Fswt1	�F��]�Ζ��`�ԂP
'�@�@�@�Fswt2	�F��]�Ζ��`�ԂQ
'�@�@�@�Fswt3	�F��]�Ζ��`�ԂR
'�@�@�@�Fsit	�F��]�Ǝ�(�J���}��؂� [XX,XX,XX])
'�@�@�@�Fsppf	�F������
'�@�@�@�Fsyi	�F�N��
'�@�@�@�Fsmi	�F����
'�@�@�@�Fsdi	�F����
'�@�@�@�Fshi	�F����
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
'�@�@�@�Fsoc	�F���R�[�h�i�����j
'�@�@�@�F
'�@�@�@�F������ �J���^�������p�p�����[�^ (�X�g�A�h up_SearchOrder ���p)
'�@�@�@�Fjt		�F�E��啪�ރR�[�h
'�@�@�@�Fjt2	�F�E��R�[�h
'�@�@�@�Fac		�F�G���A�R�[�h
'�@�@�@�Fac2	�F�s���{���R�[�h
'�@�@�@�Fwt		�F�Ζ��`�ԃR�[�h
'�@�@�@�Fkw		�F�L�[���[�h
'�@�@�@�F
'�@�@�@�F������ ���c�[���p
'�@�@�@�Fboc	�F�O��\�����R�[�h
'�@�@�@�F
'�@�@�@�F������ �ڍ׌����p�o�n�r�s�f�[�^ �i�A�h�z�b�N�Ȃr�p�k�����j
'�@�@�@�FCONF_SearchHopeJobTypeBigCode1			�F��]�E��啪�ނP
'�@�@�@�FCONF_SearchHopeJobTypeCode1			�F��]�E��P
'�@�@�@�FCONF_SearchHopeJobTypeBigCode2			�F��]�E��啪�ނQ
'�@�@�@�FCONF_SearchHopeJobTypeCode2			�F��]�E��Q
'�@�@�@�FCONF_SearchRailwayLineCode1			�F��]�����P
'�@�@�@�FCONF_SearchRailwayLineCode2			�F��]�����Q
'�@�@�@�FCONF_SearchStationCode1				�F��]�w�P
'�@�@�@�FCONF_SearchStationCode2				�F��]�w�Q
'�@�@�@�FCONF_SearchAreaCode1					�F��]�G���A�P
'�@�@�@�FCONF_SearchPrefectureCode1				�F��]�s���{���P
'�@�@�@�FCONF_SearchCity1						�F��]�s��S�P
'�@�@�@�FCONF_SearchAreaCode2					�F��]�G���A�Q
'�@�@�@�FCONF_SearchPrefectureCode2				�F��]�s���{���Q
'�@�@�@�FCONF_SearchCity2						�F��]�s��S�Q
'�@�@�@�FCONF_SearchHopeWorkingTypeCode1		�F��]�Ζ��`�ԂP
'�@�@�@�FCONF_SearchHopeWorkingTypeCode2		�F��]�Ζ��`�ԂQ
'�@�@�@�FCONF_SearchHopeWorkingTypeCode3		�F��]�Ζ��`�ԂR
'�@�@�@�FCONF_SearchHopeIndustryTypeCode		�F��]�Ǝ�(�J���}��؂� [XX,XX,XX])
'�@�@�@�FCONF_SearchPercentagePayFlag			�F������
'�@�@�@�FCONF_SearchYearlyIncome				�F�N��
'�@�@�@�FCONF_SearchMonthlyIncome				�F����
'�@�@�@�FCONF_SearchDailyIncome					�F����
'�@�@�@�FCONF_SearchHourlyIncome				�F����
'�@�@�@�FCONF_SearchWorkStartHour				�F�A�ƊJ�n���ԁi���j
'�@�@�@�FCONF_SearchWorkStartMinute				�F�A�ƊJ�n���ԁi���j
'�@�@�@�FCONF_SearchWorkEndHour					�F�A�ƏI�����ԁi���j
'�@�@�@�FCONF_SearchWorkEndMinute				�F�A�ƏI�����ԁi���j
'�@�@�@�FCONF_SearchWeeklyHolidayType			�F�T�x���
'�@�@�@�FCONF_SearchAge							�F�N��
'�@�@�@�FCONF_SearchAgreementTerm				�F�_�����
'�@�@�@�FCONF_SearchLicenseGroupCode1			�F���i�啪��
'�@�@�@�FCONF_SearchLicenseCategoryCode1		�F���i������
'�@�@�@�FCONF_SearchLicenseCode1				�F���i������
'�@�@�@�FCONF_SearchOSCode1						�F�n�r
'�@�@�@�FCONF_SearchApplicationCode1			�F�A�v���P�[�V����
'�@�@�@�FCONF_SearchDevelopmentLanguageCode1	�F�J������
'�@�@�@�FCONF_SearchDatabaseCode1				�F�f�[�^�x�[�X
'�@�@�@�FCONF_SearchKeyword						�F�������[�h
'�@�@�@�FCONF_SearchKeywordFlag					�F�������[�h�t���O [1]OR [2]AND
'�@�@�@�FCONF_SearchOrderCode					�F���R�[�h�i�����j
'�@�@�@�FCONF_SearchInexperiencedPersonFlag		�F�����i���o�����}�j
'�@�@�@�FCONF_SearchUtilizeLanguageFlag			�F�����i��w���������j
'�@�@�@�FCONF_SearchTempFlag					�F�����i�h���j�����ݖ��g�p
'�@�@�@�FCONF_SearchUITurnFlag					�F�����i�t�h�^�[���j
'�@�@�@�FCONF_SearchManyHolidayFlag				�F�����i�x���P�Q�O���ȏ�j
'�@�@�@�FCONF_SearchFlexFlag					�F�����i�t���b�N�X�^�C���j
'�@�@�@�FCONF_SP								�F���W�R�[�h�i�����ł͎g��Ȃ��B�p�����[�^�����p�ɕێ�����B�j
'�@�@�@�F
'�@�@�@�F������ �J���^�������p�o�n�r�s�f�[�^ (�X�g�A�h up_SearchOrder ���p)
'�@�@�@�FCONF_JT	�F�E��啪�ރR�[�h
'�@�@�@�FCONF_JT2	�F�E��R�[�h
'�@�@�@�FCONF_AC	�F�G���A�R�[�h
'�@�@�@�FCONF_AC2	�F�s���{���R�[�h
'�@�@�@�FCONF_WT	�F�Ζ��`�ԃR�[�h
'�@�@�@�FCONF_ST1	�F�����i���o�����}�j
'�@�@�@�FCONF_ST2	�F�����i��w���������j
'�@�@�@�FCONF_ST3	�F�����i�h���j�����ݖ��g�p
'�@�@�@�FCONF_ST4	�F�����i�t�h�^�[���j
'�@�@�@�FCONF_ST5	�F�����i�x���P�Q�O���ȏ�j
'�@�@�@�FCONF_ST6	�F�����i�t���b�N�X�^�C���j
'�@�@�@�FCONF_KW	�F�L�[���[�h
'�@�@�@�F
'�@�@�@�F������ �g�p���@
'�@�@�@�FDim oSOC
'�@�@�@�FDim sSQL
'�@�@�@�FSet oSOC = New clsSearchOrderCondition	'�������ꂽ���_�Ńp�����[�^�Ƃo�n�r�s�f�[�^����r�p�k����������Ă���
'�@�@�@�FoSOC.Top = 100	'SELECT��ŏ����ݒ�
'�@�@�@�FsSQL = oSOC.GetSQLOrderSearchDetail()	'�r�p�k���擾
'�@�@�@�F
'�X�@�V�F2007/04/05 LIS K.Kokubo �쐬
'�@�@�@�F2007/10/10 LIS K.Kokubo ���c�[���p�ϐ��ǉ�
'�@�@�@�F2007/10/31 LIS K.Kokubo TOP ??? �p�ϐ��ǉ�
'�@�@�@�F2008/01/15 LIS K.Kokubo �p�����[�^���N�G����
'******************************************************************************
Class clsSearchOrderCondition
	'�������������o�ϐ�
	Public Top						'SELECT�Ŏ擾���錏�� (SELECT TOP �� * FROM �`)
	Public JobTypeBigCode1			'��]�E��啪�ނP
	Public JobTypeCode1				'��]�E��P
	Public JobTypeBigCode2			'��]�E��啪�ނQ
	Public JobTypeCode2				'��]�E��Q
	Public RailwayLineCode1			'��]�����P
	Public RailwayLineCode2			'��]�����Q
	Public StationCode1				'��]�w�P
	Public StationCode2				'��]�w�Q
	Public AreaCode1				'��]�G���A�P
	Public PrefectureCode1			'��]�s���{���P
	Public City1					'��]�s��S�P
	Public AreaCode2				'��]�G���A�Q
	Public PrefectureCode2			'��]�s���{���Q
	Public City2					'��]�s��S�Q
	Public WorkingTypeCode1			'��]�Ζ��`�ԂP
	Public WorkingTypeCode2			'��]�Ζ��`�ԂQ
	Public WorkingTypeCode3			'��]�Ζ��`�ԂR
	Public IndustryTypeCode			'��]�Ǝ�(�J���}��؂� [XX,XX,XX])
	Public IndustryTypeCode1		'��]�Ǝ�P
	Public IndustryTypeCode2		'��]�Ǝ�Q
	Public IndustryTypeCode3		'��]�Ǝ�R
	Public PercentagePayFlag		'������
	Public YearlyIncome				'�N��
	Public MonthlyIncome			'����
	Public DailyIncome				'����
	Public HourlyIncome				'����
	Public WorkStartHour			'�A�ƊJ�n���ԁi���j
	Public WorkStartMinute			'�A�ƊJ�n���ԁi���j
	Public WorkEndHour				'�A�ƏI�����ԁi���j
	Public WorkEndMinute			'�A�ƏI�����ԁi���j
	Public WeeklyHolidayType		'�T�x���
	Public Age						'�N��
	Public AgreementTerm			'�_�����
	Public LicenseGroupCode1		'���i�啪��
	Public LicenseCategoryCode1		'���i������
	Public LicenseCode1				'���i������
	Public OSCode1					'�n�r
	Public OACode1
	Public ApplicationCode1			'�A�v���P�[�V����
	Public DevelopmentLanguageCode1	'�J������
	Public DatabaseCode1			'�f�[�^�x�[�X
	Public Keyword					'�������[�h
	Public KeywordFlag				'�������[�h�t���O [1]OR [2]AND
	Public OrderCode				'���R�[�h�i�����j
	Public Specialty
	Public InexperiencedPersonFlag	'�����i�j
	Public UtilizeLanguageFlag		'�����i�j
	Public TempFlag					'�����i�h���j
	Public UITurnFlag				'�����i�t�h�^�[�����}�j
	Public ManyHolidayFlag			'�����i�x���P�Q�O���ȏ�j
	Public FlexFlag					'�����i�t���b�N�X�j

	'�J���^����������
	Public JT	'�E��啪�ރR�[�h
	Public JT2	'�E��R�[�h
	Public AC	'�G���A�R�[�h
	Public AC2	'�s���{���R�[�h
	Public WT	'�Ζ��`�ԃR�[�h
	Public ST
	Public ST1	'����
	Public ST2	'����
	Public ST3	'����
	Public ST4	'����
	Public ST5	'����
	Public ST6	'����
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
	Public RailwayLineName1	'��]�������̂P
	Public RailwayLineName2	'��]�������̂Q
	Public StationName1
	Public StationName2
	Public AreaName1
	Public AreaName2
	Public PrefectureName1
	Public PrefectureName2
	Public WorkingTypeName1
	Public WorkingTypeName2
	Public WorkingTypeName3
	Public IndustryTypeName1
	Public IndustryTypeName2
	Public IndustryTypeName3
	Public WeeklyHolidayTypeName
	Public OSName1
	Public ApplicationName1
	Public DevelopmentLanguageName1
	Public DatabaseName1
	Public LicenseGroupName1	'���i�啪�ޖ��̂P
	Public LicenseCategoryName1	'���i�����ޖ��̂P
	Public LicenseName1		'���i���̂P

	'���̑������o�ϐ�
	Public flgEasySearch	'�J���^�������t���O [True]�J���^������ [False]�ڍ׌���
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
		'FORM�f�[�^���猟���������擾
		JobTypeBigCode1 = GetForm("CONF_SearchHopeJobTypeBigCode1", 1)
		JobTypeCode1 = GetForm("CONF_SearchHopeJobTypeCode1", 1)
		JobTypeBigCode2 = GetForm("CONF_SearchHopeJobTypeBigCode2", 1)
		JobTypeCode2 = GetForm("CONF_SearchHopeJobTypeCode2", 1)
		RailwayLineCode1 = GetForm("CONF_SearchRailwayLineCode1", 1)
		RailwayLineCode2 = GetForm("CONF_SearchRailwayLineCode2", 1)
		StationCode1 = GetForm("CONF_SearchStationCode1", 1)
		StationCode2 = GetForm("CONF_SearchStationCode2", 1)
		AreaCode1 = GetForm("CONF_SearchAreaCode1", 1)
		PrefectureCode1 = GetForm("CONF_SearchPrefectureCode1", 1)
		City1 = GetForm("CONF_SearchCity1", 1)
		AreaCode2 = GetForm("CONF_SearchAreaCode2", 1)
		PrefectureCode2 = GetForm("CONF_SearchPrefectureCode2", 1)
		City2 = GetForm("CONF_SearchCity2", 1)
		WorkingTypeCode1 = GetForm("CONF_SearchHopeWorkingTypeCode1", 1)
		WorkingTypeCode2 = GetForm("CONF_SearchHopeWorkingTypeCode2", 1)
		WorkingTypeCode3 = GetForm("CONF_SearchHopeWorkingTypeCode3", 1)
		IndustryTypeCode = GetForm("CONF_SearchHopeIndustryTypeCode", 1)
		PercentagePayFlag = GetForm("CONF_SearchPercentagePayFlag", 1)
		YearlyIncome = GetForm("CONF_SearchYearlyIncome", 1)
		MonthlyIncome = GetForm("CONF_SearchMonthlyIncome", 1)
		DailyIncome = GetForm("CONF_SearchDailyIncome", 1)
		HourlyIncome = GetForm("CONF_SearchHourlyIncome", 1)
		WorkStartHour = GetForm("CONF_SearchWorkStartHour", 1)
		WorkStartMinute = GetForm("CONF_SearchWorkStartMinute", 1)
		WorkEndHour = GetForm("CONF_SearchWorkEndHour", 1)
		WorkEndMinute = GetForm("CONF_SearchWorkEndMinute", 1)
		WeeklyHolidayType = GetForm("CONF_SearchWeeklyHolidayType", 1)
		Age = GetForm("CONF_SearchAge", 1)
		AgreementTerm = GetForm("CONF_SearchAgreementTerm", 1)
		LicenseGroupCode1 = GetForm("CONF_SearchLicenseGroupCode1", 1)
		LicenseCategoryCode1 = GetForm("CONF_SearchLicenseCategoryCode1", 1)
		LicenseCode1 = GetForm("CONF_SearchLicenseCode1", 1)
		OSCode1 = GetForm("CONF_SearchOSCode1", 1)
		ApplicationCode1 = GetForm("CONF_SearchApplicationCode1", 1)
		DevelopmentLanguageCode1 = GetForm("CONF_SearchDevelopmentLanguageCode1", 1)
		DatabaseCode1 = GetForm("CONF_SearchDatabaseCode1", 1)
		Keyword = GetForm("CONF_SearchKeyword", 1)
		KeywordFlag = GetForm("CONF_SearchKeywordFlag", 1)
		OrderCode = GetForm("CONF_SearchOrderCode", 1)
		InexperiencedPersonFlag = GetForm("CONF_SearchInexperiencedPersonFlag", 1)
		UtilizeLanguageFlag = GetForm("CONF_SearchUtilizeLanguageFlag", 1)
		TempFlag = GetForm("CONF_SearchTempFlag", 1)
		UITurnFlag = GetForm("CONF_SearchUITurnFlag", 1)
		ManyHolidayFlag = GetForm("CONF_SearchManyHolidayFlag", 1)
		FlexFlag = GetForm("CONF_SearchFlexFlag", 1)
		SP = GetForm("CONF_SP", 1)

		'�p�����[�^���猟���������擾
		If GetForm("sjtbig1", 2) <> "" Then JobTypeBigCode1 = GetForm("sjtbig1", 2)
		If GetForm("sjt1", 2) <> "" Then JobTypeCode1 = GetForm("sjt1", 2)
		If GetForm("sjtbig2", 2) <> "" Then JobTypeBigCode2 = GetForm("sjtbig2", 2)
		If GetForm("sjt2", 2) <> "" Then JobTypeCode2 = GetForm("sjt2", 2)
		If GetForm("src1", 2) <> "" Then RailwayLineCode1 = GetForm("src1", 2)
		If GetForm("src2", 2) <> "" Then RailwayLineCode2 = GetForm("src2", 2)
		If GetForm("ssc1", 2) <> "" Then StationCode1 = GetForm("ssc1", 2)
		If GetForm("ssc2", 2) <> "" Then StationCode2 = GetForm("ssc2", 2)
		If GetForm("sac1", 2) <> "" Then AreaCode1 = GetForm("sac1", 2)
		If GetForm("spc1", 2) <> "" Then PrefectureCode1 = GetForm("spc1", 2)
		If GetForm("sct1", 2) <> "" Then City1 = GetForm("sct1", 2)
		If GetForm("sac2", 2) <> "" Then AreaCode2 = GetForm("sac2", 2)
		If GetForm("spc2", 2) <> "" Then PrefectureCode2 = GetForm("spc2", 2)
		If GetForm("sct2", 2) <> "" Then City2 = GetForm("sct2", 2)
		If GetForm("swt1", 2) <> "" Then WorkingTypeCode1 = GetForm("swt1", 2)
		If GetForm("swt2", 2) <> "" Then WorkingTypeCode2 = GetForm("swt2", 2)
		If GetForm("swt3", 2) <> "" Then WorkingTypeCode3 = GetForm("swt3", 2)
		If GetForm("sit", 2) <> "" Then IndustryTypeCode = GetForm("sit", 2)
		If GetForm("sppf", 2) <> "" Then PercentagePayFlag = GetForm("sppf", 2)
		If GetForm("syi", 2) <> "" Then YearlyIncome = GetForm("syi", 2)
		If GetForm("smi", 2) <> "" Then MonthlyIncome = GetForm("smi", 2)
		If GetForm("sdi", 2) <> "" Then DailyIncome = GetForm("sdi", 2)
		If GetForm("shi", 2) <> "" Then HourlyIncome = GetForm("shi", 2)
		If GetForm("swsh", 2) <> "" Then WorkStartHour = GetForm("swsh", 2)
		If GetForm("swsm", 2) <> "" Then WorkStartMinute = GetForm("swsm", 2)
		If GetForm("sweh", 2) <> "" Then WorkEndHour = GetForm("sweh", 2)
		If GetForm("swem", 2) <> "" Then WorkEndMinute = GetForm("swem", 2)
		If GetForm("swht", 2) <> "" Then WeeklyHolidayType = GetForm("swht", 2)
		If GetForm("sage", 2) <> "" Then Age = GetForm("sage", 2)
		If GetForm("sat", 2) <> "" Then AgreementTerm = GetForm("sat", 2)
		If GetForm("slg1", 2) <> "" Then LicenseGroupCode1 = GetForm("slg1", 2)
		If GetForm("slc1", 2) <> "" Then LicenseCategoryCode1 = GetForm("slc1", 2)
		If GetForm("sl1", 2) <> "" Then LicenseCode1 = GetForm("sl1", 2)
		If GetForm("sos1", 2) <> "" Then OSCode1 = GetForm("sos1", 2)
		If GetForm("sap1", 2) <> "" Then ApplicationCode1 = GetForm("sap1", 2)
		If GetForm("sdl1", 2) <> "" Then DevelopmentLanguageCode1 = GetForm("sdl1", 2)
		If GetForm("sdb1", 2) <> "" Then DatabaseCode1 = GetForm("sdb1", 2)
		If GetForm("skw", 2) <> "" Then Keyword = GetForm("skw", 2)
		If GetForm("skwflag", 2) <> "" Then KeywordFlag = GetForm("skwflag", 2)
		If GetForm("sst", 2) <> "" Then Specialty = GetForm("sst", 2)
		If GetForm("soc", 2) <> "" Then OrderCode = GetForm("soc", 2)

		If IsRE(GetForm("sst", 2), "^[01][01][01][01][01][01]$", True) = True Then
			If Mid(GetForm("sst", 2), 1, 1) = "1" Then InexperiencedPersonFlag = "1"
			If Mid(GetForm("sst", 2), 2, 1) = "1" Then UtilizeLanguageFlag = "1"
			If Mid(GetForm("sst", 2), 3, 1) = "1" Then TempFlag = "1"
			If Mid(GetForm("sst", 2), 4, 1) = "1" Then UITurnFlag = "1"
			If Mid(GetForm("sst", 2), 5, 1) = "1" Then ManyHolidayFlag = "1"
			If Mid(GetForm("sst", 2), 6, 1) = "1" Then FlexFlag = "1"
		End If

		'�����r�b�g������
		Specialty = ""
		If InexperiencedPersonFlag & UtilizeLanguageFlag & TempFlag & UITurnFlag & ManyHolidayFlag & FlexFlag <> "" Then
			If InexperiencedPersonFlag <> "" Then: Specialty = Specialty & InexperiencedPersonFlag: Else: Specialty = Specialty & "0": End If
			If UtilizeLanguageFlag <> "" Then: Specialty = Specialty & UtilizeLanguageFlag: Else: Specialty = Specialty & "0": End If
			If TempFlag <> "" Then: Specialty = Specialty & TempFlag: Else: Specialty = Specialty & "0": End If
			If UITurnFlag <> "" Then: Specialty = Specialty & UITurnFlag: Else: Specialty = Specialty & "0": End If
			If ManyHolidayFlag <> "" Then: Specialty = Specialty & ManyHolidayFlag: Else: Specialty = Specialty & "0": End If
			If FlexFlag <> "" Then: Specialty = Specialty & FlexFlag: Else: Specialty = Specialty & "0": End If
		End If

		'��]�Ǝ�
		If IndustryTypeCode <> "" Then
			IndustryTypeCode = Replace(IndustryTypeCode, " ", "")
			Dim aHITC
			Dim idx

			aHITC = Split(IndustryTypeCode, ",")
			For idx = 0 To UBound(aHITC)
				Select Case idx
					Case 0:	IndustryTypeCode1 = aHITC(idx)
					Case 1:	IndustryTypeCode2 = aHITC(idx)
					Case 2:	IndustryTypeCode3 = aHITC(idx)
				End Select
			Next
		End If

		'�J���^�����������擾�iFORM�f�[�^�j
		JT = GetForm("CONF_JT", 1)
		JT2 = GetForm("CONF_JT2", 1)
		AC = GetForm("CONF_AC", 1)
		AC2 = GetForm("CONF_AC2", 1)
		WT = GetForm("CONF_WT", 1)
		ST1 = GetForm("CONF_ST1", 1)
		ST2 = GetForm("CONF_ST2", 1)
		ST3 = GetForm("CONF_ST3", 1)
		ST4 = GetForm("CONF_ST4", 1)
		ST5 = GetForm("CONF_ST5", 1)
		ST6 = GetForm("CONF_ST6", 1)
		KW = GetForm("CONF_KW", 1)

		'�����r�b�g������
		ST = ""
		If ST1 & ST2 & ST3 & ST4 & ST5 & ST6 <> "" Then
			ST = ""
			If ST1 <> "" Then: ST = ST & ST1: Else: ST = ST & "0": End If
			If ST2 <> "" Then: ST = ST & ST2: Else: ST = ST & "0": End If
			If ST3 <> "" Then: ST = ST & ST3: Else: ST = ST & "0": End If
			If ST4 <> "" Then: ST = ST & ST4: Else: ST = ST & "0": End If
			If ST5 <> "" Then: ST = ST & ST5: Else: ST = ST & "0": End If
			If ST6 <> "" Then: ST = ST & ST6: Else: ST = ST & "0": End If

			Specialty = ST
		End If

		'�s�n�o����
		POC = GetForm("poc", 2)

		If POC <> "" Then OrderCode = POC

		'���W
		If GetForm("sp", 2) <> "" Then SP = GetForm("sp", 2)

		'�J���^�����������擾�i�p�����[�^�j
		If GetForm("jt", 2) <> "" Then JT = GetForm("jt", 2)
		If GetForm("jt2", 2) <> "" Then JT2 = GetForm("jt2", 2)
		If GetForm("ac", 2) <> "" Then AC = GetForm("ac", 2)
		If GetForm("ac2", 2) <> "" Then AC2 = GetForm("ac2", 2)
		If GetForm("wt", 2) <> "" Then WT = GetForm("wt", 2)
		If GetForm("kw", 2) <> "" Then KW = GetForm("kw", 2)
		If IsRE(GetForm("st", 2), "^[01][01][01][01][01][01]$", True) = True Then ST1 = Mid(GetForm("st", 2), 1, 1)
		If IsRE(GetForm("st", 2), "^[01][01][01][01][01][01]$", True) = True Then ST2 = Mid(GetForm("st", 2), 2, 1)
		If IsRE(GetForm("st", 2), "^[01][01][01][01][01][01]$", True) = True Then ST3 = Mid(GetForm("st", 2), 3, 1)
		If IsRE(GetForm("st", 2), "^[01][01][01][01][01][01]$", True) = True Then ST4 = Mid(GetForm("st", 2), 4, 1)
		If IsRE(GetForm("st", 2), "^[01][01][01][01][01][01]$", True) = True Then ST5 = Mid(GetForm("st", 2), 5, 1)
		If IsRE(GetForm("st", 2), "^[01][01][01][01][01][01]$", True) = True Then ST6 = Mid(GetForm("st", 2), 6, 1)

		If JT <> "" Then JobTypeCode1 = JT
		If JT2 <> "" Then JobTypeCode1 = JT2
		If AC <> "" Then AreaCode1 = AC
		If AC2 <> "" Then PrefectureCode1 = AC2
		If WT <> "" Then WorkingTypeCode1 = WT
		If KW <> "" Then Keyword = KW

		'���������iFORM�f�[�^�j
		PC = GetForm("CONF_PC", 1)
		RC = GetForm("CONF_RC", 1)
		SC = GetForm("CONF_SC", 1)

		'���������i�p�����[�^�j
		If GetForm("pc", 2) <> "" Then PC = GetForm("pc", 2)
		If GetForm("rc", 2) <> "" Then RC = GetForm("rc", 2)
		If GetForm("sc", 2) <> "" Then SC = GetForm("sc", 2)

		If PC <> "" Then PrefectureCode1 = PC
		If RC <> "" Then RailwayLineCode1 = RC
		If SC <> "" Then StationCode1 = SC

		'���c�[��
		BOC = GetForm("boc", 2)

		'**********************************************************************
		'�␳ start
		'----------------------------------------------------------------------
		If AC = "" And AC2 <> "" Then
			sSQL = "up_GetListPrefecture '', '" & AC2 & "', ''"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then AC = oRS.Collect("AreaCode")
			Call RSClose(oRS)
		End If
		'----------------------------------------------------------------------
		'�␳ end
		'**********************************************************************

		'�f�[�^�������`�F�b�N
		Call ChkData()

		'�J���^�������E�ڍ׌�������
		Call ChkSQLType()

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
	'�T�@�v�F�R�[�h�ɑΉ��������̂��擾����
	'�쐬�ҁFLis K.Kokubo
	'�쐬���F2007/04/04 Lis K.Kokubo
	'�X�@�V�F
	'���@�l�F
	'******************************************************************************
	Private Sub SetNames()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sError

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
		End If

		'��]�����P
		If RailwayLineCode1 <> "" Then
			sSQL = "up_GetRailwayLineName '" & RailwayLineCode1 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				RailwayLineName1 = ChkStr(oRS.Collect("RailwayLineName"))
			End If
			Call RSClose(oRS)
		End If
		'��]�����Q
		If RailwayLineCode2 <> "" Then
			sSQL = "up_GetRailwayLineName '" & RailwayLineCode2 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				RailwayLineName2 = ChkStr(oRS.Collect("RailwayLineName"))
			End If
			Call RSClose(oRS)
		End If

		'��]�w�P
		If StationCode1 <> "" Then
			sSQL = "up_GetStationName '" & StationCode1 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				StationName1 = ChkStr(oRS.Collect("StationName"))
			End If
			Call RSClose(oRS)
		End If
		'��]�w�Q
		If StationCode2 <> "" Then
			sSQL = "up_GetStationName '" & StationCode2 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				StationName2 = ChkStr(oRS.Collect("StationName"))
			End If
			Call RSClose(oRS)
		End If

		'�G���A�P
		If AreaCode1 <> "" Then
			AreaName1 = GetDetail("Area", AreaCode1)
		End If

		'�G���A�Q
		If AreaCode2 <> "" Then
			AreaName2 = GetDetail("Area", AreaCode2)
		End If

		'�s���{���P
		If PrefectureCode1 <> "" Then
			PrefectureName1 = GetDetail("Prefecture", PrefectureCode1)

			If AreaCode1 = "" Then
				sSQL = "SELECT A.AreaCode, B.AreaName FROM Area AS A WITH(NOLOCK) INNER JOIN vw_Area AS B ON A.AreaCode = B.AreaCode WHERE A.PrefectureCode = '" & PrefectureCode1 & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					AreaCode1 = ChkStr(oRS.Collect("AreaCode"))
					AreaName1 = ChkStr(oRS.Collect("AreaName"))
				End If
				Call RSClose(oRS)
			End If
		End If

		'�s���{���Q
		If PrefectureCode2 <> "" Then
			PrefectureName2 = GetDetail("Prefecture", PrefectureCode2)

			If AreaCode2 = "" Then
				sSQL = "SELECT A.AreaCode, B.AreaName FROM Area AS A WITH(NOLOCK) INNER JOIN vw_Area AS B ON A.AreaCode = B.AreaCode WHERE A.PrefectureCode = '" & PrefectureCode2 & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					AreaCode2 = ChkStr(oRS.Collect("AreaCode"))
					AreaName2 = ChkStr(oRS.Collect("AreaName"))
				End If
				Call RSClose(oRS)
			End If
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

		'�Ǝ�P
		If IndustryTypeCode1 <> "" Then
			IndustryTypeName1 = GetDetail("IndustryType", IndustryTypeCode1)
		End If

		'�Ǝ�Q
		If IndustryTypeCode2 <> "" Then
			IndustryTypeName2 = GetDetail("IndustryType", IndustryTypeCode2)
		End If

		'�Ǝ�R
		If IndustryTypeCode3 <> "" Then
			IndustryTypeName3 = GetDetail("IndustryType", IndustryTypeCode3)
		End If

		'�T�x���
		If WeeklyHolidayType <> "" Then
			WeeklyHolidayTypeName = GetDetail("WeeklyHolidayType", WeeklyHolidayType)
		End If

		'�n�r
		If OSCode1 <> "" Then
			OSName1 = GetDetail("OS", OSCode1)
		End If

		'�A�v���P�[�V����
		If ApplicationCode1 <> "" Then
			ApplicationName1 = GetDetail("Application", ApplicationCode1)
		End If

		'�J������
		If DevelopmentLanguageCode1 <> "" Then
			DevelopmentLanguageName1 = GetDetail("DevelopmentLanguage", DevelopmentLanguageCode1)
		End If

		'�f�[�^�x�[�X
		If DatabaseCode1 <> "" Then
			DatabaseName1 = GetDetail("Database", DatabaseCode1)
		End If

		'���i
		If IsRE(LicenseGroupCode1, "^\d\d$", True) = True Then
			'�啪��
			sSQL = "sp_GetListLicenseGroup '" & LicenseGroupCode1 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				LicenseGroupName1 = ChkStr(oRS.Collect("GroupName"))
			End If
			Call RSClose(oRS)

			'������
			If IsRE(LicenseCategoryCode1, "^\d\d\d$", True) = True Then
				sSQL = "sp_GetListLicenseCategory '" & LicenseGroupCode1 & "', '" & LicenseCategoryCode1 & "'"
				flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
				If GetRSState(oRS) = True Then
					LicenseCategoryName1 = ChkStr(oRS.Collect("CategoryName"))
				End If
				Call RSClose(oRS)

				'�啪��
				If IsRE(LicenseCode1, "^\d\d$", True) = True Then
					sSQL = "sp_GetListLicense '" & LicenseGroupCode1 & "', '" & LicenseCategoryCode1 & "', '" & LicenseCode1 & "'"
					flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
					If GetRSState(oRS) = True Then
						LicenseName1 = ChkStr(oRS.Collect("Name"))
					End If
					Call RSClose(oRS)
				End If
			End If
		End If
	End Sub

	'******************************************************************************
	'�T�@�v�F�J���^���������ڍ׌������𔻒f����flgEasySearch��ݒ�
	'���@�l�F
	'�X�@�V�F2007/11/01 LIS K.Kokubo �쐬
	'******************************************************************************
	Private Sub ChkSQLType()
		'�J���^�������E�ڍ׌�������
		If JT & JT2 & AC & AC2 & WT & ST1 & ST2 & ST3 & ST4 & ST5 & ST6 & PC & RC & SC & KW <> "" Then
			flgEasySearch = True
		Else
			flgEasySearch = False
		End If
	End Sub

	'******************************************************************************
	'�T�@�v�F�f�[�^�̐��������`�F�b�N
	'�쐬�ҁFLis K.Kokubo
	'�쐬���F2007/04/17 Lis K.Kokubo
	'�X�@�V�F
	'���@�l�F
	'******************************************************************************
	Private Sub ChkData()
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
	End Sub

	'******************************************************************************
	'�T�@�v�F���d���ڍ׌����̏���hidden���o�͂���
	'�쐬�ҁFLis K.Kokubo
	'�쐬���F2007/04/04 Lis K.Kokubo
	'�X�@�V�F
	'���@�l�F
	'******************************************************************************
	Public Sub DspConditionHidden()
		Response.Write "<input name=""CONF_SearchHopeJobTypeBigCode1"" type=""hidden"" value=""" & JobTypeBigCode1 & """>"
		Response.Write "<input name=""CONF_SearchHopeJobTypeCode1"" type=""hidden"" value=""" & JobTypeCode1 & """>"
		Response.Write "<input name=""CONF_SearchHopeJobTypeBigCode2"" type=""hidden"" value=""" & JobTypeBigCode2 & """>"
		Response.Write "<input name=""CONF_SearchHopeJobTypeCode2"" type=""hidden"" value=""" & JobTypeCode2 & """>"
		Response.Write "<input name=""CONF_SearchRailwayLineCode1"" type=""hidden"" value=""" & RailwayLineCode1 & """>"
		Response.Write "<input name=""CONF_SearchRailwayLineCode2"" type=""hidden"" value=""" & RailwayLineCode2 & """>"
		Response.Write "<input name=""CONF_SearchStationCode1"" type=""hidden"" value=""" & StationCode1 & """>"
		Response.Write "<input name=""CONF_SearchStationCode2"" type=""hidden"" value=""" & StationCode2 & """>"
		Response.Write "<input name=""CONF_SearchAreaCode1"" type=""hidden"" value=""" & AreaCode1 & """>"
		Response.Write "<input name=""CONF_SearchPrefectureCode1"" type=""hidden"" value=""" & PrefectureCode1 & """>"
		Response.Write "<input name=""CONF_SearchCity1"" type=""hidden"" value=""" & City1 & """>"
		Response.Write "<input name=""CONF_SearchAreaCode2"" type=""hidden"" value=""" & AreaCode2 & """>"
		Response.Write "<input name=""CONF_SearchPrefectureCode2"" type=""hidden"" value=""" & PrefectureCode2 & """>"
		Response.Write "<input name=""CONF_SearchCity2"" type=""hidden"" value=""" & City2 & """>"
		Response.Write "<input name=""CONF_SearchHopeWorkingTypeCode1"" type=""hidden"" value=""" & WorkingTypeCode1 & """>"
		Response.Write "<input name=""CONF_SearchHopeWorkingTypeCode2"" type=""hidden"" value=""" & WorkingTypeCode2 & """>"
		Response.Write "<input name=""CONF_SearchHopeWorkingTypeCode3"" type=""hidden"" value=""" & WorkingTypeCode3 & """>"
		Response.Write "<input name=""CONF_SearchHopeIndustryTypeCode"" type=""hidden"" value=""" & IndustryTypeCode & """>"
		Response.Write "<input name=""CONF_SearchPercentagePayFlag"" type=""hidden"" value=""" & PercentagePayFlag & """>"
		Response.Write "<input name=""CONF_SearchYearlyIncome"" type=""hidden"" value=""" & YearlyIncome & """>"
		Response.Write "<input name=""CONF_SearchMonthlyIncome"" type=""hidden"" value=""" & MonthlyIncome & """>"
		Response.Write "<input name=""CONF_SearchDailyIncome"" type=""hidden"" value=""" & DailyIncome & """>"
		Response.Write "<input name=""CONF_SearchHourlyIncome"" type=""hidden"" value=""" & HourlyIncome & """>"
		Response.Write "<input name=""CONF_SearchWorkStartHour"" type=""hidden"" value=""" & WorkStartHour & """>"
		Response.Write "<input name=""CONF_SearchWorkStartMinute"" type=""hidden"" value=""" & WorkStartMinute & """>"
		Response.Write "<input name=""CONF_SearchWorkEndHour"" type=""hidden"" value=""" & WorkEndHour & """>"
		Response.Write "<input name=""CONF_SearchWorkEndMinute"" type=""hidden"" value=""" & WorkEndMinute & """>"
		Response.Write "<input name=""CONF_SearchWeeklyHolidayType"" type=""hidden"" value=""" & WeeklyHolidayType & """>"
		Response.Write "<input name=""CONF_SearchAge"" type=""hidden"" value=""" & Age & """>"
		Response.Write "<input name=""CONF_SearchAgreementTerm"" type=""hidden"" value=""" & AgreementTerm & """>"
		Response.Write "<input name=""CONF_SearchLicenseGroupCode1"" type=""hidden"" value=""" & LicenseGroupCode1 & """>"
		Response.Write "<input name=""CONF_SearchLicenseCategoryCode1"" type=""hidden"" value=""" & LicenseCategoryCode1 & """>"
		Response.Write "<input name=""CONF_SearchLicenseCode1"" type=""hidden"" value=""" & LicenseCode1 & """>"
		Response.Write "<input name=""CONF_SearchOSCode1"" type=""hidden"" value=""" & OSCode1 & """>"
		Response.Write "<input name=""CONF_SearchApplicationCode1"" type=""hidden"" value=""" & ApplicationCode1 & """>"
		Response.Write "<input name=""CONF_SearchDevelopmentLanguageCode1"" type=""hidden"" value=""" & DevelopmentLanguageCode1 & """>"
		Response.Write "<input name=""CONF_SearchDatabaseCode1"" type=""hidden"" value=""" & DatabaseCode1 & """>"
		Response.Write "<input name=""CONF_SearchKeyword"" type=""hidden"" value=""" & Keyword & """>"
		Response.Write "<input name=""CONF_SearchKeywordFlag"" type=""hidden"" value=""" & KeywordFlag & """>"
		Response.Write "<input name=""CONF_SearchOrderCode"" type=""hidden"" value=""" & OrderCode & """>"
		'��������
		Response.Write "<input name=""CONF_PC"" type=""hidden"" value=""" & PC & """>"
		Response.Write "<input name=""CONF_RC"" type=""hidden"" value=""" & RC & """>"
		Response.Write "<input name=""CONF_SC"" type=""hidden"" value=""" & SC & """>"
		'���W
		Response.Write "<input name=""CONF_SP"" type=""hidden"" value=""" & SP & """>"
	End Sub

	'******************************************************************************
	'�T�@�v�F���d���ڍ׌����y�[�W�֓n��GET�p�����[�^�𐶐����Ď擾�B
	'�쐬�ҁFLis Kokubo
	'�쐬���F2007/02/27
	'���@���F
	'���@�l�F������
	'�@�@�@�F�p�����[�^���܂�URL�́AIE�̐�����2048�����܂łł���̂ŁA����ɍ��킹��B
	'******************************************************************************
	Public Function GetSearchParam()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sError

		Dim odSC
		Dim odResult
		Dim idxKey
		Dim aKeys
		Dim aValues

		GetSearchParam = ""

		If JobTypeBigCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sjtbig1=" & JobTypeBigCode1
		If JobTypeCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sjt1=" & JobTypeCode1
		If JobTypeBigCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;sjtbig2=" & JobTypeBigCode2
		If JobTypeCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;sjt2=" & JobTypeCode2
		If RailwayLineCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;src1=" & RailwayLineCode1
		If RailwayLineCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;src2=" & RailwayLineCode2
		If StationCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;ssc1=" & StationCode1
		If StationCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;ssc2=" & StationCode2
		If AreaCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sac1=" & AreaCode1
		If PrefectureCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;spc1=" & PrefectureCode1
		If City1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sct1=" & Server.URLEncode(City1)
		If AreaCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;sac2=" & AreaCode2
		If PrefectureCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;spc2=" & PrefectureCode2
		If City2 <> "" Then GetSearchParam = GetSearchParam & "&amp;sct2=" & Server.URLEncode(City2)
		If WorkingTypeCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;swt1=" & WorkingTypeCode1
		If WorkingTypeCode2 <> "" Then GetSearchParam = GetSearchParam & "&amp;swt2=" & WorkingTypeCode2
		If WorkingTypeCode3 <> "" Then GetSearchParam = GetSearchParam & "&amp;swt3=" & WorkingTypeCode3
		If IndustryTypeCode <> "" Then GetSearchParam = GetSearchParam & "&amp;sit=" & IndustryTypeCode
		If PercentagePayFlag <> "" Then GetSearchParam = GetSearchParam & "&amp;sppf=" & PercentagePayFlag
		If YearlyIncome <> "" Then GetSearchParam = GetSearchParam & "&amp;syi=" & YearlyIncome
		If MonthlyIncome <> "" Then GetSearchParam = GetSearchParam & "&amp;smi=" & MonthlyIncome
		If DailyIncome <> "" Then GetSearchParam = GetSearchParam & "&amp;sdi=" & DailyIncome
		If HourlyIncome <> "" Then GetSearchParam = GetSearchParam & "&amp;shi=" & HourlyIncome
		If WorkStartHour <> "" Then GetSearchParam = GetSearchParam & "&amp;swsh=" & WorkStartHour
		If WorkStartMinute <> "" Then GetSearchParam = GetSearchParam & "&amp;swsm=" & WorkStartMinute
		If WorkEndHour <> "" Then GetSearchParam = GetSearchParam & "&amp;sweh=" & WorkEndHour
		If WorkEndMinute <> "" Then GetSearchParam = GetSearchParam & "&amp;swem=" & WorkEndMinute
		If WeeklyHolidayType <> "" Then GetSearchParam = GetSearchParam & "&amp;swht=" & WeeklyHolidayType
		If Age <> "" Then GetSearchParam = GetSearchParam & "&amp;sage=" & Age
		If AgreementTerm <> "" Then GetSearchParam = GetSearchParam & "&amp;sat=" & AgreementTerm
		If LicenseGroupCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;slg1=" & LicenseGroupCode1
		If LicenseCategoryCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;slc1=" & LicenseCategoryCode1
		If LicenseCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sl1=" & LicenseCode1
		If OSCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sos1=" & OSCode1
		If ApplicationCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sap1=" & ApplicationCode1
		If DevelopmentLanguageCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sdl1=" & DevelopmentLanguageCode1
		If DatabaseCode1 <> "" Then GetSearchParam = GetSearchParam & "&amp;sdb1=" & DatabaseCode1
		If Keyword <> "" Then GetSearchParam = GetSearchParam & "&amp;skw=" & Server.URLEncode(Keyword)
		If KeywordFlag <> "" Then GetSearchParam = GetSearchParam & "&amp;skwflag=" & KeywordFlag
		If OrderCode <> "" Then GetSearchParam = GetSearchParam & "&amp;soc=" & OrderCode
		If Specialty <> "" Then GetSearchParam = GetSearchParam & "&amp;sst=" & Specialty
		If SP <> "" Then GetSearchParam = GetSearchParam & "&amp;sp=" & SP

		If GetSearchParam <> "" Then
			'����&amp;���H�ɕϊ�
			GetSearchParam = "?" & Mid(GetSearchParam, 6)

			'�h�d�̎d�l�̓p�����[�^�̏�����Q�O�S�W�o�C�g
			GetSearchParam = Left(GetSearchParam, 2048)
		End If
	End Function

	'******************************************************************************
	'�T�@�v�F���l�[�ڍ׌����r�p�k���擾
	'�쐬�ҁFLis Kokubo
	'�쐬���F2007/04/04
	'���@���F
	'���@�l�F
	'******************************************************************************
	Function GetSQLOrderSearchDetail()
		Dim sJoin		: sJoin = ""
		Dim sWhere		: sWhere = ""
		Dim sDeclare	: sDeclare = ""
		Dim sParams		: sParams = ""
		Dim iParamNo
		Dim sFrom
		Dim sTemp
		Dim sTemp2
		Dim sTemp3
		Dim aValue
		Dim idx
		Dim sSearchCondition

		'�f�[�^�������`�F�b�N
		Call ChkData()
		'�J���^�������E�ڍ׌�������
		Call ChkSQLType()

		'******************************************************************************
		'�E�� start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		iParamNo = 0
		If JobTypeBigCode1 & JobTypeCode1 & JobTypeBigCode2 & JobTypeCode2 <> "" Then
			If JobTypeBigCode1 & JobTypeCode1 <> "" Then
				sTemp = JobTypeBigCode1
				If JobTypeCode1 <> "" Then sTemp = JobTypeCode1

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "CJT.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " + '%' "

				iParamNo = iParamNo + 1
			End If

			If JobTypeBigCode2 & JobTypeCode2 <> "" Then
				sTemp = JobTypeBigCode2
				If JobTypeCode2 <> "" Then sTemp = JobTypeCode2

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "CJT.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " + '%' "

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CJT.OrderCode FROM C_JobType AS CJT WHERE (" & sTemp2 & ")) AS CJT ON VWOC.OrderCode = CJT.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'�E�� end
		'******************************************************************************

		'******************************************************************************
		'���� start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If RailwayLineCode1 & RailwayLineCode2 <> "" Then
			If RailwayLineCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vRailwayLineCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vRailwayLineCode" & iParamNo & " = N'" & RailwayLineCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vRailwayLineCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			If RailwayLineCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vRailwayLineCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vRailwayLineCode" & iParamNo & " = N'" & RailwayLineCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vRailwayLineCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT CNS.OrderCode "
			sJoin = sJoin & "FROM C_NearbyStation AS CNS "
			sJoin = sJoin & "INNER JOIN StationStop AS SS "
			sJoin = sJoin & "ON CNS.StationCode = SS.StationCode "
			sJoin = sJoin & "INNER JOIN B_RailwayLine AS BRL "
			sJoin = sJoin & "ON SS.RailwayLineCode = BRL.RailwayLineCode "
			sJoin = sJoin & "AND BRL.RailwayLineCode IN (" & sTemp & ") "
			sJoin = sJoin & ") AS CRL "
			sJoin = sJoin & "ON VWOC.OrderCode = CRL.OrderCode "
		End If

		'------------------------------------------------------------------------------
		'���� end
		'******************************************************************************

		'******************************************************************************
		'�w start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If StationCode1 & StationCode2 <> "" Then
			If StationCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vStationCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vStationCode" & iParamNo & " = N'" & StationCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vStationCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			If StationCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vStationCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vStationCode" & iParamNo & " = N'" & StationCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vStationCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT CNS.OrderCode "
			sJoin = sJoin & "FROM C_NearbyStation AS CNS "
			sJoin = sJoin & "WHERE CNS.StationCode IN (" & sTemp & ") "
			sJoin = sJoin & ") AS CNS "
			sJoin = sJoin & "ON VWOC.OrderCode = CNS.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'�w end
		'******************************************************************************

		'******************************************************************************
		'��]�Ζ��n start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		iParamNo = 0
		If AreaCode1 & AreaCode2 <> "" Then
			sTemp = ""
			If AreaCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vAreaCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vAreaCode" & iParamNo & " = N'" & AreaCode1 & "'"

				sTemp = "AREA.AreaCode = @vAreaCode" & iParamNo & " "

				If PrefectureCode1 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vPrefectureCode" & iParamNo & " VARCHAR(3)"
					sParams = sParams & ",@vPrefectureCode" & iParamNo & " = N'" & PrefectureCode1 & "'"

					If sTemp <> "" Then sTemp = sTemp & "AND "
					sTemp = sTemp & "CWP.WorkingPlacePrefectureCode = @vPrefectureCode" & iParamNo & " "
				End If

				If PrefectureCode1 <> "" And City1 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vCity" & iParamNo & " VARCHAR(100)"
					sParams = sParams & ",@vCity" & iParamNo & " = N'" & City1 & "'"

					If sTemp <> "" Then sTemp = sTemp & "AND "
					sTemp = sTemp & "CWP.WorkingPlaceCity LIKE '%' + @vCity" & iParamNo & " + '%' "
				End If

				iParamNo = iParamNo + 1
			End If

			If sTemp <> "" Then
				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = "(" & sTemp & ") "
			End If

			sTemp = ""
			If AreaCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vAreaCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vAreaCode" & iParamNo & " = N'" & AreaCode2 & "'"

				sTemp = "AREA.AreaCode = @vAreaCode" & iParamNo & " "

				If PrefectureCode2 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vPrefectureCode" & iParamNo & " VARCHAR(3)"
					sParams = sParams & ",@vPrefectureCode" & iParamNo & " = N'" & PrefectureCode2 & "'"

					If sTemp <> "" Then sTemp = sTemp & "AND "
					sTemp = sTemp & "CWP.WorkingPlacePrefectureCode = @vPrefectureCode" & iParamNo & " "
				End If

				If PrefectureCode2 <> "" And City2 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vCity" & iParamNo & " VARCHAR(200)"
					sParams = sParams & ",@vCity" & iParamNo & " = N'" & City2 & "'"

					If sTemp <> "" Then sTemp = sTemp & "AND "
					sTemp = sTemp & "CWP.WorkingPlaceCity LIKE '%' + @vCity" & iParamNo & " + '%' "
				End If

				iParamNo = iParamNo + 1
			End If

			If sTemp <> "" Then
				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "(" & sTemp & ") "
			End If

			sJoin = sJoin & "INNER JOIN ( "
			sJoin = sJoin & "SELECT DISTINCT CWP.OrderCode "
			sJoin = sJoin & "FROM C_Info AS CWP "
			sJoin = sJoin & "INNER JOIN Area AS AREA ON CWP.WorkingPlacePrefectureCode = AREA.PrefectureCode "
			sJoin = sJoin & "WHERE " & sTemp2 & " "
			sJoin = sJoin & ") AS CWP "
			sJoin = sJoin & "ON VWOC.OrderCode = CWP.OrderCode "
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

			sJoin = sJoin & "INNER JOIN ( "
			sJoin = sJoin & "SELECT DISTINCT CWT.OrderCode "
			sJoin = sJoin & "FROM C_WorkingType AS CWT "
			sJoin = sJoin & "WHERE CWT.WorkingTypeCode IN (" & sTemp & ") "
			sJoin = sJoin & ") AS CWT "
			sJoin = sJoin & "ON VWOC.OrderCode = CWT.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'��]�Ζ��`�� end
		'******************************************************************************

		'******************************************************************************
		'��]�Ǝ� start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If IndustryTypeCode1 & IndustryTypeCode2 & IndustryTypeCode3 <> "" Then
			If IndustryTypeCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vIndustryTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vIndustryTypeCode" & iParamNo & " = N'" & IndustryTypeCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vIndustryTypeCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			If IndustryTypeCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vIndustryTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vIndustryTypeCode" & iParamNo & " = N'" & IndustryTypeCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vIndustryTypeCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			If IndustryTypeCode3 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vIndustryTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vIndustryTypeCode" & iParamNo & " = N'" & IndustryTypeCode3 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vIndustryTypeCode" & iParamNo

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN ( "
			sJoin = sJoin & "SELECT CIDST.CompanyCode "
			sJoin = sJoin & "FROM CompanyInfo AS CIDST "
			sJoin = sJoin & "WHERE CIDST.IndustryType IN (" & sTemp & ") "
			sJoin = sJoin & ") AS CIDST "
			sJoin = sJoin & "ON VWOC.CompanyCode = CIDST.CompanyCode "
		End If
		'------------------------------------------------------------------------------
		'��]�Ǝ� end
		'******************************************************************************

		'******************************************************************************
		'���� start
		'------------------------------------------------------------------------------
		'���o�����}�A��w���������AUI�^�[���A�x���P�Q�O���ȏ�
		sTemp = ""
		If InexperiencedPersonFlag = "1" Or UtilizeLanguageFlag = "1" Or UITurnFlag = "1" Or ManyHolidayFlag = "1" Then
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

			sJoin = sJoin & "INNER JOIN C_SupplementInfo AS CSP ON VWOC.OrderCode = CSP.OrderCode AND " & sTemp & " "
		End If

		'�t���b�N�X�^�C��
		sTemp = ""
		If FlexFlag = "1" Then
			sJoin = sJoin & "INNER JOIN CompanyInfo AS CMPFLEX ON VWOC.CompanyCode = CMPFLEX.CompanyCode AND CMPFLEX.CompanyKbn = '1' AND CMPFLEX.FlexTime = 'ON' "
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
		If YearlyIncome & MonthlyIncome & DailyIncome & HourlyIncome & PercentagePayFlag <> "" Then
			If YearlyIncome <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vYearlyIncome INT "
				sParams = sParams & ",@vYearlyIncome = " & YearlyIncome

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "CSLY.YearlyIncomeMin >= @vYearlyIncome "
			End If

			If MonthlyIncome <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vMonthlyIncome INT "
				sParams = sParams & ",@vMonthlyIncome = " & MonthlyIncome

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "CSLY.MonthlyIncomeMin >= @vMonthlyIncome "
			End If

			If DailyIncome <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vDailyIncome INT "
				sParams = sParams & ",@vDailyIncome = " & DailyIncome

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "CSLY.DailyIncomeMin >= @vDailyIncome "
			End If

			If HourlyIncome <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vHourlyIncome INT "
				sParams = sParams & ",@vHourlyIncome = " & HourlyIncome

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "CSLY.HourlyIncomeMin >= @vHourlyIncome "
			End If

			If sTemp <> "" Then sTemp = "(" & sTemp & ") "

			'������
			If PercentagePayFlag <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vPercentagePayFlag VARCHAR(1)"
				sParams = sParams & ",@vPercentagePayFlag = N'" & PercentagePayFlag & "'"

				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = "CSLY.PercentagePayFlag = @vPercentagePayFlag "
			End If

			sJoin = sJoin & "INNER JOIN C_Info AS CSLY ON VWOC.OrderCode = CSLY.OrderCode AND " & sTemp & " "
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
				sTemp = sTemp & "CWTM.WorkStartTime >= @vWorkStartHour + @vWorkStartMinute "
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
				sTemp = sTemp & "CWTM.WorkEndTime <= @vWorkEndHour + @vWorkEndMinute "
			End If

			If WorkStartHour <> "" And WorkEndHour <> "" Then
				If WorkStartHour < WorkEndHour Then
					'�Ζ��J�n���� < �Ζ��I�����Ԃ̏ꍇ�A��Ԃ̋Ɩ����Ԃ������悤�ɂ���
					sTemp2 = "AND CWTM.WorkStartTime < CWTM.WorkEndTime "
				End If
			End If

			sJoin = sJoin & "INNER JOIN C_WorkingCondition AS CWTM ON VWOC.OrderCode = CWTM.OrderCode AND " & sTemp & sTemp2
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

			sJoin = sJoin & "INNER JOIN C_Info AS CWHT ON VWOC.OrderCode = CWHT.OrderCode AND " & sTemp
		End If
		'------------------------------------------------------------------------------
		'�T�x end
		'******************************************************************************

		'******************************************************************************
		'�N�� start
		'------------------------------------------------------------------------------
		sTemp = ""
		If Age <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vAge INT "
			sParams = sParams & ",@vAge = " & Age

			sTemp = sTemp & "(@vAge BETWEEN ISNULL(CAGE.AgeMin, 0) AND ISNULL(CAGE.AgeMax, 255)) "

			sJoin = sJoin & "INNER JOIN C_Info AS CAGE ON VWOC.OrderCode = CAGE.OrderCode AND " & sTemp
		End If
		'------------------------------------------------------------------------------
		' �N�� end
		'******************************************************************************

		'******************************************************************************
		'�_����� start
		'------------------------------------------------------------------------------
		sTemp = ""
		If IsRE(AgreementTerm, "^[123]$", True) = True Then
			If AgreementTerm = "1" Then
				sJoin = sJoin & "INNER JOIN (SELECT OrderCode FROM C_Temp WHERE WorkPeriod <= 1 UNION SELECT OrderCode FROM C_Undertake WHERE WorkPeriod <= 1 UNION SELECT OrderCode FROM C_TTP WHERE WorkPeriod <= 1) AS CAT ON VWOC.OrderCode = CAT.OrderCode "
			ElseIf AgreementTerm = "2" Then
				sJoin = sJoin & "INNER JOIN (SELECT OrderCode FROM C_Temp WHERE WorkPeriod <= 2 UNION SELECT OrderCode FROM C_Undertake WHERE WorkPeriod <= 2 UNION SELECT OrderCode FROM C_TTP WHERE WorkPeriod <= 2) AS CAT ON VWOC.OrderCode = CAT.OrderCode "
			ElseIf AgreementTerm = "3" Then
				sJoin = sJoin & "INNER JOIN (SELECT OrderCode FROM C_Temp WHERE WorkPeriod > 3 UNION SELECT OrderCode FROM C_Undertake WHERE WorkPeriod > 3 UNION SELECT OrderCode FROM C_TTP WHERE WorkPeriod > 3) AS CAT ON VWOC.OrderCode = CAT.OrderCode "
			End If
		End If
		'------------------------------------------------------------------------------
		'�_����� end
		'******************************************************************************

		'******************************************************************************
		'�ۗL���i start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If LicenseGroupCode1 <> "" Then
			'�啪��
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vLicenseGroupCode" & iParamNo & " VARCHAR(2)"
			sParams = sParams & ",@vLicenseGroupCode" & iParamNo & " = N'" & LicenseGroupCode1 & "'"

			If sTemp <> "" Then sTemp = sTemp & "AND "
			sTemp = sTemp & "CL.GroupCode = @vLicenseGroupCode" & iParamNo & " "

			If LicenseCategoryCode1 <> "" Then
				'������
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vLicenseCategoryCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vLicenseCategoryCode" & iParamNo & " = N'" & LicenseCategoryCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CL.CategoryCode = @vLicenseCategoryCode" & iParamNo & " "

				If LicenseCode1 <> "" Then
					'������
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vLicenseCode" & iParamNo & " VARCHAR(2)"
					sParams = sParams & ",@vLicenseCode" & iParamNo & " = N'" & LicenseCode1 & "'"

					If sTemp <> "" Then sTemp = sTemp & "AND "
					sTemp = sTemp & "CL.Code = @vLicenseCode" & iParamNo & " "
				End If
			End If

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CL.OrderCode FROM C_License AS CL WHERE " & sTemp & ") AS CL ON VWOC.OrderCode = CL.OrderCode "
			iParamNo = iParamNo + 1
		End If
		'------------------------------------------------------------------------------
		'�ۗL���i end
		'******************************************************************************

		'******************************************************************************
		'�X�L�� start
		'------------------------------------------------------------------------------
		'OS
		sTemp = ""
		iParamNo = 0
		If OSCode1 <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo & " VARCHAR(20), @vSkillCode" & iParamNo & " VARCHAR(3) "
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo & " = N'OS',@vSkillCode" & iParamNo & " = N'" & OSCode1 & "'"

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CSKL.OrderCode FROM C_Skill AS CSKL WHERE CSKL.CategoryCode = @vSkillCategoryCode" & iParamNo & " AND CSKL.Code = @vSkillCode" & iParamNo & ") AS CSKL" & iParamNo & " ON VWOC.OrderCode = CSKL" & iParamNo & ".OrderCode "
			iParamNo = iParamNo + 1
		End If

		'�A�v���P�[�V����
		sTemp = ""
		If ApplicationCode1 <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo & " VARCHAR(20), @vSkillCode" & iParamNo & " VARCHAR(3) "
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo & " = N'Application',@vSkillCode" & iParamNo & " = N'" & ApplicationCode1 & "'"

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CSKL.OrderCode FROM C_Skill AS CSKL WHERE CSKL.CategoryCode = @vSkillCategoryCode" & iParamNo & " AND CSKL.Code = @vSkillCode" & iParamNo & ") AS CSKL" & iParamNo & " ON VWOC.OrderCode = CSKL" & iParamNo & ".OrderCode "
			iParamNo = iParamNo + 1
		End If

		'�J������
		sTemp = ""
		If DevelopmentLanguageCode1 <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo & " VARCHAR(20), @vSkillCode" & iParamNo & " VARCHAR(3) "
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo & " = N'DevelopmentLanguage',@vSkillCode" & iParamNo & " = N'" & DevelopmentLanguageCode1 & "'"

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CSKL.OrderCode FROM C_Skill AS CSKL WHERE CSKL.CategoryCode = @vSkillCategoryCode" & iParamNo & " AND CSKL.Code = @vSkillCode" & iParamNo & ") AS CSKL" & iParamNo & " ON VWOC.OrderCode = CSKL" & iParamNo & ".OrderCode "
			iParamNo = iParamNo + 1
		End If

		'�f�[�^�x�[�X
		sTemp = ""
		If DatabaseCode1 <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSkillCategoryCode" & iParamNo & " VARCHAR(20), @vSkillCode" & iParamNo & " VARCHAR(3) "
			sParams = sParams & ",@vSkillCategoryCode" & iParamNo & " = N'Database',@vSkillCode" & iParamNo & " = N'" & DatabaseCode1 & "'"

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CSKL.OrderCode FROM C_Skill AS CSKL WHERE CSKL.CategoryCode = @vSkillCategoryCode" & iParamNo & " AND CSKL.Code = @vSkillCode" & iParamNo & ") AS CSKL" & iParamNo & " ON VWOC.OrderCode = CSKL" & iParamNo & ".OrderCode "
			iParamNo = iParamNo + 1
		End If
		'------------------------------------------------------------------------------
		'�X�L�� end
		'******************************************************************************

		'******************************************************************************
		'�L�[���[�h start
		'------------------------------------------------------------------------------
		sTemp = ""
		If Keyword <> "" Then
			aValue = Split(Replace(Keyword, "�@", " "), " ")
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

			sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY CFTN.OrderCode) AS Num, CFTN.OrderCode FROM C_FullTextNavi AS CFTN WHERE CONTAINS(CFTN.Text, @vKeyword)) AS CFTN ON VWOC.OrderCode = CFTN.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'�L�[���[�h end
		'******************************************************************************

		'******************************************************************************
		'���R�[�h start
		'------------------------------------------------------------------------------
		If OrderCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vOrderCode VARCHAR(8) "
			sParams = sParams & ",@vOrderCode = N'" & OrderCode & "'"

			sJoin = ""
			sWhere = "WHERE VWOC.OrderCode = @vOrderCode "
		End If
		'------------------------------------------------------------------------------
		'���R�[�h end
		'******************************************************************************

		'******************************************************************************
		'�O��\�����̍ŐV���R�[�h start
		'------------------------------------------------------------------------------
		If BOC <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vBeforeOrderCode VARCHAR(8) "
			sParams = sParams & ",@vBeforeOrderCode = N'" & BOC & "'"

			sWhere = "WHERE VWOC.OrderCode > @vBeforeOrderCode "
		End If
		'------------------------------------------------------------------------------
		'�O��\�����̍ŐV���R�[�h end
		'******************************************************************************

		If flgEasySearch = False And sJoin & sWhere <> "" Then
			If CStr(Top) <> "" Then Top = "TOP " & Top
			GetSQLOrderSearchDetail = "" & _
				"SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " & _
				"SELECT " & Top & " VWOC.OrderCode " & _
				",VWOC.SortNum " & _
				",VWOC.UpdateDay " & _
				"FROM vw_OrderCode AS VWOC " & _
				sJoin & _
				sWhere & _
				"ORDER BY VWOC.SortNum ASC, VWOC.UpdateDay DESC "

			GetSQLOrderSearchDetail = "" & _
				"/*�i�r�E���l�[�ڍ׌���*/ " & _
				"EXEC sp_executesql N'" & Replace(GetSQLOrderSearchDetail, "'", "''") & "'"
			If sDeclare <> "" Then GetSQLOrderSearchDetail = GetSQLOrderSearchDetail & ",N'" & sDeclare & "'" & sParams
		Else
			GetSQLOrderSearchDetail = GetSQLOrderSearchEasy()
		End If

		If sSearchCondition <> "" Then
			sSearchCondition = "<table class=""pattern1"" border=""0"" style=""width:600px;""><thead><tr><th colspan=""2"" style=""width:588px;"">��������</th></tr></thead><tbody>" & sSearchCondition & "</tbody></table>"
		Else
			sSearchCondition = "�Ȃ�"
		End If
	End Function

	'******************************************************************************
	'�T�@�v�F���l�[�����k�n�f�������݂r�p�k���擾
	'�쐬�ҁFLis Kokubo
	'�쐬���F2007/04/04
	'���@���F
	'���@�l�F
	'******************************************************************************
	Public Function GetSQLWriteLog()
		Dim sTmpJT
		sTmpJT = JT
		If JT2 = "" Then sTmpJT = JT2

		If flgEasySearch = True Then
			'�J���^���������O
			GetSQLWriteLog = "EXEC up_Reg_LOG_SearchOrder '" & G_USERID & "'" & _
				",'" & ChkSQLStr(Request.ServerVariables("REMOTE_ADDR")) & "'" & _
				",'" & ChkSQLStr(Session.SessionID) & "'" & _
				",'" & ChkSQLStr(Request.ServerVariables("URL")) & "?" & ChkSQLStr(Request.ServerVariables("QUERY_STRING")) & "'" & _
				",'" & ChkSQLStr(Request.ServerVariables("HTTP_REFERER")) & "'" & _
				",'" & sTmpJT & "'" & _
				",'" & WT & "'" & _
				",'" & AC & "'" & _
				",'" & AC2 & "'" & _
				",'" & Specialty & "'" & _
				",''" & _
				",'" & RC & "'" & _
				",'" & SC & "'" & _
				",'" & KW & "'" & _
				",'" & Replace(SQLOrderSearch, "'", "''") & "'"
		Else
			'�ڍ׌������O
			GetSQLWriteLog = "EXEC up_Reg_LOG_SearchOrderDetail '" & G_USERID & "'" & _
				",'" & ChkSQLStr(Request.ServerVariables("REMOTE_ADDR")) & "'" & _
				",'" & ChkSQLStr(Session.SessionID) & "'" & _
				",'" & ChkSQLStr(Request.ServerVariables("URL")) & "?" & ChkSQLStr(Request.ServerVariables("QUERY_STRING")) & "'" & _
				",'" & ChkSQLStr(Request.ServerVariables("HTTP_REFERER")) & "'" & _
				",'" & JobTypeCode1 & "'" & _
				",'" & JobTypeCode2 & "'" & _
				",'" & RailwayLineCode1 & "'" & _
				",'" & StationCode1 & "'" & _
				",'" & RailwayLineCode2 & "'" & _
				",'" & StationCode2 & "'" & _
				",'" & AreaCode1 & "'" & _
				",'" & PrefectureCode1 & "'" & _
				",'" & City1 & "'" & _
				",'" & AreaCode2 & "'" & _
				",'" & PrefectureCode2 & "'" & _
				",'" & City2 & "'" & _
				",'" & WorkingTypeCode1 & "'" & _
				",'" & WorkingTypeCode2 & "'" & _
				",'" & WorkingTypeCode3 & "'" & _
				",'" & IndustryTypeCode1 & "'" & _
				",'" & IndustryTypeCode2 & "'" & _
				",'" & IndustryTypeCode3 & "'" & _
				",'" & PercentagePayFlag & "'" & _
				",'" & YearlyIncome & "'" & _
				",'" & MonthlyIncome & "'" & _
				",'" & DailyIncome & "'" & _
				",'" & HourlyIncome & "'" & _
				",'" & WorkStartHour & WorkStartMinute & "'" & _
				",'" & WorkEndHour & WorkEndMinute & "'" & _
				",'" & WeeklyHolidayType & "'" & _
				",'" & Age & "'" & _
				",'" & AgreementTerm & "'" & _
				",'" & LicenseGroupCode1 & "'" & _
				",'" & LicenseCategoryCode1 & "'" & _
				",'" & LicenseCode1 & "'" & _
				",'" & OSCode1 & "'" & _
				",'" & ApplicationCode1 & "'" & _
				",'" & DevelopmentLanguageCode1 & "'" & _
				",'" & DatabaseCode1 & "'" & _
				",'" & InexperiencedPersonFlag & "'" & _
				",'" & UtilizeLanguageFlag & "'" & _
				",'" & UITurnFlag & "'" & _
				",'" & ManyHolidayFlag & "'" & _
				",'" & FlexFlag & "'" & _
				",'" & Keyword & "'" & _
				",'" & Replace(SQLOrderSearch, "'", "''") & "'"
		End If
	End Function

	'******************************************************************************
	'�T�@�v�F���l�[�J���^�������r�p�k���擾
	'�쐬�ҁFLis Kokubo
	'�쐬���F2007/04/04
	'���@���F
	'���@�l�F
	'******************************************************************************
	Function GetSQLOrderSearchEasy()
		Dim sJoin		: sJoin = ""
		Dim sWhere		: sWhere = ""
		Dim sDeclare	: sDeclare = ""
		Dim sParams		: sParams = ""
		Dim iParamNo
		Dim sFrom
		Dim sTemp
		Dim sTemp2
		Dim sTemp3
		Dim aValue
		Dim idx
		Dim sSearchCondition

		GetSQLOrderSearchEasy = ""

		'******************************************************************************
		'�E�� start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		iParamNo = 0
		If JT & JT2 <> "" Then
			If JT2 <> "" Then
				sTemp = JobTypeCode2
				If Len(JobTypeCode2) < 7 Then sTemp = sTemp & "%"

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & JT2 & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "CJT.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " "

				iParamNo = iParamNo + 1
			ElseIf JT <> "" Then
				sTemp = JobTypeCode1
				If Len(JobTypeCode1) < 7 Then sTemp = sTemp & "%"

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & JT & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "CJT.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " + '%' "

				iParamNo = iParamNo + 1
			End If

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT CJT.OrderCode FROM C_JobType AS CJT WHERE (" & sTemp2 & ")) AS CJT ON VWOC.OrderCode = CJT.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'�E�� end
		'******************************************************************************

		'******************************************************************************
		'��]�Ζ��n start
		'------------------------------------------------------------------------------
		sTemp = ""
		sTemp2 = ""
		iParamNo = 0
		If AC & AC2 <> "" Then
			sTemp = ""
			If AC <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vAreaCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vAreaCode" & iParamNo & " = N'" & AC & "'"

				sTemp = "AREA.AreaCode = @vAreaCode" & iParamNo & " "
			End If

			If AC2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vPrefectureCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vPrefectureCode" & iParamNo & " = N'" & AC2 & "'"

				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CWP.WorkingPlacePrefectureCode = @vPrefectureCode" & iParamNo & " "
			End If

			iParamNo = iParamNo + 1

			sJoin = sJoin & "INNER JOIN ( "
			sJoin = sJoin & "SELECT DISTINCT CWP.OrderCode "
			sJoin = sJoin & "FROM C_Info AS CWP "
			sJoin = sJoin & "INNER JOIN Area AS AREA ON CWP.WorkingPlacePrefectureCode = AREA.PrefectureCode "
			sJoin = sJoin & "WHERE " & sTemp & " "
			sJoin = sJoin & ") AS CWP "
			sJoin = sJoin & "ON VWOC.OrderCode = CWP.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'��]�Ζ��n end
		'******************************************************************************

		'******************************************************************************
		'���� start
		'------------------------------------------------------------------------------
		'���o�����}�A��w���������AUI�^�[���A�x���P�Q�O���ȏ�
		sTemp = ""
		If Len(ST) >= 6 Then
			'���o�����}
			If Mid(ST, 1, 1) = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.InexperiencedPersonFlag = '1' "
			End If

			'��w��������
			If Mid(ST, 2, 1) = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.UtilizeLanguageFlag = '1' "
			End If

			'UI�^�[��
			If Mid(ST, 4, 1) = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.UITurnFlag = '1' "
			End If

			'�x���P�Q�O���ȏ�
			If Mid(ST, 5, 1) = "1" Then
				If sTemp <> "" Then sTemp = sTemp & "AND "
				sTemp = sTemp & "CSP.ManyHolidayFlag = '1' "
			End If

			sJoin = sJoin & "INNER JOIN C_SupplementInfo AS CSP ON VWOC.OrderCode = CSP.OrderCode AND " & sTemp & " "

			'�t���b�N�X�^�C��
			sTemp = ""
			If Mid(ST, 6, 1) = "1" Then
				sJoin = sJoin & "INNER JOIN CompanyInfo AS CMPFLEX ON VWOC.CompanyCode = CMPFLEX.CompanyCode AND CMPFLEX.CompanyKbn = '1' AND CMPFLEX.FlexTime = 'ON' "
			End If

'			'�h��
'			If Mid(ST, 3, 1) = "1" Then
'				If InStr(sJoin, "INNER JOIN C_WorkingType AS CWT") = 0 Then sJoin = sJoin & "INNER JOIN C_WorkingType AS CWT ON CI.OrderCode = CWT.OrderCode "
'				If sWhere <> "" Then sWhere = sWhere & "AND "
'				sWhere = sWhere & "CWT.WorkingTypeCode IN ('001', '004') "
'			End If
		End If
		'------------------------------------------------------------------------------
		'���� end
		'******************************************************************************

		'******************************************************************************
		'���� start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If RC <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vRailwayLineCode" & iParamNo & " VARCHAR(7)"
			sParams = sParams & ",@vRailwayLineCode" & iParamNo & " = N'" & RC & "'"

			If sTemp <> "" Then sTemp = sTemp & ","
			sTemp = sTemp & "@vRailwayLineCode" & iParamNo

			iParamNo = iParamNo + 1

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT CNS.OrderCode "
			sJoin = sJoin & "FROM C_NearbyStation AS CNS "
			sJoin = sJoin & "INNER JOIN StationStop AS SS "
			sJoin = sJoin & "ON CNS.StationCode = SS.StationCode "
			sJoin = sJoin & "INNER JOIN B_RailwayLine AS BRL "
			sJoin = sJoin & "ON SS.RailwayLineCode = BRL.RailwayLineCode "
			sJoin = sJoin & "AND BRL.RailwayLineCode IN (" & sTemp & ") "
			sJoin = sJoin & ") AS CRL "
			sJoin = sJoin & "ON VWOC.OrderCode = CRL.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'���� end
		'******************************************************************************

		'******************************************************************************
		'�w start
		'------------------------------------------------------------------------------
		sTemp = ""
		iParamNo = 0
		If SC <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vStationCode" & iParamNo & " VARCHAR(7)"
			sParams = sParams & ",@vStationCode" & iParamNo & " = N'" & SC & "'"

			If sTemp <> "" Then sTemp = sTemp & ","
			sTemp = sTemp & "@vStationCode" & iParamNo

			iParamNo = iParamNo + 1

			sJoin = sJoin & "INNER JOIN ("
			sJoin = sJoin & "SELECT DISTINCT CNS.OrderCode "
			sJoin = sJoin & "FROM C_NearbyStation AS CNS "
			sJoin = sJoin & "WHERE CNS.StationCode IN (" & sTemp & ") "
			sJoin = sJoin & ") AS CNS "
			sJoin = sJoin & "ON VWOC.OrderCode = CNS.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'�w end
		'******************************************************************************

		'******************************************************************************
		'�L�[���[�h start
		'------------------------------------------------------------------------------
		sTemp = ""
		If KW <> "" Then
			aValue = Split(Replace(KW, "�@", " "), " ")
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

			sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY CFTN.OrderCode) AS Num, CFTN.OrderCode FROM C_FullTextNavi AS CFTN WHERE CONTAINS(CFTN.Text, @vKeyword)) AS CFTN ON VWOC.OrderCode = CFTN.OrderCode "
		End If
		'------------------------------------------------------------------------------
		'�L�[���[�h end
		'******************************************************************************

		If CStr(Top) <> "" Then Top = "TOP " & Top
		GetSQLOrderSearchEasy = "" & _
			"SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " & _
			"SELECT " & Top & " VWOC.OrderCode " & _
			",VWOC.SortNum " & _
			",VWOC.UpdateDay " & _
			"FROM vw_OrderCode AS VWOC " & _
			sJoin & _
			sWhere & _
			"ORDER BY VWOC.SortNum ASC, VWOC.UpdateDay DESC "

		GetSQLOrderSearchEasy = "" & _
			"/*�i�r�E���l�[�J���^������*/ " & _
			"EXEC sp_executesql N'" & Replace(GetSQLOrderSearchEasy, "'", "''") & "'"
		If sDeclare <> "" Then GetSQLOrderSearchEasy = GetSQLOrderSearchEasy & ",N'" & sDeclare & "'" & sParams
	End Function

	'******************************************************************************
	'�T�@�v�F���l�[�ڍ׌��������o�͂g�s�l�k���擾
	'�쐬�ҁFLis Kokubo
	'�쐬���F2007/04/04
	'���@���F
	'���@�l�F
	'******************************************************************************
	Public Function GetHtmlSearchCondition()
		Dim sTemp
		Dim sTemp2

		If flgEasySearch = True Then Exit Function

		GetHtmlSearchCondition = ""

		'�E��
		sTemp2 = ""
		If JobTypeBigCode1 & JobTypeCode1 & JobTypeBigCode2 & JobTypeCode2 <> "" Then
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
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�E��</th><td style=""width:439px;"">" & sTemp2 & "</td></tr>"
		End If

		'�Ζ��n
		sTemp = ""
		If AreaCode1 & PrefectureCode1 & City1 & RailwayLineCode1 & RailwayLineCode1 & AreaCode2 & PrefectureCode2 & City2 & RailwayLineCode2 & RailwayLineCode2 <> "" Then
			If AreaCode1 & PrefectureCode1 & City1 & RailwayLineCode1 & RailwayLineCode1 <> "" Then
				'�G���A
				sTemp = sTemp & AreaName1

				'�s���{��
				If PrefectureCode1 <> "" Then
					sTemp = sTemp & "�@"
					sTemp = sTemp & PrefectureName1

					'�s��S
					If City1 <> "" Then
						sTemp = sTemp & "�@"
						sTemp = sTemp & City1
					End If

					'����
					If RailwayLineCode1 <> "" Then
						sTemp = sTemp & "�@"
						sTemp = sTemp & RailwayLineName1
					End If

					'�w
					If RailwayLineCode2 <> "" Then
						If sTemp <> "" Then sTemp = sTemp & "�@"
						sTemp = sTemp & StationName1 & "�w"
					End If
				End If
			End If

			If AreaCode2 & PrefectureCode2 & City2 & RailwayLineCode2 & RailwayLineCode2 <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "<br>"
				'�G���A
				sTemp = sTemp & AreaName2

				'�s���{��
				If PrefectureCode2 <> "" Then
					sTemp = sTemp & "�@"
					sTemp = sTemp & PrefectureName2

					'�s��S
					If City2 <> "" Then
						sTemp = sTemp & "�@"
						sTemp = sTemp & City2
					End If

					'����
					If RailwayLineCode2 <> "" Then
						sTemp = sTemp & "�@"
						sTemp = sTemp & RailwayLineName2
					End If

					'�w
					If RailwayLineCode2 <> "" Then
						If sTemp <> "" Then sTemp = sTemp & "�@"
						sTemp = sTemp & StationName2 & "�w"
					End If
				End If
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�Ζ��n</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
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
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�Ζ��`��</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'�Ǝ�
		sTemp = ""
		If IndustryTypeCode1 & IndustryTypeCode2 & IndustryTypeCode3 <> "" Then
			If IndustryTypeCode1 <> "" Then sTemp = sTemp & IndustryTypeName1
			If IndustryTypeCode2 <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "�@"
				sTemp = sTemp & IndustryTypeName2
			End If
			If IndustryTypeCode3 <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "�@"
				sTemp = sTemp & IndustryTypeName3
			End If
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�Ǝ�</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'������
		sTemp = ""
		If PercentagePayFlag & YearlyIncome & MonthlyIncome & DailyIncome & HourlyIncome <> "" Then
			If PercentagePayFlag = "1" Then
				sTemp = sTemp & "����������"
			ElseIf PercentagePayFlag = "0" Then
				sTemp = sTemp & "�������Ȃ�"
			End If
			If YearlyIncome <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "<br>"
				sTemp = sTemp & "�N���F" & YearlyIncome & "�`"
			End If
			If MonthlyIncome <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "<br>"
				sTemp = sTemp & "�����F" & MonthlyIncome & "�`"
			End If
			If DailyIncome <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "<br>"
				sTemp = sTemp & "�����F" & DailyIncome & "�`"
			End If
			If HourlyIncome <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "<br>"
				sTemp = sTemp & "�����F" & HourlyIncome & "�`"
			End If

			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">���^</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'����
		sTemp = ""
		If InexperiencedPersonFlag & UtilizeLanguageFlag & TempFlag & UITurnFlag & ManyHolidayFlag & FlexFlag <> "" Then
			If InexperiencedPersonFlag = "1" Then sTemp = sTemp & "�u���o���҂n�j�v"
			If UtilizeLanguageFlag = "1" Then sTemp = sTemp & "�u��w���������v"
			If TempFlag = "1" Then sTemp = sTemp & "�u�h���v"
			If UITurnFlag = "1" Then sTemp = sTemp & "�u�t�h�^�[�����}�v"
			If ManyHolidayFlag = "1" Then sTemp = sTemp & "�u�x���P�Q�O���ȏ�v"
			If FlexFlag = "1" Then sTemp = sTemp & "�u�t���b�N�X�v"

			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">����</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'�A�Ǝ���
		sTemp = ""
		If WorkStartHour & WorkStartMinute & WorkEndHour & WorkEndMinute <> "" Then
			If WorkStartHour & WorkStartMinute <> "" Then sTemp = sTemp & "�A�ƊJ�n���ԁF" & WorkStartHour & ":" & WorkStartMinute & "&nbsp;�ȍ~"
			If WorkEndHour & WorkEndMinute <> "" And sTemp <> "" Then sTemp = sTemp & "<br>"
			If WorkEndHour & WorkEndMinute <> "" Then sTemp = sTemp & "�A�ƏI�����ԁF" & WorkEndHour & ":" & WorkEndMinute & "&nbsp;�ȑO"

			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�A�Ǝ���</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'�T�x���
		sTemp = ""
		If WeeklyHolidayType <> "" Then
			sTemp = sTemp & WeeklyHolidayTypeName
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�T�x���</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'�N��
		sTemp = ""
		If Age <> "" Then
			sTemp = sTemp & Age & "�΂��܂�"
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�N��</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
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

			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�_�����</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'���i
		sTemp = ""
		If LicenseGroupCode1 & LicenseCategoryCode1 & LicenseName1 <> "" Then
			sTemp = LicenseName1
			If sTemp = "" And LicenseCategoryName1 <> "" Then sTemp = LicenseCategoryName1 & "�֘A"
			If sTemp = "" And LicenseGroupName1 <> "" Then sTemp = LicenseGroupName1
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">���i</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'�n�r
		sTemp = ""
		If OSName1 <> "" Then
			sTemp = sTemp & OSName1
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�n�r</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'�A�v���P�[�V����
		sTemp = ""
		If ApplicationName1 <> "" Then
			sTemp = sTemp & ApplicationName1
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�A�v���P�[�V����</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'�J������
		sTemp = ""
		If DevelopmentLanguageName1 <> "" Then
			sTemp = sTemp & DevelopmentLanguageName1
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�J������</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'�f�[�^�x�[�X
		sTemp = ""
		If DatabaseName1 <> "" Then
			sTemp = sTemp & DatabaseName1
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�f�[�^�x�[�X</th><td style=""width:439px;"">" & sTemp & "</td></tr>"
		End If

		'�L�[���[�h
		sTemp = ""
		If Keyword <> "" Then
			GetHtmlSearchCondition = GetHtmlSearchCondition & "<tr><th style=""width:138px;"">�L�[���[�h</th><td style=""width:439px;"">" & Keyword & "</td></tr>"
		End If

		'���R�[�h�i�����j
		If OrderCode <> "" Then
			GetHtmlSearchCondition = "<tr><th style=""width:138px;"">���R�[�h</th><td style=""width:439px;"">" & OrderCode & "</td></tr>"
		End If

		If GetHtmlSearchCondition <> "" Then
			GetHtmlSearchCondition = "<table class=""pattern1"" border=""0"" style=""width:600px;""><thead><tr><th colspan=""2"" style=""width:588px;"">��������</th></tr></thead><tbody>" & GetHtmlSearchCondition & "</tbody></table>"
		End If

	End Function
End Class
%>
