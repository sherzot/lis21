<%
'******************************************************************************
'�T�@�v�F���E�Ҍ���������ێ�����N���X
'�ց@���F��Private
'�@�@�@�FChkData					�F�f�[�^�̐��������`�F�b�N
'�@�@�@�FSetNames					�F�R�[�h�ɑΉ��������̂��擾����
'�@�@�@�F
'�@�@�@�F��Public
'�@�@�@�FClass_Initialize			�F�R���X�g���N�^
'�@�@�@�FDspConditionHidden			�F���d���ڍ׌����̏���hidden���o�͂���
'�@�@�@�FGetSearchParam				�F���d���ڍ׌����y�[�W�֓n��GET�p�����[�^�𐶐����Ď擾�B
'�@�@�@�FGetSQLStaffSearchDetail	�F���l�[�ڍ׌����r�p�k���擾
'�@�@�@�FGetSQLWriteLog				�F���l�[�ڍ׌����k�n�f�������݂r�p�k���擾
'�@�@�@�FGetHtmlSearchCondition		�F���l�[�ڍ׌��������o�͂g�s�l�k���擾
'�@�@�@�F
'���@�l�F������ �ڍ׌����p�p�����[�^ (�A�h�z�b�N�Ȃr�p�k����)
'�@�@�@�Fsdf	�F�ڍ׌����t���O ["1"]�ڍ׌���
'�@�@�@�Fordercode�F���E�҂��������鋁�l�[�̃R�[�h
'�@�@�@�Frdfrom	�F�o�^�������i���E�Ҏ��������Ŏg�p�j
'�@�@�@�Fswt	�F��]�Ζ��`�ԁi�J���}��؂�j
'�@�@�@�Fswt1	�F��]�Ζ��`�ԂP
'�@�@�@�Fswt2	�F��]�Ζ��`�ԂQ
'�@�@�@�Fshjt1	�F��]�E��P
'�@�@�@�Fshj2t	�F��]�E��Q
'�@�@�@�Fsjt1	�F�E��E��P
'�@�@�@�Fsjt2	�F�E��E��Q
'�@�@�@�Fsjp1	�F�E��E��P�̊���
'�@�@�@�Fsjp2	�F�E��E��Q�̊���
'�@�@�@�Fsccnt	�F���Љ�
'�@�@�@�Fshitc	�F��]�Ǝ�i�J���}��؂�j
'�@�@�@�Fseitc	�F�o���Ǝ�i�J���}��؂�j
'�@�@�@�Fshp	�F��]�Ζ��n�s���{���i�J���}��؂�j
'�@�@�@�Fshp1	�F��]�Ζ��n�s���{���P
'�@�@�@�Fshc1	�F��]�Ζ��n�s���{���P�̎s��S
'�@�@�@�Fshp2	�F��]�Ζ��n�s���{���Q
'�@�@�@�Fshc2	�F��]�Ζ��n�s���{���Q�̎s��S
'�@�@�@�Fsyimin	�F�N�����
'�@�@�@�Fsyimin	�F�N������
'�@�@�@�Fsmimin	�F�������
'�@�@�@�Fsmimin	�F��������
'�@�@�@�Fsdimin	�F�������
'�@�@�@�Fsdimin	�F��������
'�@�@�@�Fshimin	�F�������
'�@�@�@�Fshimin	�F��������
'�@�@�@�Fsp		�F�Z���s���{���i�J���}��؂�j
'�@�@�@�Fsc		�F�Z���s��S
'�@�@�@�Fsrlpc	�F�Z���Ŋ񉈐��p�s���{���R�[�h
'�@�@�@�Fsrlc	�F�Z���Ŋ񉈐��R�[�h
'�@�@�@�Fssc	�F�Z���Ŋ�w�R�[�h�i�J���}��؂�j
'�@�@�@�Fszpc	�F�Z���X�֔ԍ��p�s���{���R�[�h
'�@�@�@�Fszc	�F�Z���X�֔ԍ���S���i�J���}��؂�j
'�@�@�@�Fsamin	�F�N���
'�@�@�@�Fsamax	�F�N����
'�@�@�@�Fsex	�F���� [1]�j [2]��
'�@�@�@�Fsstc	�F�w���w�Z��ʃR�[�h�i�J���}��؂�j
'�@�@�@�Fssn	�F�w�Z��
'�@�@�@�Fsct	�F�w�𕶗����
'�@�@�@�Fslg1	�F�������i�P�啪��
'�@�@�@�Fslc1	�F�������i�P������
'�@�@�@�Fsl1	�F�������i�P������
'�@�@�@�Fslg2	�F�������i�Q�啪��
'�@�@�@�Fslc2	�F�������i�Q������
'�@�@�@�Fsl2	�F�������i�Q������
'�@�@�@�Fslg3	�F�������i�R�啪��
'�@�@�@�Fslc3	�F�������i�R������
'�@�@�@�Fsl3	�F�������i�R������
'�@�@�@�Fslng	�F��w�X�L��(����R�[�h)
'�@�@�@�Fslngal1�F��w�X�L��(��b���x��)
'�@�@�@�Fslngal2�F��w�X�L��(�ǉ����x��)
'�@�@�@�Fslngal3�F��w�X�L��(�앶���x��)
'�@�@�@�Fssao	�F�X�L����AND,OR�����t���O ["AND"]AND���� ["OR"]OR����
'�@�@�@�Fsos1	�F�X�L���n�r�P
'�@�@�@�Fsos2	�F�X�L���n�r�Q
'�@�@�@�Fsosp1	�F�X�L���n�r�P�g�p�N��
'�@�@�@�Fsosp2	�F�X�L���n�r�Q�g�p�N��
'�@�@�@�Fsoa1	�F�X�L���n�`�P(�p�~)
'�@�@�@�Fsap1	�F�X�L���A�v���P�[�V�����P
'�@�@�@�Fsap2	�F�X�L���A�v���P�[�V�����Q
'�@�@�@�Fsap3	�F�X�L���A�v���P�[�V�����R
'�@�@�@�Fsapp1	�F�X�L���A�v���P�[�V�����P�g�p�N��
'�@�@�@�Fsapp2	�F�X�L���A�v���P�[�V�����Q�g�p�N��
'�@�@�@�Fsapp3	�F�X�L���A�v���P�[�V�����R�g�p�N��
'�@�@�@�Fsdl1	�F�X�L���J������P
'�@�@�@�Fsdl2	�F�X�L���J������Q
'�@�@�@�Fsdlp1	�F�X�L���J������P�g�p�N��
'�@�@�@�Fsdlp2	�F�X�L���J������Q�g�p�N��
'�@�@�@�Fsdb1	�F�X�L���f�[�^�x�[�X�P
'�@�@�@�Fsdb2	�F�X�L���f�[�^�x�[�X�Q
'�@�@�@�Fsdbp1	�F�X�L���f�[�^�x�[�X�P�g�p�N��
'�@�@�@�Fsdbp2	�F�X�L���f�[�^�x�[�X�Q�g�p�N��
'�@�@�@�Fsitsao	�F�J���c�[���g�p���т�AND,OR�����t���O ["AND"]AND���� ["OR"]OR����
'�@�@�@�Fsitos1	�F�J���c�[���g�p���т̂n�r�P
'�@�@�@�Fsitos2	�F�J���c�[���g�p���т̂n�r�Q
'�@�@�@�Fsitap1	�F�J���c�[���g�p���т̃A�v���P�[�V�����P
'�@�@�@�Fsitap2	�F�J���c�[���g�p���т̃A�v���P�[�V�����Q
'�@�@�@�Fsitap3	�F�J���c�[���g�p���т̃A�v���P�[�V�����R
'�@�@�@�Fsitdl1	�F�J���c�[���g�p���т̊J������P
'�@�@�@�Fsitdl2	�F�J���c�[���g�p���т̊J������Q
'�@�@�@�Fsitdb1	�F�J���c�[���g�p���т̃f�[�^�x�[�X�P
'�@�@�@�Fsitdb2	�F�J���c�[���g�p���т̃f�[�^�x�[�X�Q
'�@�@�@�Fssprf	�F�L�[���[�h�������Ȃo�q�t���O
'�@�@�@�Fsbdf	�F�L�[���[�h�����E�����e�t���O
'�@�@�@�Fsdrf	�F�L�[���[�h�����J���c�[���t���O
'�@�@�@�Fsddf	�F�L�[���[�h�����J�����e�t���O
'�@�@�@�Fsolf	�F�L�[���[�h�������̑����i�t���O
'�@�@�@�Fsosf	�F�L�[���[�h�������̑��X�L���t���O
'�@�@�@�Fskw	�F�L�[���[�h
'�@�@�@�Fsmlf	�F���E�҂̓����F���[���t���O
'�@�@�@�Fsstf	�F���E�҃R�[�h
'�@�@�@�Fsmstf	�F�}�b�`���O�Ώۋ��E�҃R�[�h
'���@���F2007/04/05 LIS K.Kokubo �쐬
'�@�@�@�F2007/07/03 LIS K.Kokubo �����r�p�k�� WHERE���EXISTS �� INNER JOIN �ɕύX�B
'�@�@�@�F2007/07/06 LIS K.Kokubo ��]�E��ƌo���E�킪���������ŕK��%�����Ă������̂��Y�o����%�ŕ�����悤�ɂ����B
'�@�@�@�F2007/11/05 LIS K.Kokubo �s�n�o�����Ŏ擾�\�B
'�@�@�@�F2008/01/15 LIS K.Kokubo ���l�[�ڍ׌��������o�͂g�s�l�k�Ɏ����������̏������o�́B
'�@�@�@�F2009/03/26 LIS K.Kokubo �i�r������,�l�މ�ЊJ���ɔ����ACompanyKbn�����o�ϐ��̍l������ύX�B�V�����l����...[1]��ʋ��l�L��[2]�l�ޏЉ�Č�[4]�h���Č��B���l����...[1]��ʊ��[2]�l�މ��[4]�h����ЁB
'�@�@�@�F2009/07/17 LIS K.Kokubo �Z���s��S�ǉ�
'�@�@�@�F2009/08/13 LIS K.Kokubo �}�b�`���O�����p�̌������ڐF�X�ǉ�
'�@�@�@�F2009/08/19 LIS K.Kokubo �K�ޑҋ@�v�����o�b�`�p�����ǉ�
'�@�@�@�F2009/09/02 LIS K.Kokubo �o���E��̔N�������ŐE�했�Ɍo���N�������Z�������̂���������悤�ɂ���
'�@�@�@�F2010/11/02 LIS K.Kokubo �L�[���[�h�������u��]�v�u�o���v�u���i�E��w�v�ɕ�����
'�@�@�@�F2011/05/15 LIS K.Kokubo �T�[�o���v���[�X�ɂ��AMAXDOP 1 �̎w�����������
'�@�@�@�F2012/03/07 LIS K.Kokubo ���ƔN����
'******************************************************************************
Class clsSearchStaffCondition
	Public CompanyCode
	Public CompanyKbn					'���l�[�敪 [1]��ʋ��l�L�� [2]�l�މ�Ђ̏Љ�Č� [4]�l�މ�Ђ̔h���Č�
	Public UserType
	Public SearchDetailFlag
	Public SetDataFlag
	Public SpMchNoticeFlag				'�K�ޑҋ@�v�����t���O [1]�K�ޑҋ@�ʒm���[��

	'�������������o�ϐ�
	Public Top							'SELECT�Ŏ擾���錏�� (SELECT TOP �� * FROM �`)
	Public OrderCode					'���E�҂��������鋁�l�[�̃R�[�h
	Public RegistDayFrom				'�o�^�������i���E�Ҏ��������Ŏg�p�j
	Public HopeWorkingTypeCode			'��]�ٗp�`�ԁi�J���}��؂�j
	Public WorkingTypeCode1				'�ٗp�`�ԂP
	Public WorkingTypeCode2				'�ٗp�`�ԂQ
	Public HopeJobTypeCode1				'��]�E��P
	Public HopeJobTypeCode2				'��]�E��Q
	Public JobTypeCode1					'�o���E��P
	Public JobTypeCode2					'�o���E��Q
	Public JobPeriod1					'�o���E��N���P
	Public JobPeriod2					'�o���E��N���Q
	Public CareerCnt					'���Љ�
	Public HopeIndustryTypeCode			'��]�Ǝ�i�J���}��؂�j
	Public ExpIndustryTypeCode			'�o���Ǝ�
	Public HopePrefectureCode			'��]�s���{���i�J���}��؂�j
	Public HopePrefectureCode1			'��]�s���{���P
	Public HopeCity1					'��]�s��S�P
	Public HopePrefectureCode2			'��]�s���{���Q
	Public HopeCity2					'��]�s��S�Q
	Public YearlyIncomeMin				'��]�N������
	Public YearlyIncomeMax				'��]�N�����
	Public MonthlyIncomeMin				'��]��������
	Public MonthlyIncomeMax				'��]�������
	Public DailyIncomeMin				'��]��������
	Public DailyIncomeMax				'��]�������
	Public HourlyIncomeMin				'��]��������
	Public HourlyIncomeMax				'��]�������
	Public PrefectureCode				'�Z���s���{���i�J���}��؂�j
	Public City							'�Z���s��S
	Public RailwayLinePrefectureCode	'�Z���Ŋ�w�����p�s���{���R�[�h
	Public RailwayLineCode				'�Z���Ŋ�w�����R�[�h
	Public StationCode					'�Z���Ŋ�w�R�[�h�i�J���}��؂�j
	Public ZipPrefectureCode			'�Z���X�֔ԍ��p�s���{���R�[�h
	Public ZipCode						'�Z���X�֔ԍ���S��
	Public AgeMin						'�ŏ��N��
	Public AgeMax						'�ő�N��
	Public Sex							'����
	Public SchoolTypeCode				'�o���w��
	Public SchoolName					'���Ƒ�w
	Public CourseType					'�w�𕶗����
	Public FinSchoolTypeCode			'�ŏI�w���w�Z�R�[�h
	Public GraduateYearMin				'�ŏ����ƔN
	Public GraduateYearMax				'�ő呲�ƔN
	Public LicenseGroupCode1			'���i�啪�ނP
	Public LicenseCategoryCode1			'���i�����ނP
	Public LicenseCode1					'���i�����ނP
	Public LicenseGroupCode2			'���i�啪�ނQ
	Public LicenseCategoryCode2			'���i�����ނQ
	Public LicenseCode2					'���i�����ނQ
	Public LicenseGroupCode3			'���i�啪�ނR
	Public LicenseCategoryCode3			'���i�����ނR
	Public LicenseCode3					'���i�����ނR
	Public LanguageCode					'����R�[�h
	Public LanguageActionLevel1			'�����b���x��
	Public LanguageActionLevel2			'����ǉ����x��
	Public LanguageActionLevel3			'����앶���x��
	Public SkillAndOr					'�X�L������ AND OR
	Public OSCode1						'�n�r�P
	Public OSCode2						'�n�r�Q
	Public OSPeriod1					'�n�r�P�g�p�N��
	Public OSPeriod2					'�n�r�Q�g�p�N��
	Public OACode1
	Public ApplicationCode1				'�A�v���P�[�V�����P
	Public ApplicationCode2				'�A�v���P�[�V�����Q
	Public ApplicationCode3				'�A�v���P�[�V�����R
	Public ApplicationPeriod1			'�A�v���P�[�V�����P�g�p�N��
	Public ApplicationPeriod2			'�A�v���P�[�V�����Q�g�p�N��
	Public ApplicationPeriod3			'�A�v���P�[�V�����R�g�p�N��
	Public DevelopmentLanguageCode1		'�J������P
	Public DevelopmentLanguageCode2		'�J������Q
	Public DevelopmentLanguagePeriod1	'�J������P�g�p�N��
	Public DevelopmentLanguagePeriod2	'�J������Q�g�p�N��
	Public DatabaseCode1				'�f�[�^�x�[�X�P
	Public DatabaseCode2				'�f�[�^�x�[�X�Q
	Public DatabasePeriod1				'�f�[�^�x�[�X�P�g�p�N��
	Public DatabasePeriod2				'�f�[�^�x�[�X�Q�g�p�N��
	Public ITSkillAndOr					'�X�L������ AND OR
	Public ITOSCode1					'�h�s�n�r�P
	Public ITOSCode2					'�h�s�n�r�Q
	Public ITApplicationCode1			'�h�s�A�v���P�[�V�����P
	Public ITApplicationCode2			'�h�s�A�v���P�[�V�����Q
	Public ITApplicationCode3			'�h�s�A�v���P�[�V�����R
	Public ITDevelopmentLanguageCode1	'�h�s�J������P
	Public ITDevelopmentLanguageCode2	'�h�s�J������Q
	Public ITDatabaseCode1				'�h�s�f�[�^�x�[�X�P
	Public ITDatabaseCode2				'�h�s�f�[�^�x�[�X�Q
	Public KeyWord						'�t���[���[�h
	Public KeyWordFlag					'�t���[���[�h ["1"]�n�q���� ["2"]�`�m�c����
	Public KeyWordHope					'�t���[���[�h(��])
	Public KeyWordHopeFlag				'�t���[���[�h(��]) ["1"]�n�q���� ["2"]�`�m�c����
	Public KeyWordCareer				'�t���[���[�h(�o��)
	Public KeyWordCareerFlag			'�t���[���[�h(�o��) ["1"]�n�q���� ["2"]�`�m�c����
	Public KeyWordLicense				'�t���[���[�h(���i�E��w)
	Public KeyWordLicenseFlag			'�t���[���[�h(���i�E��w) ["1"]�n�q���� ["2"]�`�m�c����
	Public KeyWordPerson				'�t���[���[�h(�l�f�[�^)
	Public KeyWordPersonFlag			'�t���[���[�h(�l�f�[�^) ["1"]�n�q���� ["2"]�`�m�c����

	Public MailFlag				'���[������M�������̂��鋁�E�҂݂̂������t���O
	Public StaffCode					'���E�҃R�[�h�i�����j
	Public MchStaffCode					'�}�b�`���O�Ώۋ��E�҃R�[�h

	'�R�[�h�Ή�����
	Public HopeWorkingTypeName			'��]�ٗp�`��
	Public WorkingTypeName1				'�ٗp�`�ԂP
	Public WorkingTypeName2				'�ٗp�`�ԂQ
	Public HopeJobTypeName1				'��]�E��P
	Public HopeJobTypeName2				'��]�E��Q
	Public JobTypeName1					'�o���E��P
	Public JobTypeName2					'�o���E��Q
	Public HopeIndustryTypeName			'��]�Ǝ�
	Public ExpIndustryTypeName			'�o���Ǝ�
	Public HopePrefectureName			'��]�s���{���J���}��؂�
	Public HopePrefectureName1			'��]�s���{���P
	Public HopePrefectureName2			'��]�s���{���Q
	Public PrefectureName				'�Z���s���{��
	Public RailwayLinePrefectureName	'�Z���Ŋ�w�����p�s���{����
	Public RailwayLineName				'�Z���Ŋ񉈐���
	Public StationName					'�Z���Ŋ�w��
	Public ZipName						'�Z���X�֔ԍ���S��
	Public SchoolTypeName				'���Ɗw�Z��ʖ�
	Public FinSchoolTypeName			'�ŏI�w���w�Z��ʖ�
	Public LicenseName1					'���i�P
	Public LicenseName2					'���i�Q
	Public LicenseName3					'���i�R
	Public LanguageName					'���ꖼ
	Public LanguageActionLevelName1		'��b���x����
	Public LanguageActionLevelName2		'�ǉ����x����
	Public LanguageActionLevelName3		'�앶���x����
	Public OSName1						'�n�r�P
	Public OSName2						'�n�r�Q
	Public ApplicationName1				'�A�v���P�[�V�����P
	Public ApplicationName2				'�A�v���P�[�V�����Q
	Public ApplicationName3				'�A�v���P�[�V�����R
	Public DevelopmentLanguageName1		'�J������P
	Public DevelopmentLanguageName2		'�J������Q
	Public DatabaseName1				'�f�[�^�x�[�X�P
	Public DatabaseName2				'�f�[�^�x�[�X�Q
	Public ITOSName1					'�h�s�n�r�P
	Public ITOSName2					'�h�s�n�r�Q
	Public ITApplicationName1			'�h�s�A�v���P�[�V�����P
	Public ITApplicationName2			'�h�s�A�v���P�[�V�����Q
	Public ITApplicationName3			'�h�s�A�v���P�[�V�����R
	Public ITDevelopmentLanguageName1	'�h�s�J������P
	Public ITDevelopmentLanguageName2	'�h�s�J������Q
	Public ITDatabaseName1				'�h�s�f�[�^�x�[�X�P
	Public ITDatabaseName2				'�h�s�f�[�^�x�[�X�Q

	'���̑������o�ϐ�
	Public HtmlStaffSearch	'���������o�͂g�s�l�k��
	Public SQLStaffSearch	'�����r�p�k
	Public SQLWriteLog		'���O�������݂r�p�k

	'******************************************************************************
	'�T�@�v�F�R���X�g���N�^
	'�쐬�ҁFLis K.Kokubo
	'�쐬���F2007/04/04 Lis K.Kokubo
	'�X�@�V�F
	'���@�l�F
	'******************************************************************************
	Private Sub Class_Initialize()
		CompanyCode = Session("userid")
		'2009/03/26 LIS K.Kokubo DEL
		'CompanyKbn = Session("companykbn")
		UserType = Session("usertype")

		'�p�����[�^���猟���������擾
		Call SetData_ParamPart("setdata", "")
		Call SetData_ParamPart("sdf", "")
		Call SetData_ParamPart("ordercode", "")
		Call SetData_ParamPart("rdfrom", "")
		Call SetData_ParamPart("swt", "")
		Call SetData_ParamPart("swt1", "")
		Call SetData_ParamPart("swt2", "")
		Call SetData_ParamPart("shjt1", "")
		Call SetData_ParamPart("shjt2", "")
		Call SetData_ParamPart("sjt1", "")
		Call SetData_ParamPart("sjt2", "")
		Call SetData_ParamPart("sjp1", "")
		Call SetData_ParamPart("sjp2", "")
		Call SetData_ParamPart("sccnt", "")
		Call SetData_ParamPart("shitc", "")
		Call SetData_ParamPart("seitc", "")
		Call SetData_ParamPart("shp", "")
		Call SetData_ParamPart("shp1", "")
		Call SetData_ParamPart("shc1", "")
		Call SetData_ParamPart("shp2", "")
		Call SetData_ParamPart("shc2", "")
		Call SetData_ParamPart("syimin", "")
		Call SetData_ParamPart("syimax", "")
		Call SetData_ParamPart("smimin", "")
		Call SetData_ParamPart("smimax", "")
		Call SetData_ParamPart("sdimin", "")
		Call SetData_ParamPart("sdimax", "")
		Call SetData_ParamPart("shimin", "")
		Call SetData_ParamPart("shimax", "")
		Call SetData_ParamPart("sp", "")
		Call SetData_ParamPart("sc", "")
		Call SetData_ParamPart("srlpc","")
		Call SetData_ParamPart("srlc", "")
		Call SetData_ParamPart("ssc", "")
		Call SetData_ParamPart("szpc", "")
		Call SetData_ParamPart("szc", "")
		Call SetData_ParamPart("samin", "")
		Call SetData_ParamPart("samax", "")
		Call SetData_ParamPart("ssex", "")
		Call SetData_ParamPart("sstc", "")
		Call SetData_ParamPart("ssn", "")
		Call SetData_ParamPart("sct", "")
		Call SetData_ParamPart("sfstc", "")
		Call SetData_ParamPart("sgymin", "")
		Call SetData_ParamPart("sgymax", "")
		Call SetData_ParamPart("slg1", "")
		Call SetData_ParamPart("slc1", "")
		Call SetData_ParamPart("sl1", "")
		Call SetData_ParamPart("slg2", "")
		Call SetData_ParamPart("slc2", "")
		Call SetData_ParamPart("sl2", "")
		Call SetData_ParamPart("slg3", "")
		Call SetData_ParamPart("slc3", "")
		Call SetData_ParamPart("sl3", "")
		Call SetData_ParamPart("slng", "")
		Call SetData_ParamPart("slngal1", "")
		Call SetData_ParamPart("slngal2", "")
		Call SetData_ParamPart("slngal3", "")
		Call SetData_ParamPart("ssao", "")
		Call SetData_ParamPart("sos1", "")
		Call SetData_ParamPart("sos2", "")
		Call SetData_ParamPart("sosp1", "")
		Call SetData_ParamPart("sosp2", "")
		Call SetData_ParamPart("soa1", "")
		Call SetData_ParamPart("sap1", "")
		Call SetData_ParamPart("sap2", "")
		Call SetData_ParamPart("sap3", "")
		Call SetData_ParamPart("sapp1", "")
		Call SetData_ParamPart("sapp2", "")
		Call SetData_ParamPart("sapp3", "")
		Call SetData_ParamPart("sdl1", "")
		Call SetData_ParamPart("sdl2", "")
		Call SetData_ParamPart("sdlp1", "")
		Call SetData_ParamPart("sdlp2", "")
		Call SetData_ParamPart("sdb1", "")
		Call SetData_ParamPart("sdb2", "")
		Call SetData_ParamPart("sdbp1", "")
		Call SetData_ParamPart("sdbp2", "")
		Call SetData_ParamPart("sitsao", "")
		Call SetData_ParamPart("sitos1", "")
		Call SetData_ParamPart("sitos2", "")
		Call SetData_ParamPart("sitap1", "")
		Call SetData_ParamPart("sitap2", "")
		Call SetData_ParamPart("sitap3", "")
		Call SetData_ParamPart("sitdl1", "")
		Call SetData_ParamPart("sitdl2", "")
		Call SetData_ParamPart("sitdb1", "")
		Call SetData_ParamPart("sitdb2", "")
		Call SetData_ParamPart("skw", "")
		Call SetData_ParamPart("skwf", "")
		Call SetData_ParamPart("skwh", "")
		Call SetData_ParamPart("skwhf", "")
		Call SetData_ParamPart("skwc", "")
		Call SetData_ParamPart("skwcf", "")
		Call SetData_ParamPart("skwl", "")
		Call SetData_ParamPart("skwlf", "")
		Call SetData_ParamPart("skwp", "")
		Call SetData_ParamPart("skwpf", "")
		Call SetData_ParamPart("smlf", "")
		Call SetData_ParamPart("sstf", "")
		Call SetData_ParamPart("smstf", "")

		'�f�t�H���g�̌��������ݒ�
		If SetDataFlag = "1" Then
			Call SetDefaultCondition()
		End If

		'�f�[�^�̐������`�F�b�N
		Call ChkData()

		'�R�[�h�Ή����̎擾
		Call SetNames()

		'���l�[����SQL����
		SQLStaffSearch = GetSQLStaffSearchDetail()

		'���O��������SQL����
		SQLWriteLog = GetSQLWriteLog()

		'���l�[���������o�͂g�s�l�k��
		HtmlStaffSearch = GetHtmlSearchCondition()
	End Sub

	'******************************************************************************
	'�T�@�v�F�p�����[�^���ƃ����o�ϐ���R�t���Ēl��ݒ肷��
	'���@���FvKey	�F
	'�@�@�@�FvValue	�F
	'�@�@�@�FvFlag	�F
	'���@�l�F
	'�X�@�V�F2009/08/06 LIS K.Kokubo
	'******************************************************************************
	Private Sub SetData_ParamPart(ByVal vKey, ByVal vValue)
		If Len(vValue) = 0 Then vValue = GetForm(vKey, 2)

		Select Case vKey
			Case "setdata": SetDataFlag = vValue
			Case "sdf": SearchDetailFlag = vValue
			Case "ordercode": OrderCode = vValue
			Case "rdfrom": RegistDayFrom = vValue
			Case "swt": HopeWorkingTypeCode = vValue
			Case "swt1": WorkingTypeCode1 = vValue
			Case "swt2": WorkingTypeCode2 = vValue
			Case "shjt1": HopeJobTypeCode1 = vValue
			Case "shjt2": HopeJobTypeCode2 = vValue
			Case "sjt1": JobTypeCode1 = vValue
			Case "sjt2": JobTypeCode2 = vValue
			Case "sjp1": JobPeriod1 = vValue
			Case "sjp2": JobPeriod2 = vValue
			Case "sccnt": CareerCnt = vValue
			Case "shitc": HopeIndustryTypeCode = vValue
			Case "seitc": ExpIndustryTypeCode = vValue
			Case "shp": HopePrefectureCode = vValue
			Case "shp1": HopePrefectureCode1 = vValue
			Case "shc1": HopeCity1 = vValue
			Case "shp2": HopePrefectureCode2 = vValue
			Case "shc2": HopeCity2 = vValue
			Case "syimin": YearlyIncomeMin = vValue
			Case "syimax": YearlyIncomeMax = vValue
			Case "smimin": MonthlyIncomeMin = vValue
			Case "smimax": MonthlyIncomeMax = vValue
			Case "sdimin": DailyIncomeMin = vValue
			Case "sdimax": DailyIncomeMax = vValue
			Case "shimin": HourlyIncomeMin = vValue
			Case "shimax": HourlyIncomeMax = vValue
			Case "sp": PrefectureCode = vValue
			Case "sc": City = vValue
			Case "srlpc": RailwayLinePrefectureCode = vValue
			Case "srlc": RailwayLineCode = vValue
			Case "ssc": StationCode = vValue
			Case "szpc": ZipPrefectureCode = vValue
			Case "szc": ZipCode = vValue
			Case "samin": AgeMin = vValue
			Case "samax": AgeMax = vValue
			Case "ssex": Sex = vValue
			Case "sstc": SchoolTypeCode = vValue
			Case "ssn": SchoolName = vValue
			Case "sct": CourseType = vValue
			Case "sfstc": FinSchoolTypeCode = vValue
			Case "sgymin": GraduateYearMin = vValue
			Case "sgymax": GraduateYearMax = vValue
			Case "slg1": LicenseGroupCode1 = vValue
			Case "slc1": LicenseCategoryCode1 = vValue
			Case "sl1": LicenseCode1 = vValue
			Case "slg2": LicenseGroupCode2 = vValue
			Case "slc2": LicenseCategoryCode2 = vValue
			Case "sl2": LicenseCode2 = vValue
			Case "slg3": LicenseGroupCode3 = vValue
			Case "slc3": LicenseCategoryCode3 = vValue
			Case "sl3": LicenseCode3 = vValue
			Case "slng": LanguageCode = vValue
			Case "slngal1": LanguageActionLevel1 = vValue
			Case "slngal2": LanguageActionLevel2 = vValue
			Case "slngal3": LanguageActionLevel3 = vValue
			Case "ssao": SkillAndOr = vValue
			Case "sos1": OSCode1 = vValue
			Case "sos2": OSCode2 = vValue
			Case "sosp1": OSPeriod1 = vValue
			Case "sosp2": OSPeriod2 = vValue
			Case "soa1": OACode1 = vValue
			Case "sap1": ApplicationCode1 = vValue
			Case "sap2": ApplicationCode2 = vValue
			Case "sap3": ApplicationCode3 = vValue
			Case "sapp1": ApplicationPeriod1 = vValue
			Case "sapp2": ApplicationPeriod2 = vValue
			Case "sapp3": ApplicationPeriod3 = vValue
			Case "sdl1": DevelopmentLanguageCode1 = vValue
			Case "sdl2": DevelopmentLanguageCode2 = vValue
			Case "sdlp1": DevelopmentLanguagePeriod1 = vValue
			Case "sdlp2": DevelopmentLanguagePeriod2 = vValue
			Case "sdb1": DatabaseCode1 = vValue
			Case "sdb2": DatabaseCode2 = vValue
			Case "sdbp1": DatabasePeriod1 = vValue
			Case "sdbp2": DatabasePeriod2 = vValue
			Case "sitsao": ITSkillAndOr = vValue
			Case "sitos1": ITOSCode1 = vValue
			Case "sitos2": ITOSCode2 = vValue
			Case "sitap1": ITApplicationCode1 = vValue
			Case "sitap2": ITApplicationCode2 = vValue
			Case "sitap3": ITApplicationCode3 = vValue
			Case "sitdl1": ITDevelopmentLanguageCode1 = vValue
			Case "sitdl2": ITDevelopmentLanguageCode2 = vValue
			Case "sitdb1": ITDatabaseCode1 = vValue
			Case "sitdb2": ITDatabaseCode2 = vValue
			Case "skw": KeyWord = vValue
			Case "skwf": KeyWordFlag = vValue
			Case "skwh": KeyWordHope = vValue
			Case "skwhf": KeyWordHopeFlag = vValue
			Case "skwc": KeyWordCareer = vValue
			Case "skwcf": KeyWordCareerFlag = vValue
			Case "skwl": KeyWordLicense = vValue
			Case "skwlf": KeyWordLicenseFlag = vValue
			Case "skwp": KeyWordPerson = vValue
			Case "skwpf": KeyWordPersonFlag = vValue

			Case "smlf": MailFlag = vValue
			Case "sstf": StaffCode = vValue
			Case "smstf": MchStaffCode = vValue
		End Select
	End Sub

	'******************************************************************************
	'�T�@�v�FvOrderCode�̋��l�[�Ō���������ݒ肷��
	'���@�l�F
	'�X�@�V�F2007/12/04 LIS K.Kokubo
	'******************************************************************************
	Private Sub SetDefaultCondition()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sError

		If OrderCode = "" Then Exit Sub

		'��]�s���{��
		PrefectureCode = ""
		sSQL = "EXEC up_LstC_WorkingPlace '" & OrderCode & "';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			Set oRS.ActiveConnection = Nothing
			oRS.Filter = "WorkingPlaceSeq = 1"
			If GetRSState(oRS) = True Then HopePrefectureCode1 = ChkStr(oRS.Collect("WorkingPlacePrefectureCode"))
			oRS.Filter = 0
			oRS.Filter = "WorkingPlaceSeq = 2"
			If GetRSState(oRS) = True Then HopePrefectureCode2 = ChkStr(oRS.Collect("WorkingPlacePrefectureCode"))
		End If
		Call RSClose(oRS)

		'��]�E��
		sSQL = "SELECT CJT.JobTypeCode FROM C_JobType AS CJT WITH(NOLOCK) WHERE CJT.OrderCode = '" & OrderCode & "' AND CJT.ID = 1"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			HopeJobTypeCode1 = ChkStr(oRS.Collect("JobTypeCode"))
		End If
		Call RSClose(oRS)
		sSQL = "SELECT CJT.JobTypeCode FROM C_JobType AS CJT WITH(NOLOCK) WHERE CJT.OrderCode = '" & OrderCode & "' AND CJT.ID = 2"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			HopeJobTypeCode2 = ChkStr(oRS.Collect("JobTypeCode"))
		End If
		Call RSClose(oRS)

		'��]�Ζ��`��
		sSQL = "SELECT CWT.WorkingTypeCode FROM C_WorkingType AS CWT WITH(NOLOCK) WHERE CWT.OrderCode = '" & OrderCode & "' AND CWT.ID = 1"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			WorkingTypeCode1 = ChkStr(oRS.Collect("WorkingTypeCode"))
		End If
		Call RSClose(oRS)
	End Sub

	'******************************************************************************
	'�T�@�v�F�f�[�^�̐��������`�F�b�N
	'�쐬�ҁFLis K.Kokubo
	'�쐬���F2007/04/17 Lis K.Kokubo
	'���@���F
	'���@�l�F
	'******************************************************************************
	Private Sub ChkData()
		Dim tmp

		If UserType = "dispatch" Then
			'�h����Ƃ̏ꍇ�́u�h���v�u�Љ�\��h���v�̂�
			If WorkingTypeCode1 <> "001" And WorkingTypeCode1 <> "004" Then WorkingTypeCode1 = ""
			If WorkingTypeCode2 <> "001" And WorkingTypeCode2 <> "004" Then WorkingTypeCode2 = ""
		End If

		If UserType = "company" Then
			'��ʊ�ƁE�l�ޏЉ��Ƃ̏ꍇ�́u�h���v�u�Љ�\��h���v������
			If WorkingTypeCode1 = "001" And WorkingTypeCode1 = "004" Then WorkingTypeCode1 = ""
			If WorkingTypeCode2 = "001" And WorkingTypeCode2 = "004" Then WorkingTypeCode2 = ""
		End If

		'<���ƔN�̓��̓`�F�b�N>
		If CStr(GraduateYearMin) <> "" Then
			If IsNumber(GraduateYearMin,4,False) = True Then
				If CInt(GraduateYearMin) < 1900 Or CInt(GraduateYearMin) >= 2099 Then
					GraduateYearMin = ""
				End If
			Else
				GraduateYearMin = ""
			End If
		End If

		If CStr(GraduateYearMax) <> "" Then
			If IsNumber(GraduateYearMax,4,False) = True Then
				If CInt(GraduateYearMax) < 1900 Or CInt(GraduateYearMax) >= 2099 Then
					GraduateYearMax = ""
				End If
			Else
				GraduateYearMax = ""
			End If
		End If

		If CStr(GraduateYearMin) <> "" And CStr(GraduateYearMax) <> "" Then
			If CInt(GraduateYearMin) > CInt(GraduateYearMax) Then
				tmp = GraduateYearMin
				GraduateYearMin = GraduateYearMax
				GraduateYearMax = tmp
			End If
		End If
		'</���ƔN�̓��̓`�F�b�N>

		If KeyWordFlag = "" Then KeyWordFlag = "2"
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

		Dim idx
		Dim aValue
		Dim sXML

		'�ٗp�`�ԁi�J���}��؂�j
		If HopeWorkingTypeCode <> "" Then
			HopeWorkingTypeCode = Replace(HopeWorkingTypeCode, " ", "")
			aValue = Split(HopeWorkingTypeCode, ",")
			For idx = 0 To UBound(aValue)
				If HopeWorkingTypeName <> "" Then HopeWorkingTypeName = HopeWorkingTypeName & ",&nbsp;"
				HopeWorkingTypeName = HopeWorkingTypeName & GetDetail("WorkingType", aValue(idx))
			Next
		End If

		'�ٗp�`�ԂP
		If WorkingTypeCode1 <> "" Then
			WorkingTypeName1 = GetDetail("WorkingType", WorkingTypeCode1)
		End If

		'�ٗp�`�ԂQ
		If WorkingTypeCode2 <> "" Then
			WorkingTypeName2 = GetDetail("WorkingType", WorkingTypeCode2)
		End If

		'��]�E��P
		If IsRE(HopeJobTypeCode1, "^\d\d$", True) = True Then
			sSQL = "sp_GetListJobTypeBig '" & HopeJobTypeCode1 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				HopeJobTypeName1 = ChkStr(oRS.Collect("BigClassName"))
			End If
			Call RSClose(oRS)
		ElseIf IsRE(HopeJobTypeCode1, "^\d\d\d\d\d\d\d$", True) = True Then
			sSQL = "sp_GetListJobType '" & Left(HopeJobTypeCode1, 2) & "', '" & Mid(HopeJobTypeCode1, 3, 2) & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				HopeJobTypeName1 = ChkStr(oRS.Collect("MiddleClassName"))
			End If
			Call RSClose(oRS)
		End If
		'��]�E��Q
		If IsRE(HopeJobTypeCode2, "^\d\d$", True) = True Then
			sSQL = "sp_GetListJobTypeBig '" & HopeJobTypeCode2 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				HopeJobTypeName2 = ChkStr(oRS.Collect("BigClassName"))
			End If
			Call RSClose(oRS)
		ElseIf IsRE(HopeJobTypeCode2, "^\d\d\d\d\d\d\d$", True) = True Then
			sSQL = "sp_GetListJobType '" & Left(HopeJobTypeCode2, 2) & "', '" & Mid(HopeJobTypeCode2, 3, 2) & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				HopeJobTypeName2 = ChkStr(oRS.Collect("MiddleClassName"))
			End If
			Call RSClose(oRS)
		End If

		'�o���E��P
		If IsRE(JobTypeCode1, "^\d\d$", True) = True Then
			sSQL = "sp_GetListJobTypeBig '" & JobTypeCode1 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				JobTypeName1 = ChkStr(oRS.Collect("BigClassName"))
			End If
			Call RSClose(oRS)
		ElseIf IsRE(JobTypeCode1, "^\d\d\d\d\d\d\d$", True) = True Then
			sSQL = "sp_GetListJobType '" & Left(JobTypeCode1, 2) & "', '" & Mid(JobTypeCode1, 3, 2) & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				JobTypeName1 = ChkStr(oRS.Collect("MiddleClassName"))
			End If
			Call RSClose(oRS)
		End If
		'�o���E��Q
		If IsRE(JobTypeCode2, "^\d\d$", True) = True Then
			sSQL = "sp_GetListJobTypeBig '" & JobTypeCode2 & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				JobTypeName2 = ChkStr(oRS.Collect("BigClassName"))
			End If
			Call RSClose(oRS)
		ElseIf IsRE(JobTypeCode2, "^\d\d\d\d\d\d\d$", True) = True Then
			sSQL = "sp_GetListJobType '" & Left(JobTypeCode2, 2) & "', '" & Mid(JobTypeCode2, 3, 2) & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				JobTypeName2 = ChkStr(oRS.Collect("MiddleClassName"))
			End If
			Call RSClose(oRS)
		End If

		'��]�Ǝ�
		If HopeIndustryTypeCode <> "" Then
			HopeIndustryTypeCode = Replace(HopeIndustryTypeCode, " ", "")
			aValue = Split(HopeIndustryTypeCode, ",")
			For idx = 0 To UBound(aValue)
				If HopeIndustryTypeName <> "" Then HopeIndustryTypeName = HopeIndustryTypeName & ",&nbsp;"
				HopeIndustryTypeName = HopeIndustryTypeName & GetDetail("IndustryType", aValue(idx))
			Next
		End If

		'�o���Ǝ�
		If ExpIndustryTypeCode <> "" Then
			ExpIndustryTypeCode = Replace(ExpIndustryTypeCode, " ", "")
			aValue = Split(ExpIndustryTypeCode, ",")
			For idx = 0 To UBound(aValue)
				If ExpIndustryTypeName <> "" Then ExpIndustryTypeName = ExpIndustryTypeName & ",&nbsp;"
				ExpIndustryTypeName = ExpIndustryTypeName & GetDetail("IndustryType", aValue(idx))
			Next
		End If

		'��]�s���{���J���}��؂�
		If HopePrefectureCode <> "" Then
			HopePrefectureCode = Replace(HopePrefectureCode, " ", "")
			aValue = Split(HopePrefectureCode, ",")
			For idx = 0 To UBound(aValue)
				If HopePrefectureName <> "" Then HopePrefectureName = HopePrefectureName & ",&nbsp;"
				HopePrefectureName = HopePrefectureName & GetDetail("Prefecture", aValue(idx))
			Next
		End If

		'��]�s���{���P
		If HopePrefectureCode1 <> "" Then
			HopePrefectureName1 = GetDetail("Prefecture", HopePrefectureCode1)
		End If

		'��]�s���{���Q
		If HopePrefectureCode2 <> "" Then
			HopePrefectureName2 = GetDetail("Prefecture", HopePrefectureCode2)
		End If

		'�Z���s���{����
		If PrefectureCode <> "" Then
			PrefectureCode = Replace(PrefectureCode, " ", "")
			aValue = Split(PrefectureCode, ",")
			For idx = 0 To UBound(aValue)
				If PrefectureName <> "" Then PrefectureName = PrefectureName & ",&nbsp;"
				PrefectureName = PrefectureName & GetDetail("Prefecture", aValue(idx))
			Next
		End If

		'�Z���Ŋ�w�����p�s���{����
		If RailwayLinePrefectureCode <> "" Then
			RailwayLinePrefectureName = RailwayLinePrefectureName & GetDetail("Prefecture", RailwayLinePrefectureCode)
		End If

		'�Z���Ŋ񉈐���
		If RailwayLineCode <> "" Then
			sSQL = "EXEC up_DtlB_RailwayLine '" & RailwayLineCode & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				RailwayLineName = ChkStr(oRS.Collect("RailwayLineName"))
			End If
			Call RSClose(oRS)
		End If

		'�Z���Ŋ�w��
		If StationCode <> "" Then
			StationCode = Replace(StationCode, " ", "")
			aValue = Split(StationCode, ",")
			sXML = ""
			For idx = 0 To UBound(aValue)
				sXML = sXML & "<station><stationcode>" & aValue(idx) & "</stationcode></station>"
			Next
			sXML = "<root>" & sXML & "</root>"
			sSQL = "EXEC up_LstB_Station_XML '" & sXML & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Do While GetRSState(oRS) = True
				If StationName <> "" Then StationName = StationName & ","
				StationName = StationName & ChkStr(oRS.Collect("StationName"))

				oRS.MoveNext
			Loop
			Call RSClose(oRS)
		End If

		'�Z���X�֔ԍ���S��
		If ZipCode <> "" Then
			ZipCode = Replace(ZipCode, " ", "")
			aValue = Split(ZipCode, ",")
			For idx = 0 To UBound(aValue)
				If ZipName <> "" Then ZipName = ZipName & ",&nbsp;"
				ZipName = ZipName & Left(aValue(idx), 3) & "-" & Right(aValue(idx), 1) & "XXX"
			Next
		End If

		'���Ɗw���w�Z��ʖ�
		If SchoolTypeCode <> "" Then
			SchoolTypeCode = Replace(SchoolTypeCode, " ", "")
			aValue = Split(SchoolTypeCode, ",")
			For idx = 0 To UBound(aValue)
				If SchoolTypeName <> "" Then SchoolTypeName = SchoolTypeName & ",&nbsp;"
				SchoolTypeName = SchoolTypeName & GetDetail("SchoolType", aValue(idx))
			Next
		End If

		'�ŏI�w���w�Z��ʖ�
		If FinSchoolTypeCode <> "" Then
			FinSchoolTypeName = GetDetail("SchoolType", FinSchoolTypeCode)
		End If

		'���i�P
		If LicenseGroupCode1 <> "" And LicenseCategoryCode1 <> "" And LicenseCode1 <> "" Then
			sSQL = "SELECT LicenseName FROM vw_License WHERE GroupCode = '" & LicenseGroupCode1 & "' AND CategoryCode = '" & LicenseCategoryCode1 & "' AND Code = '" & Right(LicenseCode1, 2) & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then LicenseName1 = ChkStr(oRS.Collect("LicenseName"))
			Call RSClose(oRS)
		ElseIf LicenseGroupCode1 <> "" And LicenseCategoryCode1 <> "" Then
			sSQL = "SELECT LicenseCategoryName FROM vw_License WHERE GroupCode = '" & LicenseGroupCode1 & "' AND CategoryCode = '" & LicenseCategoryCode1 & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then LicenseName1 = ChkStr(oRS.Collect("LicenseCategoryName")) & "(������)"
			Call RSClose(oRS)
		ElseIf LicenseGroupCode1 <> "" Then
			sSQL = "SELECT LicenseGroupName FROM vw_License WHERE GroupCode = '" & LicenseGroupCode1 & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then LicenseName1 = ChkStr(oRS.Collect("LicenseGroupName")) & "(�啪��)"
			Call RSClose(oRS)
		End If

		'���i�Q
		If LicenseGroupCode2 <> "" And LicenseCategoryCode2 <> "" And LicenseCode2 <> "" Then
			sSQL = "SELECT LicenseName FROM vw_License WHERE GroupCode = '" & LicenseGroupCode2 & "' AND CategoryCode = '" & LicenseCategoryCode2 & "' AND Code = '" & Right(LicenseCode2, 2) & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then LicenseName2 = ChkStr(oRS.Collect("LicenseName"))
			Call RSClose(oRS)
		ElseIf LicenseGroupCode2 <> "" And LicenseCategoryCode2 <> "" Then
			sSQL = "SELECT LicenseCategoryName FROM vw_License WHERE GroupCode = '" & LicenseGroupCode2 & "' AND CategoryCode = '" & LicenseCategoryCode2 & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then LicenseName2 = ChkStr(oRS.Collect("LicenseCategoryName")) & "(������)"
			Call RSClose(oRS)
		ElseIf LicenseGroupCode2 <> "" Then
			sSQL = "SELECT LicenseGroupName FROM vw_License WHERE GroupCode = '" & LicenseGroupCode2 & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then LicenseName2 = ChkStr(oRS.Collect("LicenseGroupName")) & "(�啪��)"
			Call RSClose(oRS)
		End If

		'���i�R
		If LicenseGroupCode3 <> "" And LicenseCategoryCode3 <> "" And LicenseCode3 <> "" Then
			sSQL = "SELECT LicenseName FROM vw_License WHERE GroupCode = '" & LicenseGroupCode3 & "' AND CategoryCode = '" & LicenseCategoryCode3 & "' AND Code = '" & Right(LicenseCode3, 2) & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then LicenseName3 = ChkStr(oRS.Collect("LicenseName"))
			Call RSClose(oRS)
		ElseIf LicenseGroupCode3 <> "" And LicenseCategoryCode3 <> "" Then
			sSQL = "SELECT LicenseCategoryName FROM vw_License WHERE GroupCode = '" & LicenseGroupCode3 & "' AND CategoryCode = '" & LicenseCategoryCode3 & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then LicenseName3 = ChkStr(oRS.Collect("LicenseCategoryName")) & "(������)"
			Call RSClose(oRS)
		ElseIf LicenseGroupCode3 <> "" Then
			sSQL = "SELECT LicenseGroupName FROM vw_License WHERE GroupCode = '" & LicenseGroupCode3 & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then LicenseName3 = ChkStr(oRS.Collect("LicenseGroupName")) & "(�啪��)"
			Call RSClose(oRS)
		End If

		'����
		If LanguageCode <> "" Then
			LanguageName = GetLanguageName(LanguageCode)
		End If

		'��b���x��
		If LanguageActionLevel1 <> "" Then
			LanguageActionLevelName1 = GetLanguageActionLevelName("1", LanguageActionLevel1)
		End If

		'�ǉ����x��
		If LanguageActionLevel2 <> "" Then
			LanguageActionLevelName2 = GetLanguageActionLevelName("2", LanguageActionLevel2)
		End If

		'�앶���x��
		If LanguageActionLevel3 <> "" Then
			LanguageActionLevelName3 = GetLanguageActionLevelName("3", LanguageActionLevel3)
		End If

		'�n�r�P
		If OSCode1 <> "" Then
			OSName1 = GetDetail("OS", OSCode1)
		End If

		'�n�r�Q
		If OSCode2 <> "" Then
			OSName2 = GetDetail("OS", OSCode2)
		End If

		'�A�v���P�[�V�����P
		If ApplicationCode1 <> "" Then
			ApplicationName1 = GetDetail("Application", ApplicationCode1)
		End If

		'�A�v���P�[�V�����Q
		If ApplicationCode2 <> "" Then
			ApplicationName2 = GetDetail("Application", ApplicationCode2)
		End If

		'�A�v���P�[�V�����R
		If ApplicationCode3 <> "" Then
			ApplicationName3 = GetDetail("Application", ApplicationCode3)
		End If

		'�J������P
		If DevelopmentLanguageCode1 <> "" Then
			DevelopmentLanguageName1 = GetDetail("DevelopmentLanguage", DevelopmentLanguageCode1)
		End If

		'�J������Q
		If DevelopmentLanguageCode2 <> "" Then
			DevelopmentLanguageName2 = GetDetail("DevelopmentLanguage", DevelopmentLanguageCode2)
		End If

		'�f�[�^�x�[�X�P
		If DatabaseCode1 <> "" Then
			DatabaseName1 = GetDetail("Database", DatabaseCode1)
		End If

		'�f�[�^�x�[�X�Q
		If DatabaseCode2 <> "" Then
			DatabaseName2 = GetDetail("Database", DatabaseCode2)
		End If

		'�h�s�n�r�P
		If ITOSCode1 <> "" Then
			ITOSName1 = GetDetail("OS", ITOSCode1)
		End If

		'�h�s�n�r�Q
		If ITOSCode2 <> "" Then
			ITOSName2 = GetDetail("OS", ITOSCode2)
		End If

		'�h�s�A�v���P�[�V�����P
		If ITApplicationCode1 <> "" Then
			ITApplicationName1 = GetDetail("Application", ITApplicationCode1)
		End If

		'�h�s�A�v���P�[�V�����Q
		If ITApplicationCode2 <> "" Then
			ITApplicationName2 = GetDetail("Application", ITApplicationCode2)
		End If

		'�h�s�A�v���P�[�V�����R
		If ITApplicationCode3 <> "" Then
			ITApplicationName3 = GetDetail("Application", ITApplicationCode3)
		End If

		'�h�s�J������P
		If ITDevelopmentLanguageCode1 <> "" Then
			ITDevelopmentLanguageName1 = GetDetail("DevelopmentLanguage", ITDevelopmentLanguageCode1)
		End If

		'�h�s�J������Q
		If ITDevelopmentLanguageCode2 <> "" Then
			ITDevelopmentLanguageName2 = GetDetail("DevelopmentLanguage", ITDevelopmentLanguageCode2)
		End If

		'�h�s�f�[�^�x�[�X�P
		If ITDatabaseCode1 <> "" Then
			ITDatabaseName1 = GetDetail("Database", ITDatabaseCode1)
		End If

		'�h�s�f�[�^�x�[�X�Q
		If ITDatabaseCode2 <> "" Then
			ITDatabaseName2 = GetDetail("Database", ITDatabaseCode2)
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
	End Sub

	'******************************************************************************
	'�T�@�v�F���E�Ҍ����y�[�W�֓n��GET�p�����[�^�𐶐����Ď擾�B
	'���@���F
	'���@�l�F������
	'�@�@�@�F�p�����[�^���܂�URL�́AIE�̐�����2048�����܂łł���̂ŁA����ɍ��킹��B
	'���@���F2007/02/27 LIS K.Kokubo �쐬
	'�@�@�@�F2009/06/24 LIS K.Kokubo GetSearchParamBase()����p�����[�^���擾����
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

		Dim sParam

		sParam = ""

		If OrderCode <> "" Then sParam = sParam & "&amp;ordercode=" & OrderCode
		sParam = sParam & GetSearchParamBase()

		If sParam <> "" Then
			'����&amp;���H�ɕϊ�
			sParam = "?" & Mid(sParam, 6)

			'�h�d�̎d�l�̓p�����[�^�̏�����Q�O�S�W�o�C�g
			sParam = Left(sParam, 2048)
		End If

		GetSearchParam = sParam
	End Function

	'******************************************************************************
	'�T�@�v�F���������̊�{����(���R�[�h������������)���擾
	'���@���F
	'���@�l�F
	'���@���F2009/06/24 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Function GetSearchParamBase()
		Dim sParam
		sParam = ""

		If SearchDetailFlag <> "" Then sParam = sParam & "&amp;sdf=" & SearchDetailFlag
		If RegistDayFrom <> "" Then sParam = sParam & "&amp;rdfrom=" & RegistDayFrom
		If HopeWorkingTypeCode <> "" Then sParam = sParam & "&amp;swt=" & HopeWorkingTypeCode
		If WorkingTypeCode1 <> "" Then sParam = sParam & "&amp;swt1=" & WorkingTypeCode1
		If WorkingTypeCode2 <> "" Then sParam = sParam & "&amp;swt2=" & WorkingTypeCode2
		If HopeJobTypeCode1 <> "" Then sParam = sParam & "&amp;shjt1=" & HopeJobTypeCode1
		If HopeJobTypeCode2 <> "" Then sParam = sParam & "&amp;shjt2=" & HopeJobTypeCode2
		If JobTypeCode1 <> "" Then sParam = sParam & "&amp;sjt1=" & JobTypeCode1
		If JobTypeCode2 <> "" Then sParam = sParam & "&amp;sjt2=" & JobTypeCode2
		If JobPeriod1 <> "" Then sParam = sParam & "&amp;sjp1=" & JobPeriod1
		If JobPeriod2 <> "" Then sParam = sParam & "&amp;sjp2=" & JobPeriod2
		If CareerCnt <> "" Then sParam = sParam & "&amp;sccnt=" & CareerCnt
		If HopeIndustryTypeCode <> "" Then sParam = sParam & "&amp;shitc=" & HopeIndustryTypeCode
		If ExpIndustryTypeCode <> "" Then sParam = sParam & "&amp;seitc=" & ExpIndustryTypeCode
		If HopePrefectureCode <> "" Then sParam = sParam & "&amp;shp=" & HopePrefectureCode
		If HopePrefectureCode1 <> "" Then sParam = sParam & "&amp;shp1=" & HopePrefectureCode1
		If HopeCity1 <> "" Then sParam = sParam & "&amp;shc1=" & Server.URLEncode(HopeCity1)
		If HopePrefectureCode2 <> "" Then sParam = sParam & "&amp;shp2=" & HopePrefectureCode2
		If HopeCity2 <> "" Then sParam = sParam & "&amp;shc2=" & Server.URLEncode(HopeCity2)
		If YearlyIncomeMin <> "" Then sParam = sParam & "&amp;syimin=" & YearlyIncomeMin
		If YearlyIncomeMax <> "" Then sParam = sParam & "&amp;syimax=" & YearlyIncomeMax
		If MonthlyIncomeMin <> "" Then sParam = sParam & "&amp;smimin=" & MonthlyIncomeMin
		If MonthlyIncomeMax <> "" Then sParam = sParam & "&amp;smimax=" & MonthlyIncomeMax
		If DailyIncomeMin <> "" Then sParam = sParam & "&amp;sdimin=" & DailyIncomeMin
		If DailyIncomeMax <> "" Then sParam = sParam & "&amp;sdimax=" & DailyIncomeMax
		If HourlyIncomeMin <> "" Then sParam = sParam & "&amp;shimin=" & HourlyIncomeMin
		If HourlyIncomeMax <> "" Then sParam = sParam & "&amp;shimax=" & HourlyIncomeMax
		If PrefectureCode <> "" Then sParam = sParam & "&amp;sp=" & PrefectureCode
		If City <> "" Then sParam = sParam & "&amp;sc=" & Server.URLEncode(City)
		If RailwayLinePrefectureCode <> "" Then sParam = sParam & "&amp;srlpc=" & RailwayLinePrefectureCode
		If RailwayLineCode <> "" Then sParam = sParam & "&amp;srlc=" & RailwayLineCode
		If StationCode <> "" Then sParam = sParam & "&amp;ssc=" & StationCode
		If ZipPrefectureCode <> "" Then sParam = sParam & "&amp;szpc=" & ZipPrefectureCode
		If ZipCode <> "" Then sParam = sParam & "&amp;szc=" & ZipCode
		If AgeMin <> "" Then sParam = sParam & "&amp;samin=" & AgeMin
		If AgeMax <> "" Then sParam = sParam & "&amp;samax=" & AgeMax
		If Sex <> "" Then sParam = sParam & "&amp;ssex=" & Sex
		If SchoolTypeCode <> "" Then sParam = sParam & "&amp;sstc=" & SchoolTypeCode
		If SchoolName <> "" Then sParam = sParam & "&amp;ssn=" & Server.URLEncode(SchoolName)
		If CourseType <> "" Then sParam = sParam & "&amp;sct=" & CourseType
		If FinSchoolTypeCode <> "" Then sParam = sParam & "&amp;sfstc=" & FinSchoolTypeCode
		If GraduateYearMin <> "" Then sParam = sParam & "&amp;sgymin=" & GraduateYearMin
		If GraduateYearMax <> "" Then sParam = sParam & "&amp;sgymax=" & GraduateYearMax
		If LicenseGroupCode1 <> "" Then sParam = sParam & "&amp;slg1=" & LicenseGroupCode1
		If LicenseCategoryCode1 <> "" Then sParam = sParam & "&amp;slc1=" & LicenseCategoryCode1
		If LicenseCode1 <> "" Then sParam = sParam & "&amp;sl1=" & LicenseCode1
		If LicenseGroupCode2 <> "" Then sParam = sParam & "&amp;slg2=" & LicenseGroupCode2
		If LicenseCategoryCode2 <> "" Then sParam = sParam & "&amp;slc2=" & LicenseCategoryCode2
		If LicenseCode2 <> "" Then sParam = sParam & "&amp;sl2=" & LicenseCode2
		If LicenseGroupCode3 <> "" Then sParam = sParam & "&amp;slg3=" & LicenseGroupCode3
		If LicenseCategoryCode3 <> "" Then sParam = sParam & "&amp;slc3=" & LicenseCategoryCode3
		If LicenseCode3 <> "" Then sParam = sParam & "&amp;sl3=" & LicenseCode3
		If LanguageCode <> "" Then sParam = sParam & "&amp;slng=" & LanguageCode
		If LanguageActionLevel1 <> "" Then sParam = sParam & "&amp;slngal1=" & LanguageActionLevel1
		If LanguageActionLevel2 <> "" Then sParam = sParam & "&amp;slngal2=" & LanguageActionLevel2
		If LanguageActionLevel3 <> "" Then sParam = sParam & "&amp;slngal3=" & LanguageActionLevel3
		If OSCode1 & OSCode2 & ApplicationCode1 & ApplicationCode2 & ApplicationCode3 & DevelopmentLanguageCode1 & DevelopmentLanguageCode2 & DatabaseCode1 & DatabaseCode2 <> "" And SkillAndOr <> "" Then sParam = sParam & "&amp;ssao=" & SkillAndOr
		If OSCode1 <> "" Then sParam = sParam & "&amp;sos1=" & OSCode1
		If OSCode2 <> "" Then sParam = sParam & "&amp;sos2=" & OSCode2
		If OSCode1 <> "" And OSPeriod1 <> "" Then sParam = sParam & "&amp;sosp1=" & OSPeriod1
		If OSCode2 <> "" And OSPeriod2 <> "" Then sParam = sParam & "&amp;sosp2=" & OSPeriod2
		If OACode1 <> "" Then sParam = sParam & "&amp;soa1=" & OACode1
		If ApplicationCode1 <> "" Then sParam = sParam & "&amp;sap1=" & ApplicationCode1
		If ApplicationCode2 <> "" Then sParam = sParam & "&amp;sap2=" & ApplicationCode2
		If ApplicationCode3 <> "" Then sParam = sParam & "&amp;sap3=" & ApplicationCode3
		If ApplicationPeriod1 <> "" And ApplicationPeriod1 <> "" Then sParam = sParam & "&amp;sapp1=" & ApplicationPeriod1
		If ApplicationPeriod2 <> "" And ApplicationPeriod2 <> "" Then sParam = sParam & "&amp;sapp2=" & ApplicationPeriod2
		If ApplicationPeriod3 <> "" And ApplicationPeriod3 <> "" Then sParam = sParam & "&amp;sapp3=" & ApplicationPeriod3
		If DevelopmentLanguageCode1 <> "" Then sParam = sParam & "&amp;sdl1=" & DevelopmentLanguageCode1
		If DevelopmentLanguageCode2 <> "" Then sParam = sParam & "&amp;sdl2=" & DevelopmentLanguageCode2
		If DevelopmentLanguageCode1 <> "" And DevelopmentLanguagePeriod1 <> "" Then sParam = sParam & "&amp;sdlp1=" & DevelopmentLanguagePeriod1
		If DevelopmentLanguageCode2 <> "" And DevelopmentLanguagePeriod2 <> "" Then sParam = sParam & "&amp;sdlp2=" & DevelopmentLanguagePeriod2
		If DatabaseCode1 <> "" Then sParam = sParam & "&amp;sdb1=" & DatabaseCode1
		If DatabaseCode2 <> "" Then sParam = sParam & "&amp;sdb2=" & DatabaseCode2
		If DatabaseCode1 <> "" And DatabasePeriod1 <> "" Then sParam = sParam & "&amp;sdbp1=" & DatabasePeriod1
		If DatabaseCode2 <> "" And DatabasePeriod2 <> "" Then sParam = sParam & "&amp;sdbp2=" & DatabasePeriod2
		If (ITOSCode1 & ITOSCode2 & ITApplicationCode1 & ITApplicationCode2 & ITApplicationCode3 & ITDevelopmentLanguageCode1 & ITDevelopmentLanguageCode2 & ITDatabaseCode1 & ITDatabaseCode2 <> "") And ITSkillAndOr <> "" Then sParam = sParam & "&amp;sitsao=" & ITSkillAndOr
		If ITOSCode1 <> "" Then sParam = sParam & "&amp;sitos1=" & ITOSCode1
		If ITOSCode2 <> "" Then sParam = sParam & "&amp;sitos2=" & ITOSCode2
		If ITApplicationCode1 <> "" Then sParam = sParam & "&amp;sitap1=" & ITApplicationCode1
		If ITApplicationCode2 <> "" Then sParam = sParam & "&amp;sitap2=" & ITApplicationCode2
		If ITApplicationCode3 <> "" Then sParam = sParam & "&amp;sitap3=" & ITApplicationCode3
		If ITDevelopmentLanguageCode1 <> "" Then sParam = sParam & "&amp;sitdl1=" & ITDevelopmentLanguageCode1
		If ITDevelopmentLanguageCode2 <> "" Then sParam = sParam & "&amp;sitdl2=" & ITDevelopmentLanguageCode2
		If ITDatabaseCode1 <> "" Then sParam = sParam & "&amp;sitdb1=" & ITDatabaseCode1
		If ITDatabaseCode2 <> "" Then sParam = sParam & "&amp;sitdb2=" & ITDatabaseCode2
		If KeyWord <> "" Then sParam = sParam & "&amp;skw=" & Server.URLEncode(KeyWord)
		If KeyWord <> "" And KeyWordFlag <> "" Then sParam = sParam & "&amp;skwf=" & KeyWordFlag
		If KeyWordHope <> "" Then sParam = sParam & "&amp;skwh=" & Server.URLEncode(KeyWordHope)
		If KeyWordHope <> "" And KeyWordHopeFlag <> "" Then sParam = sParam & "&amp;skwhf=" & KeyWordHopeFlag
		If KeyWordCareer <> "" Then sParam = sParam & "&amp;skwc=" & Server.URLEncode(KeyWordCareer)
		If KeyWordCareer <> "" And KeyWordCareerFlag <> "" Then sParam = sParam & "&amp;skwcf=" & KeyWordCareerFlag
		If KeyWordLicense <> "" Then sParam = sParam & "&amp;skwl=" & Server.URLEncode(KeyWordLicense)
		If KeyWordLicense <> "" And KeyWordLicenseFlag <> "" Then sParam = sParam & "&amp;skwlf=" & KeyWordLicenseFlag
		If KeyWordPerson <> "" Then sParam = sParam & "&amp;skwp=" & Server.URLEncode(KeyWordPerson)
		If KeyWordPerson <> "" And KeyWordPersonFlag <> "" Then sParam = sParam & "&amp;skwpf=" & KeyWordPersonFlag
		If MailFlag <> "" Then sParam = sParam & "&amp;smlf=" & MailFlag
		If StaffCode <> "" Then sParam = sParam & "&amp;sstf=" & StaffCode

		GetSearchParamBase = sParam
	End Function

	'******************************************************************************
	'�T�@�v�F���l�[�ڍ׌����k�n�f�������݂r�p�k���擾
	'�쐬�ҁFLis Kokubo
	'�쐬���F2007/04/04
	'���@���F
	'���@�l�F
	'******************************************************************************
	Public Function GetSQLWriteLog()
		GetSQLWriteLog = "EXEC up_Reg_LOG_SearchStaffDetail '" & CompanyCode & "'" & _
			",'" & ChkSQLStr(Request.ServerVariables("REMOTE_ADDR")) & "'" & _
			",'" & ChkSQLStr(Session.SessionID) & "'" & _
			",'" & ChkSQLStr(Request.ServerVariables("URL")) & "?" & ChkSQLStr(Request.ServerVariables("QUERY_STRING")) & "'" & _
			",'" & ChkSQLStr(Request.ServerVariables("HTTP_REFERER")) & "'" & _
			",'" & WorkingTypeCode1 & "'" & _
			",'" & WorkingTypeCode2 & "'" & _
			",'" & HopeJobTypeCode1 & "'" & _
			",'" & HopeJobTypeCode2 & "'" & _
			",'" & JobTypeCode1 & "'" & _
			",'" & JobTypeCode2 & "'" & _
			",'" & JobPeriod1 & "'" & _
			",'" & JobPeriod2 & "'" & _
			",'" & HopePrefectureCode1 & "'" & _
			",'" & HopeCity1 & "'" & _
			",'" & HopePrefectureCode2 & "'" & _
			",'" & HopeCity2 & "'" & _
			",'" & PrefectureCode & "'" & _
			",'" & AgeMin & "'" & _
			",'" & AgeMax & "'" & _
			",'" & LicenseGroupCode1 & "'" & _
			",'" & LicenseCategoryCode1 & "'" & _
			",'" & LicenseCode1 & "'" & _
			",'" & LicenseGroupCode2 & "'" & _
			",'" & LicenseCategoryCode2 & "'" & _
			",'" & LicenseCode2 & "'" & _
			",'" & LicenseGroupCode3 & "'" & _
			",'" & LicenseCategoryCode3 & "'" & _
			",'" & LicenseCode3 & "'" & _
			",'" & SkillAndOr & "'" & _
			",'" & OSCode1 & "'" & _
			",'" & OSCode2 & "'" & _
			",'" & ApplicationCode1 & "'" & _
			",'" & ApplicationCode2 & "'" & _
			",'" & ApplicationCode3 & "'" & _
			",'" & DevelopmentLanguageCode1 & "'" & _
			",'" & DevelopmentLanguageCode2 & "'" & _
			",'" & DatabaseCode1 & "'" & _
			",'" & DatabaseCode2 & "'" & _
			",'" & ITSkillAndOr & "'" & _
			",'" & ITOSCode1 & "'" & _
			",'" & ITOSCode2 & "'" & _
			",'" & ITApplicationCode1 & "'" & _
			",'" & ITApplicationCode2 & "'" & _
			",'" & ITApplicationCode3 & "'" & _
			",'" & ITDevelopmentLanguageCode1 & "'" & _
			",'" & ITDevelopmentLanguageCode2 & "'" & _
			",'" & ITDatabaseCode1 & "'" & _
			",'" & ITDatabaseCode2 & "'" & _
			",''" & _
			",''" & _
			",'" & KeyWord & "'" & _
			",'" & ChkSQLStr(SQLStaffSearch) & "'"
	End Function

	'******************************************************************************
	'�T�@�v�F���l�[�ڍ׌��������o�͂g�s�l�k���擾
	'�쐬�ҁFLis K.Kokubo
	'�쐬���F2007/04/04
	'���@���F
	'���@�l�F
	'******************************************************************************
	Public Function GetHtmlSearchCondition()
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sError

		Dim sTemp
		Dim sHTML
		Dim idx
		Dim aValue

		GetHtmlSearchCondition = ""
		sHTML = ""

		'���݂̋��l�[
		sTemp = ""
		If OrderCode <> "" Then
			sHTML = sHTML & "<b>[&nbsp;���l�[&nbsp;]</b><a href=""/order/order_detail.asp?ordercode=" & OrderCode & """ target=""_blank"">" & OrderCode & "</a>&nbsp;"
		End If

		If SearchDetailFlag = "1" Then
			'�ڍ׌���

			'�ٗp�`��
			sTemp = ""
			If HopeWorkingTypeCode <> "" Then
				sTemp = HopeWorkingTypeName
			ElseIf WorkingTypeCode1 & WorkingTypeCode2 <> "" Then
				If WorkingTypeCode1 <> "" Then sTemp = sTemp & WorkingTypeName1
				If WorkingTypeCode2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & "�@"
					sTemp = sTemp & WorkingTypeName2
				End If
			End If
			If sTemp <> "" Then sHTML = sHTML & "<b>[&nbsp;�ٗp�`��&nbsp;]</b>" & sTemp & "&nbsp;"

			'��]�E��
			sTemp = ""
			If HopeJobTypeCode1 & HopeJobTypeCode2 <> "" Then
				If HopeJobTypeCode1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & HopeJobTypeName1
				End If
				If HopeJobTypeCode2 <> "" Then
					sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & HopeJobTypeName2
				End If

				sHTML = sHTML & "<b>[&nbsp;��]�E��&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�o���E��
			sTemp = ""
			If JobTypeCode1 & JobTypeCode2 <> "" Then
				If JobTypeCode1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & JobTypeName1
					If JobPeriod1 <> "" Then sTemp = sTemp & "�i" & JobPeriod1 & "�N�ȏ�j"
				End If
				If JobTypeCode2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & JobTypeName2
					If JobPeriod2 <> "" Then sTemp = sTemp & "�i" & JobPeriod2 & "�N�ȏ�j"
				End If

				sHTML = sHTML & "<b>[&nbsp;�o���E��&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'���Љ�
			sTemp = ""
			If CareerCnt <> "" Then
				sHTML = sHTML & "<b>[&nbsp;���Љ�&nbsp;]</b>" & CareerCnt & "��܂�&nbsp;"
			End If

			'��]�Ǝ�
			sTemp = ""
			If HopeIndustryTypeCode <> "" Then
				sHTML = sHTML & "<b>[&nbsp;��]�Ǝ�&nbsp;]</b>" & HopeIndustryTypeName & "&nbsp;"
			End If

			'�o���Ǝ�
			sTemp = ""
			If ExpIndustryTypeCode <> "" Then
				sHTML = sHTML & "<b>[&nbsp;�o���Ǝ�&nbsp;]</b>" & ExpIndustryTypeName & "&nbsp;"
			End If

			'�Ζ��n
			sTemp = ""
			If HopePrefectureCode <> "" Then
				sTemp = HopePrefectureName
			ElseIf HopePrefectureCode1 & HopeCity1 & HopePrefectureCode2 & HopeCity2 <> "" Then
				If HopePrefectureCode1 & HopeCity1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					If HopePrefectureCode1 <> "" Then sTemp = sTemp & HopePrefectureName1
					If HopePrefectureCode1 <> "" And HopeCity1 <> "" Then sTemp = sTemp & "�@"
					If HopeCity1 <> "" Then sTemp = sTemp & HopeCity1
				End If

				If HopePrefectureCode2 & HopeCity2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					If HopePrefectureCode2 <> "" Then sTemp = sTemp & HopePrefectureName2
					If HopePrefectureCode2 <> "" And HopeCity2 <> "" Then sTemp = sTemp & "�@"
					If HopeCity2 <> "" Then sTemp = sTemp & HopeCity2
				End If
			End If
			If sTemp <> "" Then
				sHTML = sHTML & "<b>[&nbsp;�Ζ��n&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'���^
			sTemp = ""
			If YearlyIncomeMin & YearlyIncomeMax & MonthlyIncomeMin & MonthlyIncomeMax & DailyIncomeMin & DailyIncomeMax & HourlyIncomeMin & HourlyIncomeMax <> "" Then
				If YearlyIncomeMin & YearlyIncomeMax <> "" Then
					sTemp = sTemp & "�N���F"
					If YearlyIncomeMin <> "" Then sTemp = sTemp & GetJapaneseYen(YearlyIncomeMin) & "&nbsp;"
					sTemp = sTemp & "�`&nbsp;"
					If YearlyIncomeMax <> "" Then sTemp = sTemp & GetJapaneseYen(YearlyIncomeMax) & "&nbsp;"
				End If

				If MonthlyIncomeMin & MonthlyIncomeMax <> "" Then
					sTemp = sTemp & "�����F"
					If MonthlyIncomeMin <> "" Then sTemp = sTemp & GetJapaneseYen(MonthlyIncomeMin) & "&nbsp;"
					sTemp = sTemp & "�`&nbsp;"
					If MonthlyIncomeMax <> "" Then sTemp = sTemp & GetJapaneseYen(MonthlyIncomeMax) & "&nbsp;"
				End If

				If DailyIncomeMin & DailyIncomeMax <> "" Then
					sTemp = sTemp & "�����F"
					If DailyIncomeMin <> "" Then sTemp = sTemp & GetJapaneseYen(DailyIncomeMin) & "&nbsp;"
					sTemp = sTemp & "�`&nbsp;"
					If DailyIncomeMax <> "" Then sTemp = sTemp & GetJapaneseYen(DailyIncomeMax) & "&nbsp;"
				End If

				If HourlyIncomeMin & HourlyIncomeMax <> "" Then
					sTemp = sTemp & "�����F"
					If HourlyIncomeMin <> "" Then sTemp = sTemp & GetJapaneseYen(HourlyIncomeMin) & "&nbsp;"
					sTemp = sTemp & "�`&nbsp;"
					If HourlyIncomeMax <> "" Then sTemp = sTemp & GetJapaneseYen(HourlyIncomeMax) & "&nbsp;"
				End If
				sHTML = sHTML & "<b>[&nbsp;���^&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�Z��
			sTemp = ""
			If PrefectureCode & City <> "" Then
				If PrefectureCode <> "" Then
					If sTemp <> "" Then sTemp = sTemp & "&nbsp;"
					sTemp = sTemp & "�s���{���F" & PrefectureName
				End If
				If City <> "" Then
					If sTemp <> "" Then sTemp = sTemp & "&nbsp;"
					sTemp = sTemp & "�s��S�F" & City
				End If

				sHTML = sHTML & "<b>[&nbsp;�Z��&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�Ŋ񉈐��E�w
			sTemp = ""
			If RailwayLineCode & StationCode <> "" Then
				If RailwayLinePrefectureCode <> "" Then
					If sTemp <> "" Then sTemp = sTemp & "&nbsp;"
					sTemp = sTemp & "�s���{���F" & RailwayLinePrefectureName
				End If
				If RailwayLineCode <> "" Then
					If sTemp <> "" Then sTemp = sTemp & "&nbsp;"
					sTemp = sTemp & "�Ŋ񉈐��F" & RailwayLineName
				End If
				If StationCode <> "" Then
					If sTemp <> "" Then sTemp = sTemp & "&nbsp;"
					sTemp = sTemp & "�Ŋ�w�F" & StationName
				End If

				sHTML = sHTML & "<b>[&nbsp;�Ŋ񉈐�,�w&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�Z��
			sTemp = ""
			If ZipCode <> "" Then
				sHTML = sHTML & "<b>[&nbsp;�Z���n�ߗ�&nbsp;]</b>" & ZipName & "&nbsp;"
			End If

			'�N��
			sTemp = ""
			If AgeMin & AgeMax <> "" Then
				If AgeMin <> "" Then sTemp = sTemp & AgeMin & "��"
				sTemp = sTemp & "�`"
				If AgeMax <> "" Then sTemp = sTemp & AgeMax & "��"

				sHTML = sHTML & "<b>[&nbsp;�N��&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'����
			sTemp = ""
			If Sex <> "" Then
				If Sex = "1" Then
					sTemp = sTemp & "�j��"
				ElseIf Sex = "2" Then
					sTemp = sTemp & "����"
				End If

				sHTML = sHTML & "<b>[&nbsp;����&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�o���w��
			sTemp = ""
			If SchoolTypeCode <> "" Then
				sHTML = sHTML & "<b>[&nbsp;�o���w��&nbsp;]</b>" & SchoolTypeName & "&nbsp;"
			End If

			'�w�Z��
			sTemp = ""
			If SchoolName <> "" Then
				aValue = Split(SchoolName, ",")
				For idx = LBound(aValue) To UBound(aValue)
					If sTemp <> "" Then sTemp = sTemp & ","
					sTemp = sTemp & aValue(idx)
				Next

				sHTML = sHTML & "<b>[&nbsp;�w�Z��&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�����敪
			sTemp = ""
			If CourseType <> "" Then
				sHTML = sHTML & "<b>[&nbsp;�����敪&nbsp;]</b>" & GetDetail("CourseType", CourseType) & "&nbsp;"
			End If

			'�ŏI�w��
			sTemp = ""
			If FinSchoolTypeName & GraduateYearMin & GraduateYearMax <> "" Then
				If FinSchoolTypeName <> "" Then sTemp = sTemp & FinSchoolTypeName & "��"

				If sTemp <> "" Then sTemp = sTemp & "&nbsp;"

				If GraduateYearMin & GraduateYearMax <> "" Then
					If GraduateYearMin <> "" Then sTemp = sTemp & GraduateYearMin & "�N��"
					sTemp = sTemp & "�`"
					If GraduateYearMax <> "" Then sTemp = sTemp & GraduateYearMax & "�N��"
				End If

				sHTML = sHTML & "<b>[&nbsp;�ŏI�w��&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'���i
			sTemp = ""
			If LicenseName1 & LicenseName2 & LicenseName3 <> "" Then
				If LicenseName1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & LicenseName1
				End If
				If LicenseName2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & LicenseName2
				End If
				If LicenseName3 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & LicenseName3
				End If

				sHTML = sHTML & "<b>[&nbsp;���i&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'��w�X�L��
			sTemp = ""
			If LanguageName & LanguageActionLevelName1 & LanguageActionLevelName2 & LanguageActionLevelName3 <> "" Then
				If LanguageName <> "" Then sTemp = sTemp & "����F" & LanguageName
				If LanguageActionLevelName1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & "��b���x�����u" & LanguageActionLevelName1 & "�v�ȏ�"
				End If
				If LanguageActionLevelName2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & "�ǉ����x�����u" & LanguageActionLevelName2 & "�v�ȏ�"
				End If
				If LanguageActionLevelName3 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & "�앶���x�����u" & LanguageActionLevelName3 & "�v�ȏ�"
				End If

				sHTML = sHTML & "<b>[&nbsp;��w�X�L��&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�n�r
			sTemp = ""
			If OSName1 & OSName2 <> "" Then
				If OSName1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & OSName1
					If OSPeriod1 <> "" Then sTemp = sTemp & "(" & OSPeriod1 & "�N�ȏ�g�p)"
				End If
				If OSName2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & OSName2
					If OSPeriod2 <> "" Then sTemp = sTemp & "(" & OSPeriod2 & "�N�ȏ�g�p)"
				End If

				sHTML = sHTML & "<b>[&nbsp;�n�r&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�A�v���P�[�V����
			sTemp = ""
			If ApplicationName1 <> "" Then
				If ApplicationName1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & ApplicationName1
					If ApplicationPeriod1 <> "" Then sTemp = sTemp & "(" & ApplicationPeriod1 & "�N�ȏ�g�p)"
				End If
				If ApplicationName2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & ApplicationName2
					If ApplicationPeriod2 <> "" Then sTemp = sTemp & "(" & ApplicationPeriod2 & "�N�ȏ�g�p)"
				End If
				If ApplicationName3 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & ApplicationName3
					If ApplicationPeriod3 <> "" Then sTemp = sTemp & "(" & ApplicationPeriod3 & "�N�ȏ�g�p)"
				End If

				sHTML = sHTML & "<b>[&nbsp;�A�v���P�[�V����&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�J������
			sTemp = ""
			If DevelopmentLanguageName1 & DevelopmentLanguageName2 <> "" Then
				If DevelopmentLanguageName1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & DevelopmentLanguageName1
					If DevelopmentLanguagePeriod1 <> "" Then sTemp = sTemp & "(" & DevelopmentLanguagePeriod1 & "�N�ȏ�g�p)"
				End If
				If DevelopmentLanguageName2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & DevelopmentLanguageName2
					If DevelopmentLanguagePeriod2 <> "" Then sTemp = sTemp & "(" & DevelopmentLanguagePeriod2 & "�N�ȏ�g�p)"
				End If

				sHTML = sHTML & "<b>[&nbsp;�J������&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�f�[�^�x�[�X
			sTemp = ""
			If DatabaseName1 & DatabaseName2 <> "" Then
				If DatabaseName1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & DatabaseName1
					If DatabasePeriod1 <> "" Then sTemp = sTemp & "(" & DatabasePeriod1 & "�N�ȏ�g�p)"
				End If
				If DatabaseName2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & DatabaseName2
					If DatabasePeriod2 <> "" Then sTemp = sTemp & "(" & DatabasePeriod2 & "�N�ȏ�g�p)"
				End If

				sHTML = sHTML & "<b>[&nbsp;�f�[�^�x�[�X&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�h�s�n�r
			sTemp = ""
			If ITOSName1 & ITOSName2 <> "" Then
				If ITOSName1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & ITOSName1
				End If
				If ITOSName2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & ITOSName2
				End If

				sHTML = sHTML & "<b>[&nbsp;IT�E��&nbsp;�n�r&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�h�s�A�v���P�[�V����
			sTemp = ""
			If ITApplicationName1 <> "" Then
				If ITApplicationName1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & ITApplicationName1
				End If
				If ITApplicationName2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & ITApplicationName2
				End If
				If ITApplicationName3 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & ITApplicationName3
				End If

				sHTML = sHTML & "<b>[&nbsp;IT�E��&nbsp;�A�v���P�[�V����&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�h�s�J������
			sTemp = ""
			If ITDevelopmentLanguageName1 & ITDevelopmentLanguageName2 <> "" Then
				If ITDevelopmentLanguageName1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & ITDevelopmentLanguageName1
				End If
				If ITDevelopmentLanguageName2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & ITDevelopmentLanguageName2
				End If

				sHTML = sHTML & "<b>[&nbsp;IT�E��&nbsp;�J������&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�h�s�f�[�^�x�[�X
			sTemp = ""
			If ITDatabaseName1 & ITDatabaseName2 <> "" Then
				If ITDatabaseName1 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & ITDatabaseName1
				End If
				If ITDatabaseName2 <> "" Then
					If sTemp <> "" Then sTemp = sTemp & ",&nbsp;"
					sTemp = sTemp & ITDatabaseName2
				End If

				sHTML = sHTML & "<b>[&nbsp;IT�E��&nbsp;�f�[�^�x�[�X&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�L�[���[�h
			sTemp = ""
			If KeyWord <> "" Then
				If KeyWordFlag = "1" Then
					sTemp = sTemp & "(OR����)"
				Else
					sTemp = sTemp & "(AND����)"
				End If
				sTemp = sTemp & KeyWord

				sHTML = sHTML & "<b>[&nbsp;�L�[���[�h&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�L�[���[�h(��])
			sTemp = ""
			If KeyWordHope <> "" Then
				If KeyWordHopeFlag = "1" Then
					sTemp = sTemp & "(OR����)"
				Else
					sTemp = sTemp & "(AND����)"
				End If
				sTemp = sTemp & KeyWordHope

				sHTML = sHTML & "<b>[&nbsp;�L�[���[�h(��])&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�L�[���[�h(�o��)
			sTemp = ""
			If KeyWordCareer <> "" Then
				If KeyWordCareerFlag = "1" Then
					sTemp = sTemp & "(OR����)"
				Else
					sTemp = sTemp & "(AND����)"
				End If
				sTemp = sTemp & KeyWordCareer

				sHTML = sHTML & "<b>[&nbsp;�L�[���[�h(�o��)&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�L�[���[�h(���i�E��w)
			sTemp = ""
			If KeyWordLicense <> "" Then
				If KeyWordLicenseFlag = "1" Then
					sTemp = sTemp & "(OR����)"
				Else
					sTemp = sTemp & "(AND����)"
				End If
				sTemp = sTemp & KeyWordLicense

				sHTML = sHTML & "<b>[&nbsp;�L�[���[�h(���i�E��w)&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�L�[���[�h(���Ȃo�q)
			sTemp = ""
			If KeyWordPerson <> "" Then
				If KeyWordPersonFlag = "1" Then
					sTemp = sTemp & "(OR����)"
				Else
					sTemp = sTemp & "(AND����)"
				End If
				sTemp = sTemp & KeyWordPerson

				sHTML = sHTML & "<b>[&nbsp;�L�[���[�h(���Ȃo�q)&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'���E�҂̓���
			sTemp = ""
			'���[������M�������̂��鋁�E�҂̂�
			If MailFlag <> "" Then
				sHTML = sHTML & "<b>[&nbsp;���E�҂̓���&nbsp;]</b>&nbsp;"
				If MailFlag = "1" Then
					sHTML = sHTML & "�M�Ђ̋��l��񈶂ĂɃ��[���𑗐M�������Ƃ̂��鋁�E��&nbsp;"
				ElseIf MailFlag = "2" Then
					sHTML = sHTML & "���[���̂��Ƃ�̎��т��������E��&nbsp;"
				ElseIf MailFlag = "3" Then
					sHTML = sHTML & "���[���𑗐M�������ԐM�̖������E��&nbsp;"
				End If
			End If

			'���E�҃R�[�h�i�����j
			If StaffCode <> "" Then
				sHTML = "<b>[&nbsp;���E�҃R�[�h&nbsp;]</b>" & StaffCode & "&nbsp;"
			End If

			If sHTML <> "" Then
				sHTML = "<div class=""description""><p class=""m0"">" & sHTML & "</p></div>"
			End If
		Else
			'��������

			'�ٗp�`��
			sTemp = ""
			sSQL = "sp_GetDataWorkingType '" & OrderCode & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Do While GetRSState(oRS) = True
				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & ChkStr(oRS.Collect("WorkingTypeName"))

				oRS.MoveNext
			Loop
			Call RSClose(oRS)
			If sTemp <> "" Then
				sHTML = sHTML & "<b>[&nbsp;�ٗp�`��&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'��]�E��
			sTemp = ""
			sSQL = "sp_GetDataJobType '" & OrderCode & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Do While GetRSState(oRS) = True
				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & ChkStr(oRS.Collect("JobTypeName"))

				oRS.MoveNext
			Loop
			Call RSClose(oRS)
			If sTemp <> "" Then
				sHTML = sHTML & "<b>[&nbsp;��]�E��&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			'�Ζ��n
			sTemp = ""
			sSQL = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
			sSQL = sSQL & "EXEC up_LstC_WorkingPlace '" & OrderCode & "';"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Do While GetRSState(oRS) = True
				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & ChkStr(oRS.Collect("WorkingPlacePrefectureName"))

				oRS.MoveNext
			Loop
			Call RSClose(oRS)
			If sTemp <> "" Then
				sHTML = sHTML & "<b>[&nbsp;��]�Ζ��n&nbsp;]</b>" & sTemp & "&nbsp;"
			End If

			If sHTML <> "" Then
				sHTML = "<div class=""description""><p class=""m0"">" & sHTML & "</p></div>"
			End If
		End If

		GetHtmlSearchCondition = sHTML
	End Function

	'******************************************************************************
	'�T�@�v�F���l�[�ڍ׌����r�p�k���擾 ver.3
	'�쐬�ҁFLis Kokubo
	'�쐬���F2007/04/04
	'���@���F
	'���@�l�F
	'******************************************************************************
	Public Function GetSQLStaffSearchDetail()
		Dim sSQL

		Dim idx
		Dim sJoin		: sJoin = ""
		Dim sWhere		: sWhere = ""
		Dim sDeclare		: sDeclare = ""
		Dim sParams		: sParams = ""
		Dim sDWT		: sDWT = ""
		Dim sFrom
		Dim sTemp
		Dim sTemp2
		Dim iParamNo
		Dim iParamNo2
		Dim aValue
		Dim sSearchCondition

		'���Ћ��l�[�`�F�b�N
		If ChkMyOrder(dbconn, CompanyCode, OrderCode) <> "1" Then Exit Function

		sSQL = ""

		'<�ٗp�`��>
		sTemp = ""
		iParamNo = 1
		If WorkingTypeCode1 & WorkingTypeCode2 <> "" Then
			If WorkingTypeCode1 <> "" Then
				sTemp = sTemp & "PHWT.WorkingTypeCode = @vWorkingTypeCode" & iParamNo & " "

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkingTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vWorkingTypeCode" & iParamNo & " = N'" & WorkingTypeCode1 & "'"
				iParamNo = iParamNo + 1
			End If
			If WorkingTypeCode2 <> "" Then
				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "PHWT.WorkingTypeCode = @vWorkingTypeCode" & iParamNo & " "

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vWorkingTypeCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vWorkingTypeCode" & iParamNo & " = N'" & WorkingTypeCode2 & "'"
				iParamNo = iParamNo + 1
			End If
			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT PHWT.StaffCode FROM P_HopeWorkingType AS PHWT WHERE " & sTemp & ") AS PHWT ON PNS.StaffCode = PHWT.StaffCode" & vbCrLf
		End If
		'</�ٗp�`��>

		'<��]�E��>
		sTemp = ""
		sTemp2 = ""
		iParamNo = 1
		If HopeJobTypeCode1 & HopeJobTypeCode2 <> "" Then
			sTemp = ""
			If HopeJobTypeCode1 <> "" Then
				sTemp = HopeJobTypeCode1
				If Len(HopeJobTypeCode1) < 7 Then sTemp = sTemp & "%"

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vHopeJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vHopeJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "(A.JobTypeCode LIKE @vHopeJobTypeCode" & iParamNo & ") "

				iParamNo = iParamNo + 1
			End If

			sTemp = ""
			If HopeJobTypeCode2 <> "" Then
				sTemp = HopeJobTypeCode2
				If Len(HopeJobTypeCode2) < 7 Then sTemp = sTemp & "%"

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vHopeJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vHopeJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
				sTemp2 = sTemp2 & "(A.JobTypeCode LIKE @vHopeJobTypeCode" & iParamNo & ") "

				iParamNo = iParamNo + 1
			End If

			If sTemp2 <> "" Then
				sTemp2 = Trim(sTemp2)
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_HopeJobType AS A WHERE (" & sTemp2 & ")) AS A ON PNS.StaffCode = A.StaffCode" & vbCrLf
			End If
		End If
		'</��]�E��>

		'<�o���E��>
		sTemp = ""
		sTemp2 = ""
		iParamNo = 1
		If JobTypeCode1 & JobTypeCode2 <> "" Then
			sTemp = ""
			If JobTypeCode1 <> "" Then
				sTemp = JobTypeCode1
				If Len(JobTypeCode1) < 7 Then sTemp = sTemp & "%"

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "UNION "
				sTemp2 = sTemp2 & "SELECT DISTINCT A.StaffCode FROM P_CareerHistory AS A WHERE A.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " "

				If JobPeriod1 <> "" Then
					sTemp2 = sTemp2 & "GROUP BY A.StaffCode,A.JobTypeCode HAVING SUM(A.Period) >= @vJobPeriod" & iParamNo & " "
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vJobPeriod" & iParamNo & " FLOAT"
					sParams = sParams & ",@vJobPeriod" & iParamNo & " = " & JobPeriod1
				End If

				iParamNo = iParamNo + 1
			End If

			sTemp = ""
			If JobTypeCode2 <> "" Then
				sTemp = JobTypeCode2
				If Len(JobTypeCode2) < 7 Then sTemp = sTemp & "%"

				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vJobTypeCode" & iParamNo & " VARCHAR(7)"
				sParams = sParams & ",@vJobTypeCode" & iParamNo & " = N'" & sTemp & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "UNION "
				sTemp2 = sTemp2 & "SELECT DISTINCT A.StaffCode FROM P_CareerHistory AS A WHERE A.JobTypeCode LIKE @vJobTypeCode" & iParamNo & " "

				If JobPeriod2 <> "" Then
					sTemp2 = sTemp2 & "GROUP BY A.StaffCode,A.JobTypeCode HAVING SUM(A.Period) >= @vJobPeriod" & iParamNo & " "
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vJobPeriod" & iParamNo & " FLOAT"
					sParams = sParams & ",@vJobPeriod" & iParamNo & " = " & JobPeriod2
				End If

				iParamNo = iParamNo + 1
			End If
		End If

		If sTemp2 <> "" Then
			sJoin = sJoin & "INNER JOIN (" & Trim(sTemp2) & ") AS PCH ON PNS.StaffCode = PCH.StaffCode" & vbCrLf
		End If
		'</�o���E��>

		'<���Љ�>
		If CareerCnt <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vCareerCnt INT"
			sParams = sParams & ",@vCareerCnt = " & CareerCnt

			sJoin = sJoin & "INNER JOIN (SELECT A.StaffCode FROM P_CareerHistory AS A GROUP BY A.StaffCode HAVING COUNT(*) <= @vCareerCnt) AS PCCNT ON PNS.StaffCode = PCCNT.StaffCode" & vbCrLf
		End If
		'</���Љ�>

		'<��]�Ǝ�>
		sTemp = ""
		sTemp2 = ""
		iParamNo = 1
		If HopeIndustryTypeCode <> "" Then
			aValue = Split(Replace(HopeIndustryTypeCode, " ", ""), ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vHopeIndustryTypeCode" & idx & " VARCHAR(3)"
				sParams = sParams & ",@vHopeIndustryTypeCode" & idx & " = N'" & aValue(idx) & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & ","
				sTemp2 = sTemp2 & "@vHopeIndustryTypeCode" & idx
			Next

			If sTemp2 <> "" Then
				sTemp = sTemp & "A.IndustryTypeCode IN (" & sTemp2 & ") "
			End If

			iParamNo = iParamNo + 1
		End If

		If sTemp <> "" Then
			sTemp = Trim(sTemp)
			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_HopeIndustryType AS A WHERE (" & sTemp & ")) AS PHIT ON PNS.StaffCode = PHIT.StaffCode" & vbCrLf
		End If
		'</��]�Ǝ�>

		'<�o���Ǝ�>
		sTemp = ""
		sTemp2 = ""
		iParamNo = 1
		If ExpIndustryTypeCode <> "" Then
			aValue = Split(Replace(ExpIndustryTypeCode, " ", ""), ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vExpIndustryTypeCode" & idx & " VARCHAR(3)"
				sParams = sParams & ",@vExpIndustryTypeCode" & idx & " = N'" & aValue(idx) & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & ","
				sTemp2 = sTemp2 & "@vExpIndustryTypeCode" & idx
			Next

			If sTemp2 <> "" Then
				sTemp = sTemp & "A.IndustryTypeCode IN (" & sTemp2 & ") "
			End If

			iParamNo = iParamNo + 1
		End If

		If sTemp <> "" Then
			sTemp = Trim(sTemp)
			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_CareerHistory AS A WHERE (" & sTemp & ")) AS PCIT ON PNS.StaffCode = PCIT.StaffCode" & vbCrLf
		End If
		'</�o���Ǝ�>

		'<�Ζ��n>
		sTemp = ""
		iParamNo = 1
		If HopePrefectureCode <> "" Then
			aValue = Split(Replace(HopePrefectureCode, " ", ""), ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vHopePrefectureCode" & idx & " VARCHAR(3)"
				sParams = sParams & ",@vHopePrefectureCode" & idx & " = N'" & aValue(idx) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vHopePrefectureCode" & idx
			Next

			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_HopeWorkingPlace AS A WHERE A.PrefectureCode IN (" & sTemp & ")) AS PHWP ON PNS.StaffCode = PHWP.StaffCode" & vbCrLf
			End If
		ElseIf HopePrefectureCode1 & HopePrefectureCode2 <> "" Then
			If HopePrefectureCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vHopePrefectureCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vHopePrefectureCode" & iParamNo & " = N'" & HopePrefectureCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "(PHWP.PrefectureCode = @vHopePrefectureCode" & iParamNo & " "

				If HopeCity1 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vHopeCity" & iParamNo & " VARCHAR(100)"
					sParams = sParams & ",@vHopeCity" & iParamNo & " = N'%" & HopeCity1 & "%'"

					sTemp = sTemp & " AND PHWP.City LIKE @vHopeCity" & iParamNo & " "
				End If

				sTemp = Trim(sTemp) & ") "

				iParamNo = iParamNo + 1
			End If

			If HopePrefectureCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vHopePrefectureCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vHopePrefectureCode" & iParamNo & " = N'" & HopePrefectureCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "(PHWP.PrefectureCode = @vHopePrefectureCode" & iParamNo & " "

				If HopeCity2 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vHopeCity" & iParamNo & " VARCHAR(100)"
					sParams = sParams & ",@vHopeCity" & iParamNo & " = N'%" & HopeCity2 & "%'"

					sTemp = sTemp & " AND PHWP.City LIKE @vHopeCity" & iParamNo & " "
				End If

				sTemp = Trim(sTemp) & ") "

				iParamNo = iParamNo + 1
			End If

			If sTemp <> "" Then
				sTemp = Trim(sTemp)
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT PHWP.StaffCode FROM P_HopeWorkingPlace AS PHWP WHERE (" & sTemp & ")) AS PHWP ON PNS.StaffCode = PHWP.StaffCode" & vbCrLf
			End If
		End If
		'</�Ζ��n>

		'<���^>
		sTemp = ""
		If YearlyIncomeMin & YearlyIncomeMax & MonthlyIncomeMin & MonthlyIncomeMax & DailyIncomeMin & DailyIncomeMax & HourlyIncomeMin & HourlyIncomeMax <> "" Then
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
					sTemp = sTemp & "(COALESCE(A.YearlyIncomeMin, 0) > 0 AND (A.YearlyIncomeMin BETWEEN @vYearlyIncomeMin AND @vYearlyIncomeMax)) OR (COALESCE(A.YearlyIncomeMax, 0) > 0 AND (A.YearlyIncomeMax BETWEEN @vYearlyIncomeMin AND @vYearlyIncomeMax)) "
				ElseIf YearlyIncomeMin <> "" Then
					'�N�������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "(COALESCE(A.YearlyIncomeMin, 0) > 0 AND A.YearlyIncomeMin >= @vYearlyIncomeMin) OR (COALESCE(A.YearlyIncomeMax, 0) > 0 AND A.YearlyIncomeMax >= @vYearlyIncomeMin) "
				ElseIf YearlyIncomeMax <> "" Then
					'�N������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "(COALESCE(A.YearlyIncomeMin, 0) > 0 AND A.YearlyIncomeMin <= @vYearlyIncomeMax) OR (COALESCE(A.YearlyIncomeMin, 0) = 0 AND COALESCE(A.YearlyIncomeMax, 0) > 0 AND A.YearlyIncomeMax <= @vYearlyIncomeMax) "
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
					sTemp = sTemp & "(COALESCE(A.MonthlyIncomeMin, 0) > 0 AND (A.MonthlyIncomeMin BETWEEN @vMonthlyIncomeMin AND @vMonthlyIncomeMax)) OR (COALESCE(A.MonthlyIncomeMax, 0) > 0 AND (A.MonthlyIncomeMax BETWEEN @vMonthlyIncomeMin AND @vMonthlyIncomeMax)) "
				ElseIf MonthlyIncomeMin <> "" Then
					'���������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "(COALESCE(A.MonthlyIncomeMin, 0) > 0 AND A.MonthlyIncomeMin >= @vMonthlyIncomeMin) OR (COALESCE(A.MonthlyIncomeMax, 0) > 0 AND A.MonthlyIncomeMax >= @vMonthlyIncomeMin) "
				ElseIf MonthlyIncomeMax <> "" Then
					'��������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "(COALESCE(A.MonthlyIncomeMin, 0) > 0 AND A.MonthlyIncomeMin <= @vMonthlyIncomeMax) OR (COALESCE(A.MonthlyIncomeMin, 0) = 0 AND COALESCE(A.MonthlyIncomeMax, 0) > 0 AND A.MonthlyIncomeMax <= @vMonthlyIncomeMax) "
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
					sTemp = sTemp & "(COALESCE(A.DailyIncomeMin, 0) > 0 AND (A.DailyIncomeMin BETWEEN @vDailyIncomeMin AND @vDailyIncomeMax)) OR (COALESCE(A.DailyIncomeMax, 0) > 0 AND (A.DailyIncomeMax BETWEEN @vDailyIncomeMin AND @vDailyIncomeMax)) "
				ElseIf DailyIncomeMin <> "" Then
					'���������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "(COALESCE(A.DailyIncomeMin, 0) > 0 AND A.DailyIncomeMin >= @vDailyIncomeMin) OR (COALESCE(A.DailyIncomeMax, 0) > 0 AND A.DailyIncomeMax >= @vDailyIncomeMin) "
				ElseIf DailyIncomeMax <> "" Then
					'��������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "(COALESCE(A.DailyIncomeMin, 0) > 0 AND A.DailyIncomeMin <= @vDailyIncomeMax) OR (COALESCE(A.DailyIncomeMin, 0) = 0 AND COALESCE(A.DailyIncomeMax, 0) > 0 AND A.DailyIncomeMax <= @vDailyIncomeMax) "
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
					sTemp = sTemp & "(COALESCE(A.HourlyIncomeMin, 0) > 0 AND (A.HourlyIncomeMin BETWEEN @vHourlyIncomeMin AND @vHourlyIncomeMax)) OR (COALESCE(A.HourlyIncomeMax, 0) > 0 AND (A.HourlyIncomeMax BETWEEN @vHourlyIncomeMin AND @vHourlyIncomeMax)) "
				ElseIf HourlyIncomeMin <> "" Then
					'���������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "(COALESCE(A.HourlyIncomeMin, 0) > 0 AND A.HourlyIncomeMin >= @vHourlyIncomeMin) OR (COALESCE(A.HourlyIncomeMax, 0) > 0 AND A.HourlyIncomeMax >= @vHourlyIncomeMin) "
				ElseIf HourlyIncomeMax <> "" Then
					'��������̂ݓ��͂�����ꍇ
					sTemp = sTemp & "(COALESCE(A.HourlyIncomeMin, 0) > 0 AND A.HourlyIncomeMin <= @vHourlyIncomeMax) OR (COALESCE(A.HourlyIncomeMin, 0) = 0 AND COALESCE(A.HourlyIncomeMax, 0) > 0 AND A.HourlyIncomeMax <= @vHourlyIncomeMax) "
				End If
			End If
			'</����>

			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT A.StaffCode FROM P_HopeWorkingCondition AS A WHERE " & Trim(sTemp) & ") AS PSLY ON PNS.StaffCode = PSLY.StaffCode "
			End If
		End If
		'<���^>

		'<�Z���s���{��>
		sTemp = ""
		sTemp2 = ""
		iParamNo = 1
		If PrefectureCode <> "" Then
			aValue = Split(Replace(PrefectureCode, " ", ""), ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vPrefectureCode" & idx & " VARCHAR(3)"
				sParams = sParams & ",@vPrefectureCode" & idx & " = N'" & aValue(idx) & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & ","
				sTemp2 = sTemp2 & "@vPrefectureCode" & idx
			Next

			If sTemp2 <> "" Then
				sTemp = sTemp & "A.PrefectureCode IN (" & sTemp2 & ") "
			End If

			iParamNo = iParamNo + 1
		End If

		If sTemp <> "" Then
			sTemp = Trim(sTemp)
			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_Info AS A WHERE (" & sTemp & ")) AS ADR ON PNS.StaffCode = ADR.StaffCode" & vbCrLf
		End If
		'</�Z���s���{��>

		'<�Z���s��S>
		sTemp = ""
		iParamNo = 1
		If City <> "" Then
			aValue = Split(City, " ")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vCity" & idx & " VARCHAR(100)"
				sParams = sParams & ",@vCity" & idx & " = N'" & aValue(idx) & "'"

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "A.City LIKE '%' + @vCity" & idx & "+ '%' "
			Next

			iParamNo = iParamNo + 1
		End If

		If sTemp <> "" Then
			sTemp = Trim(sTemp)
			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_Info AS A WHERE (" & sTemp & ")) AS ADRCITY ON PNS.StaffCode = ADRCITY.StaffCode" & vbCrLf
		End If
		'</�Z���s��S>

		'<�Z���Ŋ񉈐�,�w>
		sTemp = ""
		iParamNo = 1
		If RailwayLinePrefectureCode <> "" And RailwayLineCode & StationCode <> "" Then
			If StationCode <> "" Then
				'<�w>
				aValue = Split(StationCode, ",")
				For idx = LBound(aValue) To UBound(aValue)
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vStationCode" & idx & " VARCHAR(7)"
					sParams = sParams & ",@vStationCode" & idx & " = N'" & aValue(idx) & "'"

					If sTemp <> "" Then sTemp = sTemp & ","
					sTemp = sTemp & "@vStationCode" & idx
				Next

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_NearbyStation AS A WHERE A.StationCode IN (" & sTemp & ")) AS PS ON PNS.StaffCode = PS.StaffCode" & vbCrLf
				'</�w>
			ElseIf RailwayLineCode <> "" Then
				'<����>
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vRailwayLinePrefectureCode VARCHAR(3)"
				sParams = sParams & ",@vRailwayLinePrefectureCode = N'" & RailwayLinePrefectureCode & "'"
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vRailwayLineCode VARCHAR(7)"
				sParams = sParams & ",@vRailwayLineCode = N'" & RailwayLineCode & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_NearbyStation AS A WHERE EXISTS(SELECT * FROM StationStop AS B INNER JOIN B_Station AS C ON B.StationCode = C.StationCode AND C.PrefectureCode = @vRailwayLinePrefectureCode WHERE A.StationCode = B.StationCode AND B.RailwayLineCode = @vRailwayLineCode)) AS PRL ON PNS.StaffCode = PRL.StaffCode" & vbCrLf
				'</����>
			End If
		End If
		'</�Z���Ŋ񉈐�,�w>

		'<�Z���n�ߗ�>
		sTemp = ""
		iParamNo = 1
		If ZipCode <> "" Then
			aValue = Split(ZipCode, ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vPost_U" & idx & " VARCHAR(3)"
				sParams = sParams & ",@vPost_U" & idx & " = N'" & Left(aValue(idx), 3) & "'"
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vPost_L" & idx & " VARCHAR(1)"
				sParams = sParams & ",@vPost_L" & idx & " = N'" & Right(aValue(idx), 1) & "'"

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "(A.Post_U = @vPost_U" & idx & " AND A.Post_L LIKE @vPost_L" & idx & " + '%') "
			Next

			sJoin = sJoin & "INNER JOIN (SELECT A.StaffCode FROM P_Info AS A WHERE " & Trim(sTemp) & ") AS PZIP ON PNS.StaffCode = PZIP.StaffCode" & vbCrLf
		End If
		'</�Z���n�ߗ�>

		'<�N��>
		sTemp = ""
		iParamNo = 1
		If AgeMin <> "" And AgeMax <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vAgeMin" & iParamNo & " TINYINT"
			sParams = sParams & ",@vAgeMin" & iParamNo & " = " & AgeMin
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vAgeMax" & iParamNo & " TINYINT"
			sParams = sParams & ",@vAgeMax" & iParamNo & " = " & AgeMax

			sJoin = sJoin & "INNER JOIN (SELECT PAGE.StaffCode FROM P_Info AS PAGE WHERE (PAGE.Age BETWEEN @vAgeMin" & iParamNo & " AND @vAgeMax" & iParamNo & ")) AS PAGE ON PNS.StaffCode = PAGE.StaffCode" & vbCrLf
		ElseIf AgeMin <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vAgeMin" & iParamNo & " TINYINT"
			sParams = sParams & ",@vAgeMin" & iParamNo & " = " & AgeMin

			sJoin = sJoin & "INNER JOIN (SELECT PAGE.StaffCode FROM P_Info AS PAGE WHERE PAGE.Age >= @vAgeMin" & iParamNo & ") AS PAGE ON PNS.StaffCode = PAGE.StaffCode" & vbCrLf
		ElseIf AgeMax <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vAgeMax" & iParamNo & " TINYINT"
			sParams = sParams & ",@vAgeMax" & iParamNo & " = " & AgeMax

			sJoin = sJoin & "INNER JOIN (SELECT PAGE.StaffCode FROM P_Info AS PAGE WHERE PAGE.Age <= @vAgeMax" & iParamNo & ") AS PAGE ON PNS.StaffCode = PAGE.StaffCode" & vbCrLf
		End If
		'<�N��>

		'<����>
		If Sex <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSex VARCHAR(1)"
			sParams = sParams & ",@vSex = N'" & Sex & "'"

			sJoin = sJoin & "INNER JOIN (SELECT A.StaffCode FROM P_Info AS A WHERE A.Sex = @vSex) AS PSEX ON PNS.StaffCode = PSEX.StaffCode" & vbCrLf
		End If
		'</����>

		'<�w��>
		If SchoolTypeCode <> "" Then
			aValue = Split(SchoolTypeCode, ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSchoolTypeCode" & idx & " VARCHAR(3)"
				sParams = sParams & ",@vSchoolTypeCode" & idx & " = N'" & aValue(idx) & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vSchoolTypeCode" & idx
			Next

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_EducateHistory AS A WHERE A.SchoolTypeCode IN (" & sTemp & ") AND A.GraduateTypeCode IN ('001','003')) AS PST ON PNS.StaffCode = PST.StaffCode" & vbCrLf
		End If
		'</�w��>

		'<���Ƒ�w>
		sTemp = ""
		If SchoolName <> "" Then
			aValue = Split(SchoolName, ",")
			For idx = LBound(aValue) To UBound(aValue)
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSchoolName" & idx & " VARCHAR(100)"
				sParams = sParams & ",@vSchoolName" & idx & " = N'" & aValue(idx) & "'"

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "(A.SchoolName + B.SchoolTypeName LIKE '%' + @vSchoolName" & idx & " + '%') "
			Next

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_EducateHistory AS A INNER JOIN vw_SchoolType AS B ON A.SchoolTypeCode = B.SchoolTypeCode WHERE A.SchoolTypeCode IN ('006','007') AND (" & Trim(sTemp) & ")) AS PSN ON PNS.StaffCode = PSN.StaffCode" & vbCrLf
		End If
		'</���Ƒ�w>

		'<�w�𕶗��敪>
		If CourseType <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vCourseType VARCHAR(1)"
			sParams = sParams & ",@vCourseType = N'" & CourseType & "'"

			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_EducateHistory AS A WHERE A.CourseType = @vCourseType) AS PSCT ON PNS.StaffCode = PSCT.StaffCode" & vbCrLf
		End If

		'<�ŏI�w��>
		sTemp = ""
		sTemp2 = ""
		If FinSchoolTypeCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vFinSchoolTypeCode VARCHAR(3)"
			sParams = sParams & ",@vFinSchoolTypeCode = '" & FinSchoolTypeCode & "' "

			If sTemp <> "" Then sTemp = sTemp & "AND "
			sTemp = sTemp & "A.SchoolTypeCode = @vFinSchoolTypeCode "
		End If
		If GraduateYearMin <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vGraduateYearMin VARCHAR(4)"
			sParams = sParams & ",@vGraduateYearMin = '" & GraduateYearMin & "' "
		End If
		If GraduateYearMax <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vGraduateYearMax VARCHAR(4)"
			sParams = sParams & ",@vGraduateYearMax = '" & GraduateYearMax & "' "
		End If

		If GraduateYearMin <> "" And GraduateYearMax <> "" Then
			If sTemp <> "" Then sTemp = sTemp & "AND "
			sTemp = sTemp & "(A.GraduateDay BETWEEN @vGraduateYearMin+'0101' AND @vGraduateYearMax+'1231') "
		ElseIf GraduateYearMin <> "" Then
			If sTemp <> "" Then sTemp = sTemp & "AND "
			sTemp = sTemp & "A.GraduateDay >= @vGraduateYearMin+'0101' "
		ElseIf GraduateYearMax <> "" Then
			If sTemp <> "" Then sTemp = sTemp & "AND "
			sTemp = sTemp & "A.GraduateDay <= @vGraduateYearMax+'1231' "
		End If

		If sTemp <> "" Then
			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_EducateHistory AS A INNER JOIN (SELECT A.StaffCode,MAX(A.GraduateDay) AS GraduateDay FROM P_EducateHistory AS A GROUP BY A.StaffCode) AS B ON A.StaffCode = B.StaffCode AND A.GraduateDay = B.GraduateDay WHERE " & RTrim(sTemp) & ") AS PGY ON PNS.StaffCode = PGY.StaffCode" & vbCrLf
		End If
		'</�ŏI�w��>

		'<�ۗL���i>
		sTemp = ""
		iParamNo = 1
		'���i�P
		If LicenseGroupCode1 & LicenseCategoryCode1 & LicenseCode1 <> "" Then
			sTemp2 = ""

			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vLicenseGroupCode" & iParamNo & " VARCHAR(2)"
			sParams = sParams & ",@vLicenseGroupCode" & iParamNo & " = N'" & LicenseGroupCode1 & "'"
			sTemp2 = sTemp2 & "PL.GroupCode = @vLicenseGroupCode" & iParamNo & " "

			If LicenseCategoryCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vLicenseCategoryCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vLicenseCategoryCode" & iParamNo & " = N'" & LicenseCategoryCode1 & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "AND "
				sTemp2 = sTemp2 & "PL.CategoryCode = @vLicenseCategoryCode" & iParamNo & " "
			End If

			If LicenseCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vLicenseCode" & iParamNo & " VARCHAR(2)"
				sParams = sParams & ",@vLicenseCode" & iParamNo & " = N'" & LicenseCode1 & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "AND "
				sTemp2 = sTemp2 & "PL.Code = @vLicenseCode" & iParamNo & " "
			End If

			sTemp = sTemp & "(" & Trim(sTemp2) & ") "

			iParamNo = iParamNo + 1
		End If
		'���i�Q
		If LicenseGroupCode2 & LicenseCategoryCode2 & LicenseCode2 <> "" Then
			sTemp2 = ""

			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vLicenseGroupCode" & iParamNo & " VARCHAR(2)"
			sParams = sParams & ",@vLicenseGroupCode" & iParamNo & " = N'" & LicenseGroupCode2 & "'"
			sTemp2 = sTemp2 & "PL.GroupCode = @vLicenseGroupCode" & iParamNo & " "

			If LicenseCategoryCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vLicenseCategoryCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vLicenseCategoryCode" & iParamNo & " = N'" & LicenseCategoryCode2 & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "AND "
				sTemp2 = sTemp2 & "PL.CategoryCode = @vLicenseCategoryCode" & iParamNo & " "
			End If

			If LicenseCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vLicenseCode" & iParamNo & " VARCHAR(2)"
				sParams = sParams & ",@vLicenseCode" & iParamNo & " = N'" & LicenseCode2 & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "AND "
				sTemp2 = sTemp2 & "PL.Code = @vLicenseCode" & iParamNo & " "
			End If

			If sTemp <> "" Then sTemp = sTemp & "OR "
			sTemp = sTemp & "(" & Trim(sTemp2) & ") "

			iParamNo = iParamNo + 1
		End If
		'���i�R
		If LicenseGroupCode3 & LicenseCategoryCode3 & LicenseCode3 <> "" Then
			sTemp2 = ""

			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vLicenseGroupCode" & iParamNo & " VARCHAR(2)"
			sParams = sParams & ",@vLicenseGroupCode" & iParamNo & " = N'" & LicenseGroupCode3 & "'"
			sTemp2 = sTemp2 & "PL.GroupCode = @vLicenseGroupCode" & iParamNo & " "

			If LicenseCategoryCode3 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vLicenseCategoryCode" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vLicenseCategoryCode" & iParamNo & " = N'" & LicenseCategoryCode3 & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "AND "
				sTemp2 = sTemp2 & "PL.CategoryCode = @vLicenseCategoryCode" & iParamNo & " "
			End If

			If LicenseCode3 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vLicenseCode" & iParamNo & " VARCHAR(2)"
				sParams = sParams & ",@vLicenseCode" & iParamNo & " = N'" & LicenseCode3 & "'"

				If sTemp2 <> "" Then sTemp2 = sTemp2 & "AND "
				sTemp2 = sTemp2 & "PL.Code = @vLicenseCode" & iParamNo & " "
			End If

			If sTemp <> "" Then sTemp = sTemp & "OR "
			sTemp = sTemp & "(" & Trim(sTemp2) & ") "

			iParamNo = iParamNo + 1
		End If

		If sTemp <> "" Then
			sJoin = sJoin & "INNER JOIN (SELECT DISTINCT PL.StaffCode FROM P_License AS PL WHERE (" & sTemp & ")) AS PL ON PNS.StaffCode = PL.StaffCode" & vbCrLf
		End If
		'</�ۗL���i>

		'<��w�X�L��>
		sTemp = ""
		sTemp2 = ""
		'����
		If LanguageCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vLanguageCode VARCHAR(3)"
			sParams = sParams & ",@vLanguageCode = N'" & LanguageCode & "'"

			If sTemp <> "" Then sTemp = sTemp & "AND "
			sTemp = sTemp & "A.LanguageCode = @vLanguageCode "

			'����A�N�V�������x��
			If LanguageActionLevel1 & LanguageActionLevel2 & LanguageActionLevel3 <> "" Then
				'��b���x��
				If LanguageActionLevel1 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vLanguageActionLevel1 TINYINT"
					sParams = sParams & ",@vLanguageActionLevel1 = N'" & LanguageActionLevel1 & "'"

					If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
					sTemp2 = sTemp2 & "(A.StaffCode = B.StaffCode AND A.LanguageSeq = B.LanguageSeq AND B.LanguageActionCode = '1' AND B.LanguageActionLevel >= @vLanguageActionLevel1) "
				End If
				'�ǉ����x��
				If LanguageActionLevel2 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vLanguageActionLevel2 TINYINT"
					sParams = sParams & ",@vLanguageActionLevel2 = N'" & LanguageActionLevel2 & "'"

					If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
					sTemp2 = sTemp2 & "(A.StaffCode = B.StaffCode AND A.LanguageSeq = B.LanguageSeq AND B.LanguageActionCode = '2' AND B.LanguageActionLevel >= @vLanguageActionLevel2) "
				End If
				'�앶���x��
				If LanguageActionLevel3 <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vLanguageActionLevel3 TINYINT"
					sParams = sParams & ",@vLanguageActionLevel3 = N'" & LanguageActionLevel3 & "'"

					If sTemp2 <> "" Then sTemp2 = sTemp2 & "OR "
					sTemp2 = sTemp2 & "(A.StaffCode = B.StaffCode AND A.LanguageSeq = B.LanguageSeq AND B.LanguageActionCode = '3' AND B.LanguageActionLevel >= @vLanguageActionLevel3) "
				End If

				sTemp = sTemp & "AND "
				sTemp = sTemp & "EXISTS(SELECT * FROM P_Skill_LanguageLevel AS B WHERE " & Trim(sTemp2) & ") "
			End If

			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_Skill_Language AS A WHERE " & Trim(sTemp) & ") AS PLNG ON PNS.StaffCode = PLNG.StaffCode" & vbCrLf
			End If
		Else
			'����w�薳���̌���A�N�V�������x���̂�
			sTemp = ""
			'��b���x��
			If LanguageActionLevel1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vLanguageActionLevel1 TINYINT"
				sParams = sParams & ",@vLanguageActionLevel1 = N'" & LanguageActionLevel1 & "'"

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "(A.LanguageActionCode = '1' AND A.LanguageActionLevel >= @vLanguageActionLevel1) "
			End If
			'�ǉ����x��
			If LanguageActionLevel2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vLanguageActionLevel2 TINYINT"
				sParams = sParams & ",@vLanguageActionLevel2 = N'" & LanguageActionLevel2 & "'"

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "(A.LanguageActionCode = '2' AND A.LanguageActionLevel >= @vLanguageActionLevel2) "
			End If
			'�앶���x��
			If LanguageActionLevel3 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vLanguageActionLevel3 TINYINT"
				sParams = sParams & ",@vLanguageActionLevel3 = N'" & LanguageActionLevel3 & "'"

				If sTemp <> "" Then sTemp = sTemp & "OR "
				sTemp = sTemp & "(A.LanguageActionCode = '3' AND A.LanguageActionLevel >= @vLanguageActionLevel3) "
			End If
			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.StaffCode FROM P_Skill_LanguageLevel AS A WHERE " & Trim(sTemp) & ") AS PLNGLVL ON PNS.StaffCode = PLNGLVL.StaffCode" & vbCrLf
			End If
		End If
		'<��w�X�L��>

		'<�X�L��>
		If SkillAndOr = "AND" Then
			'**************************************
			'** AND start
			'**************************************
			'OA
			iParamNo = 1
			If OACode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOA" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vOA" & iParamNo & " = N'" & OACode1 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT OA" & iParamNo & ".StaffCode FROM P_Skill AS OA" & iParamNo & " WHERE OA" & iParamNo & ".CategoryCode = 'OA' AND OA" & iParamNo & ".Code = @vOA" & iParamNo & ") AS OA" & iParamNo & " ON PNS.StaffCode = OA" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If

			'OS
			iParamNo = 1
			If OSCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOS" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vOS" & iParamNo & " = N'" & OSCode1 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT OS" & iParamNo & ".StaffCode FROM P_Skill AS OS" & iParamNo & " WHERE OS" & iParamNo & ".CategoryCode = 'OS' AND OS" & iParamNo & ".Code = @vOS" & iParamNo & ") AS OS" & iParamNo & " ON PNS.StaffCode = OS" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If
			If OSCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOS" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vOS" & iParamNo & " = N'" & OSCode2 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT OS" & iParamNo & ".StaffCode FROM P_Skill AS OS" & iParamNo & " WHERE OS" & iParamNo & ".CategoryCode = 'OS' AND OS" & iParamNo & ".Code = @vOS" & iParamNo & ") AS OS" & iParamNo & " ON PNS.StaffCode = OS" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If

			'�A�v���P�[�V����
			sTemp = ""
			iParamNo = 1
			If ApplicationCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vAPP" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vAPP" & iParamNo & " = N'" & ApplicationCode1 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT APP" & iParamNo & ".StaffCode FROM P_Skill AS APP" & iParamNo & " WHERE APP" & iParamNo & ".CategoryCode = 'Application' AND APP" & iParamNo & ".Code = '" & ApplicationCode1 & "') AS APP" & iParamNo & " ON PNS.StaffCode = APP" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If
			If ApplicationCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vAPP" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vAPP" & iParamNo & " = N'" & ApplicationCode2 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT APP2.StaffCode FROM P_Skill AS APP" & iParamNo & " WHERE APP" & iParamNo & ".CategoryCode = 'Application' AND APP" & iParamNo & ".Code = '" & ApplicationCode2 & "') AS APP" & iParamNo & " ON PNS.StaffCode = APP" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If
			If ApplicationCode3 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vAPP" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vAPP" & iParamNo & " = N'" & ApplicationCode3 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT APP3.StaffCode FROM P_Skill AS APP" & iParamNo & " WHERE APP" & iParamNo & ".CategoryCode = 'Application' AND APP" & iParamNo & ".Code = '" & ApplicationCode3 & "') AS APP" & iParamNo & " ON PNS.StaffCode = APP" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If

			'�J������
			sTemp = ""
			iParamNo = 1
			If DevelopmentLanguageCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vDL" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vDL" & iParamNo & " = N'" & DevelopmentLanguageCode1 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT DL1.StaffCode FROM P_Skill AS DL" & iParamNo & " WHERE DL" & iParamNo & ".CategoryCode = 'DevelopmentLanguage' AND DL" & iParamNo & ".Code = '" & DevelopmentLanguageCode1 & "') AS DL" & iParamNo & " ON PNS.StaffCode = DL" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If
			If DevelopmentLanguageCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vDL" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vDL" & iParamNo & " = N'" & DevelopmentLanguageCode2 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT DL2.StaffCode FROM P_Skill AS DL" & iParamNo & " WHERE DL" & iParamNo & ".CategoryCode = 'DevelopmentLanguage' AND DL" & iParamNo & ".Code = '" & DevelopmentLanguageCode2 & "') AS DL" & iParamNo & " ON PNS.StaffCode = DL" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If

			'�f�[�^�x�[�X
			sTemp = ""
			iParamNo = 1
			If DatabaseCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vDB" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vDB" & iParamNo & " = N'" & DatabaseCode1 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT DB1.StaffCode FROM P_Skill AS DB" & iParamNo & " WHERE DB" & iParamNo & ".CategoryCode = 'Database' AND DB" & iParamNo & ".Code = '" & DatabaseCode1 & "') AS DB" & iParamNo & " ON PNS.StaffCode = DB" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If
			If DatabaseCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vDB" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vDB" & iParamNo & " = N'" & DatabaseCode2 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT DB2.StaffCode FROM P_Skill AS DB" & iParamNo & " WHERE DB" & iParamNo & ".CategoryCode = 'Database' AND DB" & iParamNo & ".Code = '" & DatabaseCode2 & "') AS DB" & iParamNo & " ON PNS.StaffCode = DB" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If
			'**************************************
			'** AND end
			'**************************************
		Else
			'**************************************
			'** OR start
			'**************************************
			'OA
			sTemp = ""
			sTemp2 = ""
			iParamNo = 1
			If OACode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOA" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vOA" & iParamNo & " = N'" & OACode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vOA" & iParamNo
			End If
			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT OA" & iParamNo & ".StaffCode FROM P_Skill AS OA" & iParamNo & " WHERE OA" & iParamNo & ".CategoryCode = 'OA' AND OA" & iParamNo & ".Code IN (" & sTemp & ")) AS OA" & iParamNo & " ON PNS.StaffCode = OA" & iParamNo & ".StaffCode" & vbCrLf
			End If

			'OS
			sTemp = ""
			iParamNo = 1
			If OSCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOS" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vOS" & iParamNo & " = N'" & OSCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vOS" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If OSCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOS" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vOS" & iParamNo & " = N'" & OSCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vOS" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT OS.StaffCode FROM P_Skill AS OS WHERE OS.CategoryCode = 'OS' AND OS.Code IN (" & sTemp & ")) AS OS ON PNS.StaffCode = OS.StaffCode" & vbCrLf
			End If

			'�A�v���P�[�V����
			sTemp = ""
			iParamNo = 1
			If ApplicationCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vAPP" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vAPP" & iParamNo & " = N'" & ApplicationCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vAPP" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If ApplicationCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vAPP" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vAPP" & iParamNo & " = N'" & ApplicationCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vAPP" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If ApplicationCode3 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vAPP" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vAPP" & iParamNo & " = N'" & ApplicationCode3 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vAPP" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT APP.StaffCode FROM P_Skill AS APP WHERE APP.CategoryCode = 'Application' AND APP.Code IN (" & sTemp & ")) AS APP ON PNS.StaffCode = APP.StaffCode" & vbCrLf
			End If

			'�J������
			sTemp = ""
			iParamNo = 1
			If DevelopmentLanguageCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vDL" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vDL" & iParamNo & " = N'" & DevelopmentLanguageCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vDL" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If DevelopmentLanguageCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vDL" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vDL" & iParamNo & " = N'" & DevelopmentLanguageCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vDL" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT DL.StaffCode FROM P_Skill AS DL WHERE DL.CategoryCode = 'DevelopmentLanguage' AND DL.Code IN (" & sTemp & ")) AS DL ON PNS.StaffCode = DL.StaffCode" & vbCrLf
			End If

			'�f�[�^�x�[�X
			sTemp = ""
			iParamNo = 1
			If DatabaseCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vDB" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vDB" & iParamNo & " = N'" & DatabaseCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vDB" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If DatabaseCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vDB" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vDB" & iParamNo & " = N'" & DatabaseCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vDB" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT DB.StaffCode FROM P_Skill AS DB WHERE DB.CategoryCode = 'Database' AND DB.Code IN (" & sTemp & ")) AS DB ON PNS.StaffCode = DB.StaffCode" & vbCrLf
			End If
			'**************************************
			'** OR end
			'**************************************
		End If
		'<�X�L��>

		'<�h�s�E���ڍ�>
		sTemp = ""
		sTemp2 = ""
		If ITSkillAndOr = "AND" Then
			'**************************************
			'** AND start
			'**************************************
			'OS
			iParamNo = 1
			If ITOSCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITOS" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITOS" & iParamNo & " = N'" & ITOSCode1 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITOS" & iParamNo & ".StaffCode FROM P_DevelopmentTool AS ITOS" & iParamNo & " WHERE ITOS" & iParamNo & ".CategoryCode = 'OS' AND ITOS" & iParamNo & ".Code = @vITOS" & iParamNo & ") AS ITOS" & iParamNo & " ON PNS.StaffCode = ITOS" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If
			If ITOSCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITOS" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITOS" & iParamNo & " = N'" & ITOSCode2 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITOS" & iParamNo & ".StaffCode FROM P_DevelopmentTool AS ITOS" & iParamNo & " WHERE ITOS" & iParamNo & ".CategoryCode = 'OS' AND ITOS" & iParamNo & ".Code = @vITOS" & iParamNo & ") AS ITOS" & iParamNo & " ON PNS.StaffCode = ITOS" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If

			'�A�v���P�[�V����
			sTemp = ""
			iParamNo = 1
			If ITApplicationCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITAPP" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITAPP" & iParamNo & " = N'" & ITApplicationCode1 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITAPP" & iParamNo & ".StaffCode FROM P_DevelopmentTool AS ITAPP" & iParamNo & " WHERE ITAPP" & iParamNo & ".CategoryCode = 'Application' AND ITAPP" & iParamNo & ".Code = @vITAPP" & iParamNo & ") AS ITAPP" & iParamNo & " ON PNS.StaffCode = ITAPP" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If
			If ITApplicationCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITAPP" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITAPP" & iParamNo & " = N'" & ITApplicationCode2 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITAPP" & iParamNo & ".StaffCode FROM P_DevelopmentTool AS ITAPP" & iParamNo & " WHERE ITAPP" & iParamNo & ".CategoryCode = 'Application' AND ITAPP" & iParamNo & ".Code = @vITAPP" & iParamNo & ") AS ITAPP" & iParamNo & " ON PNS.StaffCode = ITAPP" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If
			If ITApplicationCode3 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITAPP" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITAPP" & iParamNo & " = N'" & ITApplicationCode3 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITAPP" & iParamNo & ".StaffCode FROM P_DevelopmentTool AS ITAPP" & iParamNo & " WHERE ITAPP" & iParamNo & ".CategoryCode = 'Application' AND ITAPP" & iParamNo & ".Code = @vITAPP" & iParamNo & ") AS ITAPP" & iParamNo & " ON PNS.StaffCode = ITAPP" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If

			'�J������
			sTemp = ""
			iParamNo = 1
			If ITDevelopmentLanguageCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITDL" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITDL" & iParamNo & " = N'" & ITDevelopmentLanguageCode1 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITDL" & iParamNo & ".StaffCode FROM P_DevelopmentTool AS ITDL" & iParamNo & " WHERE ITDL" & iParamNo & ".CategoryCode = 'DevelopmentLanguage' AND ITDL" & iParamNo & ".Code = @vITDL" & iParamNo & ") AS ITDL" & iParamNo & " ON PNS.StaffCode = ITDL" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If
			If ITDevelopmentLanguageCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITDL" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITDL" & iParamNo & " = N'" & ITDevelopmentLanguageCode2 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITDL" & iParamNo & ".StaffCode FROM P_DevelopmentTool AS ITDL" & iParamNo & " WHERE ITDL" & iParamNo & ".CategoryCode = 'DevelopmentLanguage' AND ITDL" & iParamNo & ".Code = @vITDL" & iParamNo & ") AS ITDL" & iParamNo & " ON PNS.StaffCode = ITDL" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If

			'�f�[�^�x�[�X
			sTemp = ""
			iParamNo = 1
			If ITDatabaseCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITDB" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITDB" & iParamNo & " = N'" & ITDatabaseCode1 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITDB" & iParamNo & ".StaffCode FROM P_DevelopmentTool AS ITDB" & iParamNo & " WHERE ITDB" & iParamNo & ".CategoryCode = 'Database' AND ITDB" & iParamNo & ".Code = @vITDB" & iParamNo & ") AS ITDB" & iParamNo & " ON PNS.StaffCode = ITDB" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If
			If ITDatabaseCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITDB" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITDB" & iParamNo & " = N'" & ITDatabaseCode2 & "'"

				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITDB" & iParamNo & ".StaffCode FROM P_DevelopmentTool AS ITDB" & iParamNo & " WHERE ITDB" & iParamNo & ".CategoryCode = 'Database' AND ITDB" & iParamNo & ".Code = @vITDB" & iParamNo & ") AS ITDB" & iParamNo & " ON PNS.StaffCode = ITDB" & iParamNo & ".StaffCode" & vbCrLf

				iParamNo = iParamNo + 1
			End If
			'**************************************
			'** AND end
			'**************************************
		Else
			'**************************************
			'** OR start
			'**************************************
			'OS
			sTemp = ""
			iParamNo = 1
			If ITOSCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITOS" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITOS" & iParamNo & " = N'" & ITOSCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vITOS" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If ITOSCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITOS" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITOS" & iParamNo & " = N'" & ITOSCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vITOS" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITOS.StaffCode FROM P_DevelopmentTool AS ITOS WHERE ITOS.CategoryCode = 'OS' AND ITOS.Code IN (" & sTemp & ")) AS ITOS ON PNS.StaffCode = ITOS.StaffCode" & vbCrLf
			End If

			'�A�v���P�[�V����
			sTemp = ""
			iParamNo = 1
			If ITApplicationCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITAPP" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITAPP" & iParamNo & " = N'" & ITApplicationCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vITAPP" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If ITApplicationCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITAPP" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITAPP" & iParamNo & " = N'" & ITApplicationCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vITAPP" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If ITApplicationCode3 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITAPP" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITAPP" & iParamNo & " = N'" & ITApplicationCode3 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vITAPP" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITAPP.StaffCode FROM P_DevelopmentTool AS ITAPP WHERE ITAPP.CategoryCode = 'Application' AND ITAPP.Code IN (" & sTemp & ")) AS ITAPP ON PNS.StaffCode = ITAPP.StaffCode" & vbCrLf
			End If

			'�J������
			sTemp = ""
			iParamNo = 1
			If ITDevelopmentLanguageCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITDL" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITDL" & iParamNo & " = N'" & ITDevelopmentLanguageCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vITDL" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If ITDevelopmentLanguageCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITDL" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITDL" & iParamNo & " = N'" & ITDevelopmentLanguageCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vITDL" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITDL.StaffCode FROM P_DevelopmentTool AS ITDL WHERE ITDL.CategoryCode = 'DevelopmentLanguage' AND ITDL.Code IN (" & sTemp & ")) AS ITDL ON PNS.StaffCode = ITDL.StaffCode" & vbCrLf
			End If

			'�f�[�^�x�[�X
			sTemp = ""
			iParamNo = 1
			If ITDatabaseCode1 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITDB" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITDB" & iParamNo & " = N'" & ITDatabaseCode1 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vITDB" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If ITDatabaseCode2 <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vITDB" & iParamNo & " VARCHAR(3)"
				sParams = sParams & ",@vITDB" & iParamNo & " = N'" & ITDatabaseCode2 & "'"

				If sTemp <> "" Then sTemp = sTemp & ","
				sTemp = sTemp & "@vITDB" & iParamNo

				iParamNo = iParamNo + 1
			End If
			If sTemp <> "" Then
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT ITDB.StaffCode FROM P_DevelopmentTool AS ITDB WHERE ITDB.CategoryCode = 'Database' AND ITDB.Code IN (" & sTemp & ")) AS ITDB ON PNS.StaffCode = ITDB.StaffCode" & vbCrLf
			End If
			'**************************************
			'** OR end
			'**************************************
		End If
		'<�h�s�E���ڍ�>

		'<�L�[���[�h>
		sTemp = ""
		sTemp2 = ""
		If KeyWord <> "" Then
			aValue = Split(Replace(KeyWord, "�@", " "), " ")
			For idx = LBound(aValue) To UBound(aValue)
				If sTemp <> "" Then
					If KeyWordFlag = "1" Then
						sTemp = sTemp & " OR "
					ElseIf KeyWordFlag = "2" Then
						sTemp = sTemp & " AND "
					End If
				End If
				sTemp = sTemp & "FORMSOF(THESAURUS, " & aValue(idx) & "*)"
			Next
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vKeyWord VARCHAR(400)"
			sParams = sParams & ",@vKeyWord = N'" & sTemp & "'"

			sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY PFTN.StaffCode) AS Num, PFTN.StaffCode FROM (SELECT A.StaffCode FROM FTIStaffHopeNAVI AS A WHERE CONTAINS(A.Txt,@vKeyWord) UNION SELECT A.StaffCode FROM FTIStaffCareerNAVI AS A WHERE CONTAINS(A.Txt,@vKeyWord) UNION SELECT A.StaffCode FROM FTIStaffLicenseNAVI AS A WHERE CONTAINS(A.Txt,@vKeyWord)) AS PFTN) AS KW ON PNS.StaffCode = KW.StaffCode" & vbCrLf
		End If
		'</�L�[���[�h>

		'<�L�[���[�h(��])>
		sTemp = ""
		sTemp2 = ""
		If KeyWordHope <> "" Then
			aValue = Split(Replace(KeyWordHope, "�@", " "), " ")
			For idx = LBound(aValue) To UBound(aValue)
				If sTemp <> "" Then
					If KeyWordHopeFlag = "1" Then
						sTemp = sTemp & " OR "
					ElseIf KeyWordHopeFlag = "2" Then
						sTemp = sTemp & " AND "
					End If
				End If
				sTemp = sTemp & "FORMSOF(THESAURUS, " & aValue(idx) & "*)"
			Next
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vKeyWordHope VARCHAR(400)"
			sParams = sParams & ",@vKeyWordHope = N'" & sTemp & "'"

			sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY FSH.StaffCode) AS Num, FSH.StaffCode FROM FTIStaffHopeNAVI AS FSH WHERE CONTAINS(FSH.Txt, @vKeyWordHope)) AS KWH ON PNS.StaffCode = KWH.StaffCode" & vbCrLf
		End If
		'</�L�[���[�h(��])>

		'<�L�[���[�h(�o��)>
		sTemp = ""
		sTemp2 = ""
		If KeyWordCareer <> "" Then
			aValue = Split(Replace(KeyWordCareer, "�@", " "), " ")
			For idx = LBound(aValue) To UBound(aValue)
				If sTemp <> "" Then
					If KeyWordCareerFlag = "1" Then
						sTemp = sTemp & " OR "
					ElseIf KeyWordCareerFlag = "2" Then
						sTemp = sTemp & " AND "
					End If
				End If
				sTemp = sTemp & "FORMSOF(THESAURUS, " & aValue(idx) & "*)"
			Next
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vKeyWordCareer VARCHAR(400)"
			sParams = sParams & ",@vKeyWordCareer = N'" & sTemp & "'"

			sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY FSC.StaffCode) AS Num, FSC.StaffCode FROM FTIStaffCareerNAVI AS FSC WHERE CONTAINS(FSC.Txt, @vKeyWordCareer)) AS KWC ON PNS.StaffCode = KWC.StaffCode" & vbCrLf
		End If
		'</�L�[���[�h(�o��)>

		'<�L�[���[�h(���i�E��w)>
		sTemp = ""
		sTemp2 = ""
		If KeyWordLicense <> "" Then
			aValue = Split(Replace(KeyWordLicense, "�@", " "), " ")
			For idx = LBound(aValue) To UBound(aValue)
				If sTemp <> "" Then
					If KeyWordLicenseFlag = "1" Then
						sTemp = sTemp & " OR "
					ElseIf KeyWordLicenseFlag = "2" Then
						sTemp = sTemp & " AND "
					End If
				End If
				sTemp = sTemp & "FORMSOF(THESAURUS, " & aValue(idx) & "*)"
			Next
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vKeyWordLicense VARCHAR(400)"
			sParams = sParams & ",@vKeyWordLicense = N'" & sTemp & "'"

			sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY FSL.StaffCode) AS Num, FSL.StaffCode FROM FTIStaffLicenseNAVI AS FSL WHERE CONTAINS(FSL.Txt, @vKeyWordLicense)) AS KWL ON PNS.StaffCode = KWL.StaffCode" & vbCrLf
		End If
		'</�L�[���[�h(���i�E��w)>

		'<�L�[���[�h(���Ȃo�q)>
		sTemp = ""
		sTemp2 = ""
		If KeyWordPerson <> "" Then
			aValue = Split(Replace(KeyWordPerson, "�@", " "), " ")
			For idx = LBound(aValue) To UBound(aValue)
				If sTemp <> "" Then
					If KeyWordPersonFlag = "1" Then
						sTemp = sTemp & " OR "
					ElseIf KeyWordPersonFlag = "2" Then
						sTemp = sTemp & " AND "
					End If
				End If
				sTemp = sTemp & "FORMSOF(THESAURUS, " & aValue(idx) & "*)"
			Next
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vKeyWordPerson VARCHAR(400)"
			sParams = sParams & ",@vKeyWordPerson = N'" & sTemp & "'"

			sJoin = sJoin & "INNER JOIN (SELECT ROW_NUMBER() OVER(ORDER BY FSP.StaffCode) AS Num, FSP.StaffCode FROM FTIStaffPersonNAVI AS FSP WHERE CONTAINS(FSP.Txt, @vKeyWordPerson)) AS KWP ON PNS.StaffCode = KWP.StaffCode" & vbCrLf
		End If
		'</�L�[���[�h(���i�E��w)>

		'<���[������M�������̂��鋁�E�҂̂�>
		If MailFlag <> "" Then
			If InStr(sDeclare,"@vOrderCode ") = 0 Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOrderCode VARCHAR(8)"
				sParams = sParams & ",@vOrderCode = N'" & OrderCode & "'"
			End If
			If InStr(sDeclare,"@vCompanyCode ") = 0 Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vCompanyCode VARCHAR(8)"
				sParams = sParams & ",@vCompanyCode = N'" & CompanyCode & "'"
			End If

			If MailFlag = "1" Then
				'���[������M�������̂��鋁�E��
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.SenderCode AS StaffCode FROM MailHistory AS A WHERE A.SenderCode LIKE 'S%' AND A.OrderCode = @vOrderCode) AS MR ON PNS.StaffCode = MR.StaffCode" & vbCrLf
			ElseIf MailFlag = "2" Then
				'���[���̂��Ƃ�̎��т��������E��
				If sWhere <> "" Then sWhere = sWhere & "AND "
				sWhere = sWhere & "NOT EXISTS(SELECT * FROM MailHistory AS Z WHERE PNS.StaffCode IN (Z.SenderCode,Z.ReceiverCode) AND Z.OrderCode = @vOrderCode) "
			ElseIf MailFlag = "3" Then
				'���[���𑗐M�������ԐM�̖������E��
				sJoin = sJoin & "INNER JOIN (SELECT DISTINCT A.ReceiverCode AS StaffCode FROM MailHistory AS A WHERE A.OrderCode = @vOrderCode AND A.SenderCode = @vCompanyCode AND NOT EXISTS(SELECT * FROM MailHistory AS Z WHERE A.OrderCode = Z.OrderCode AND A.ReceiverCode = Z.SenderCode)) AS ML ON PNS.StaffCode = ML.StaffCode" & vbCrLf
			End If
		End If
		'</���[������M�������̂��鋁�E�҂̂�>

		'<�X�^�b�t�R�[�h����>
		If StaffCode <> "" Then
			sJoin = ""

			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vStaffCode VARCHAR(8)"
			sParams = sParams & ",@vStaffCode = N'" & StaffCode & "'"

			sWhere = "PNS.StaffCode LIKE @vStaffCode "
		End If
		'</�X�^�b�t�R�[�h����>

		'<�}�b�`���O�Ώۋ��E�҃R�[�h>
		If MchStaffCode <> "" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vMchStaffCode VARCHAR(8)"
			sParams = sParams & ",@vMchStaffCode = N'" & MchStaffCode & "'"

			sWhere = "PNS.StaffCode = @vMchStaffCode "
		End If
		'</�}�b�`���O�Ώۋ��E�҃R�[�h>

		'<�K�ޑҋ@�v�����ʒm���[��>
		If SpMchNoticeFlag = "1" Then
			If sDeclare <> "" Then sDeclare = sDeclare & ","
			sDeclare = sDeclare & "@vSpMchCompanyCode VARCHAR(8)"
			sParams = sParams & ",@vSpMchCompanyCode = N'" & CompanyCode & "'"

			'�X�V�����P�O���ȓ��̐l�����Ώ�
			sWhere = ""
			sWhere = sWhere & "PUI.UpdateDay >= CONVERT(VARCHAR(8),DATEADD(DAY,-9,GETDATE()),112) "
			sWhere = sWhere & "AND NOT EXISTS(SELECT * FROM CMPPaySpMchNotice AS EXT WHERE PUI.StaffCode = EXT.StaffCode AND EXT.CompanyCode = @vSpMchCompanyCode) "
		End If
		'</�K�ޑҋ@�v�����ʒm���[��>

		If sWhere <> "" Then sWhere = "AND " & sWhere

		If SearchDetailFlag = "1" Then
			sSQL = ""
			sSQL = sSQL & "SELECT"
			If CStr(Top) <> "" Then sSQL = sSQL & " TOP " & Top
			sSQL = sSQL & " PUI.StaffCode, PUI.LastAccessDay, PUI.UpdateDay, CASE WHEN WA.StaffCode IS NOT NULL THEN '1' ELSE '0' END AS WAFlag" & vbCrLf
			sSQL = sSQL & "FROM P_NaviSearch AS PNS WITH(NOLOCK)" & vbCrLf
			sSQL = sSQL & "INNER JOIN P_UserInfo AS PUI WITH(NOLOCK) "
			sSQL = sSQL & "ON PNS.StaffCode = PUI.StaffCode "
			sSQL = sSQL & "AND PUI.RegistCommit = '1'" & vbCrLf
			sSQL = sSQL & "LEFT JOIN WorkerAlarm AS WA WITH(NOLOCK) "
			sSQL = sSQL & "ON PNS.StaffCode = WA.StaffCode " & vbCrLf
			sSQL = sSQL & sJoin
			sSQL = sSQL & "WHERE PNS.CompanyKbn = '" & CompanyKbn & "'" & vbCrLf & sWhere & vbCrLf
			sSQL = sSQL & "ORDER BY WAFlag DESC, StaffCode DESC" & vbCrLf
			'<2011/05/15 LIS K.Kokubo �T�[�o�ύX�ɂ��MAXDOP�̎w����������� />
			'sSQL = sSQL & "OPTION(MAXDOP 1)"

			GetSQLStaffSearchDetail = "" & _
				"/*�i�r�E���E�ҏڍ׌���*/" & vbCrLf & _
				"/* " & G_USERID & " */" & vbCrLf & _
				"SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf & _
				"EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
			If sDeclare <> "" Then GetSQLStaffSearchDetail = GetSQLStaffSearchDetail & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams
		Else
			'���E�Ҍ����Ώۂ̋��l�[�̏��R�[�h
			sDeclare = sDeclare & "@vOrderCode VARCHAR(8)"
			sParams = sParams & ",@vOrderCode = N'" & OrderCode & "'"
			'�o�^���w��
			If RegistDayFrom <> "" Then
				sDeclare = sDeclare & ",@vRegistDayFrom VARCHAR(8)"
				sParams = sParams & ",@vRegistDayFrom = N'" & RegistDayFrom & "'"
			End If

			sSQL = ""
			sSQL = sSQL & "SELECT"
			If CStr(Top) <> "" Then sSQL = sSQL & " TOP " & Top
			sSQL = sSQL & " PUI.StaffCode "
			sSQL = sSQL & ",PUI.LastAccessDay "
			sSQL = sSQL & ",PUI.UpdateDay "
			sSQL = sSQL & ",CASE WHEN WA.StaffCode IS NOT NULL THEN '1' ELSE '0' END WAFlag" & vbCrLf
			sSQL = sSQL & "FROM P_NaviSearch AS PNS WITH(NOLOCK)" & vbCrLf
			sSQL = sSQL & "INNER JOIN P_UserInfo AS PUI "
			sSQL = sSQL & "ON PNS.StaffCode = PUI.StaffCode "
			sSQL = sSQL & "AND PUI.RegistCommit = '1' "

			If RegistDayFrom <> "" Then
				sSQL = sSQL & vbCrLf & "INNER JOIN P_UserInfo AS PRD ON PNS.StaffCode = PRD.StaffCode AND PRD.RegistDay >= CONVERT(DATETIME, @vRegistDayFrom) "
			End If

			sSQL = sSQL & vbCrLf
			sSQL = sSQL & "INNER JOIN ("
			sSQL = sSQL & "SELECT DISTINCT PHWT.StaffCode "
			sSQL = sSQL & "FROM P_HopeWorkingType AS PHWT "
			sSQL = sSQL & "INNER JOIN C_WorkingType AS CWT "
			sSQL = sSQL & "ON PHWT.WorkingTypeCode = CWT.WorkingTypeCode "
			sSQL = sSQL & "WHERE CWT.OrderCode = @vOrderCode "
			sSQL = sSQL & ") AS PHWT "
			sSQL = sSQL & "ON PNS.StaffCode = PHWT.StaffCode" & vbCrLf
			sSQL = sSQL & "INNER JOIN ("
			sSQL = sSQL & "SELECT DISTINCT PHJT.StaffCode "
			sSQL = sSQL & "FROM P_HopeJobType AS PHJT "
			sSQL = sSQL & "INNER JOIN C_JobType AS CJT "
			sSQL = sSQL & "ON PHJT.JobTypeCode = CJT.JobTypeCode "
			sSQL = sSQL & "WHERE CJT.OrderCode = @vOrderCode "
			sSQL = sSQL & ") AS PHJT "
			sSQL = sSQL & "ON PNS.StaffCode = PHJT.StaffCode" & vbCrLf
			sSQL = sSQL & "INNER JOIN ("
			sSQL = sSQL & "SELECT DISTINCT PHWP.StaffCode "
			sSQL = sSQL & "FROM P_HopeWorkingPlace AS PHWP "
			sSQL = sSQL & "INNER JOIN C_WorkingPlace AS CWP "
			sSQL = sSQL & "ON PHWP.PrefectureCode = CWP.PrefectureCode "
			sSQL = sSQL & "WHERE CWP.OrderCode = @vOrderCode "
			sSQL = sSQL & ") AS PHWP "
			sSQL = sSQL & "ON PNS.StaffCode = PHWP.StaffCode "
			sSQL = sSQL & "LEFT JOIN WorkerAlarm AS WA "
			sSQL = sSQL & "ON PNS.StaffCode = WA.StaffCode" & vbCrLf
			sSQL = sSQL & "WHERE PNS.CompanyKbn = ("
			sSQL = sSQL & "SELECT TOP 1 CMP.CompanyKbn "
			sSQL = sSQL & "FROM CompanyInfo AS CMP "
			sSQL = sSQL & "INNER JOIN C_Info AS CINF "
			sSQL = sSQL & "ON CMP.CompanyCode = CINF.CompanyCode "
			sSQL = sSQL & "AND CINF.OrderCode = @vOrderCode "
			sSQL = sSQL & ")" & vbCrLf
			sSQL = sSQL & "OPTION(MAXDOP 1)" & vbCrLf

			GetSQLStaffSearchDetail = "" & _
				"/*�i�r�E���E�Ҏ�������*/" & vbCrLf & _
				"/* " & G_USERID & " */" & vbCrLf & _
				"SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf & _
				"EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"

			If sDeclare <> "" Then GetSQLStaffSearchDetail = GetSQLStaffSearchDetail & vbCrLf & ",N'" & sDeclare & "'" & vbCrLf & sParams
		End If
'Response.Write GetSQLStaffSearchDetail
	End Function

	'******************************************************************************
	'�T�@�v�F�p�����[�^�����񂩂烁���o�ϐ��̐ݒ�
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
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
		If HopeCity1 <> "" Then HopeCity1 = getURLDecode(HopeCity1,"sjis")
		If HopeCity2 <> "" Then HopeCity2 = getURLDecode(HopeCity2,"sjis")
		If City <> "" Then City = getURLDecode(City,"sjis")
		If SchoolName <> "" Then SchoolName = getURLDecode(SchoolName,"sjis")
		If KeyWord <> "" Then KeyWord = getURLDecode(KeyWord,"sjis")
		If KeyWordHope <> "" Then KeyWordHope = getURLDecode(KeyWordHope,"sjis")
		If KeyWordCareer <> "" Then KeyWordCareer = getURLDecode(KeyWordCareer,"sjis")
		If KeyWordLicense <> "" Then KeyWordLicense = getURLDecode(KeyWordLicense,"sjis")
		If KeyWordPerson <> "" Then KeyWordPerson = getURLDecode(KeyWordPerson,"sjis")
		'</URL�G���R�[�h����Ă��镶������f�R�[�h>

		Call SetNames()
	End Function

	'******************************************************************************
	'�T�@�v�F���ʃ}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get MchCount
		Dim iCnt

		iCnt = 0

		If ChkMch01 = True Then iCnt = iCnt + 1
		If ChkMch02 = True Then iCnt = iCnt + 1
		If ChkMch03 = True Then iCnt = iCnt + 1
		If ChkMch04 = True Then iCnt = iCnt + 1
		If ChkMch05 = True Then iCnt = iCnt + 1
		If ChkMch06 = True Then iCnt = iCnt + 1
		If ChkMch07 = True Then iCnt = iCnt + 1
		If ChkMch08 = True Then iCnt = iCnt + 1
		If ChkMch09 = True Then iCnt = iCnt + 1
		If ChkMch10 = True Then iCnt = iCnt + 1
		If ChkMch11 = True Then iCnt = iCnt + 1
		If ChkMch12 = True Then iCnt = iCnt + 1
		If ChkMch13 = True Then iCnt = iCnt + 1
		If ChkMch14 = True Then iCnt = iCnt + 1
		If ChkMch15 = True Then iCnt = iCnt + 1
		If ChkMch16 = True Then iCnt = iCnt + 1
		If ChkMch17 = True Then iCnt = iCnt + 1
		If ChkMch18 = True Then iCnt = iCnt + 1
		If ChkMch19 = True Then iCnt = iCnt + 1
		If ChkMch20 = True Then iCnt = iCnt + 1
		If ChkMch21 = True Then iCnt = iCnt + 1
		If ChkMch22 = True Then iCnt = iCnt + 1
		If ChkMch23 = True Then iCnt = iCnt + 1

		MchCount = iCnt
	End Property

	'******************************************************************************
	'�T�@�v�F���ʃ}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch01
		ChkMch01 = False
		If Sex <> "" Then ChkMch01 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�N��}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch02
		ChkMch02 = False
		If AgeMin & AgeMax <> "" Then ChkMch02 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�Z��(�s���{��)�}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch03
		ChkMch03 = False
		If PrefectureCode <> "" Then ChkMch03 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�Z��(����,�w)�}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch04
		ChkMch04 = False
		If RailwayLinePrefectureCode <> "" And RailwayLineCode & StationCode <> "" Then ChkMch04 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�Z��(�ߗ׌���)�}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch05
		ChkMch05 = False
		If ZipCode <> "" Then ChkMch05 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�o���w���}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch06
		ChkMch06 = False
		If SchoolTypeCode <> "" Then ChkMch06 = True
	End Property

	'******************************************************************************
	'�T�@�v�F���Ƒ�w�}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch07
		ChkMch07 = False
		If SchoolName <> "" Then ChkMch07 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�w�𕶗���ʃ}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch08
		ChkMch08 = False
		If CourseType <> "" Then ChkMch08 = True
	End Property

	'******************************************************************************
	'�T�@�v�F���L���i�}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch09
		ChkMch09 = False
		If LicenseGroupCode1 & LicenseCategoryCode1 & LicenseCode1 & _
			LicenseGroupCode2 & LicenseCategoryCode2 & LicenseCode2 & _
			LicenseGroupCode3 & LicenseCategoryCode3 & LicenseCode3 <> "" Then ChkMch09 = True
	End Property

	'******************************************************************************
	'�T�@�v�F��w�X�L���}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch10
		ChkMch10 = False
		If LanguageCode & LanguageActionLevel1 & LanguageActionLevel2 & LanguageActionLevel3 <> "" Then ChkMch10 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�n�r�X�L���}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch11
		ChkMch11 = False
		If OSCode1 & OSCode2 <> "" Then ChkMch11 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�A�v���P�[�V�����X�L���}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch12
		ChkMch12 = False
		If ApplicationCode1 & ApplicationCode2 & ApplicationCode3 <> "" Then ChkMch12 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�J������X�L���}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch13
		ChkMch13 = False
		If DevelopmentLanguageCode1 & DevelopmentLanguageCode2 <> "" Then ChkMch13 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�f�[�^�x�[�X�X�L���}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch14
		ChkMch14 = False
		If DatabaseCode1 & DatabaseCode2 <> "" Then ChkMch14 = True
	End Property

	'******************************************************************************
	'�T�@�v�F��]�E��}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch15
		ChkMch15 = False
		If HopeJobTypeCode1 & HopeJobTypeCode2 <> "" Then ChkMch15 = True
	End Property

	'******************************************************************************
	'�T�@�v�F��]�Ǝ�}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch16
		ChkMch16 = False
		If HopeIndustryTypeCode <> "" Then ChkMch16 = True
	End Property

	'******************************************************************************
	'�T�@�v�F��]�Ζ��`�ԃ}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch17
		ChkMch17 = False
		If HopeWorkingTypeCode & WorkingTypeCode1 & WorkingTypeCode2 <> "" Then ChkMch17 = True
	End Property

	'******************************************************************************
	'�T�@�v�F��]�Ζ��n(�s���{��)�}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch18
		ChkMch18 = False
		If HopePrefectureCode <> "" Then ChkMch18 = True
	End Property

	'******************************************************************************
	'�T�@�v�F��]���^�}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch19
		ChkMch19 = False
		If YearlyIncomeMin & YearlyIncomeMax & MonthlyIncomeMin & MonthlyIncomeMax & DailyIncomeMin & DailyIncomeMax & HourlyIncomeMin & HourlyIncomeMax <> "" Then ChkMch19 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�o���E��}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch20
		ChkMch20 = False
		If JobTypeCode1 & JobTypeCode1 <> "" Then ChkMch20 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�o���Ǝ�}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch21
		ChkMch21 = False
		If ExpIndustryTypeCode <> "" Then ChkMch21 = True
	End Property

	'******************************************************************************
	'�T�@�v�F���Љ񐔃}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch22
		ChkMch22 = False
		If CareerCnt <> "" Then ChkMch22 = True
	End Property

	'******************************************************************************
	'�T�@�v�F�t���[���[�h�}�b�`���O�����`�F�b�N
	'���@�l�F
	'���@���F2009/08/04 LIS K.Kokubo �쐬
	'******************************************************************************
	Public Property Get ChkMch23
		ChkMch23 = False
		If KeyWord <> "" Then ChkMch23 = True
	End Property
End Class
%>
