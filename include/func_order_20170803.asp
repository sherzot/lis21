<%
'**********************************************************************************************************************
'�T�@�v�F���l�[�ꗗ�y�[�W /order/order_list_entity.asp
'�@�@�@�F���l�[�ڍ׃y�[�W /order/order_detail_entity.asp
'�@�@�@�F��Ə��y�[�W /order/company_order.asp
'�@�@�@�F��L�y�[�W�ŏo�͗p�̊֐��Q�����̃t�@�C���ɗp�ӂ���B
'�@�@�@�Fmain_pics
'�@�@�@�F�������@�O������@������
'�@�@�@�F�v���O�C���N���[�h
'�@�@�@�F/config/personel.asp
'�@�@�@�F/include/commonfunc.asp
'��@���F�������@���l�[�ꗗ�y�[�W�o�͗p�@������
'�@�@�@�FDspOrderListDetail			�F���l�[�ꗗ�y�[�W�̊e���l�[�P�ʂ�\��
'�@�@�@�FDspOrderListDetail2		�F���l�[�ꗗ�����уo�[�W����1
'�@�@�@�FDspOrderListDetail3		�F���l�[�ꗗ�����уo�[�W����2
'�@�@�@�FDspOrderListDetail4		�F���l�[�ꗗ�����уo�[�W����3 (2�d����Ή�����)
'�@�@�@�FDspPageControl				�F���l�[�ꗗ�y�[�W�̃y�[�W�R���g���[��
'�@�@�@�F
'�@�@�@�F�������@��Ə��y�[�W�o�͗p�@������
'�@�@�@�FDspCompanyInfo				�F��Ə��̊�{�����o��
'       :DspCompanyInfoNEO          �F��Ə��̊�{�����o��(NEO�p)
'�@�@�@�FDspCompanyPR				�F��Ə��̂o�q�����o��
'�@�@�@�F
'�@�@�@�F�������@���l�[�ڍ׃y�[�W�o�͗p�@������
'�@�@�@�FDspLisOrderCompanyInfo		�F���l�[�ڍ׃y�[�W�̃��X�̏Љ��E�h�����Ə����o��
'�@�@�@�FDspTempOrderCompanyInfo	�F���l�[�ڍ׃y�[�W�̔h����Ƃ̔h�����Ə����o��
'�@�@�@�FDspOrderControlButton		�F���l�[�ڍ׃y�[�W�̃R���g���[���{�^���i���O�C���ς݃��[�U�p�j
'�@�@�@�FJSOrderControlButton		�F���l�[�ڍ׃y�[�W�̃R���g���[���{�^���ŗ��p����javascript�̏o��
'�@�@�@�FFrmOrderControlButton		�F���l�[�ڍ׃y�[�W�̃R���g���[���{�^���ŗ��p����FORM�f�[�^�̏o��
'�@�@�@�FDspOrderCompanyName		�F���l�[�ڍ׃y�[�W�̊�Ɩ����o��
'�@�@�@�FDspOrderShowTypeSwitch		�F���l�[�ڍ׃y�[�W�̉�Џ��E�E����E�C���^�r���[�؂�ւ��{�^���ƎQ�Ɖ񐔂��o��
'�@�@�@�FDspOrderCatchCopy			�F���l�[�ڍ׃y�[�W�̃L���b�`�R�s�[�����i�傫���摜�Ȃǁj���o��
'�@�@�@�FDspOrderCatchCopy_OldPlan		�F���l�[�ڍ׃y�[�W�̉ߋ����l�o�i�[���o��
'�@�@�@�FDspOrderFreePR				�F���l�[�ڍ׃y�[�W�̃t���[�o�q���o��
'�@�@�@�FDspOrderPictureNow			�F���l�[�ڍ׃y�[�W�̏������摜���o��
'�@�@�@�FDspOrderBackGround			�F���l�[�ڍ׃y�[�W�̗̍p�̔w�i���o��
'�@�@�@�FDspBusiness				�F���l�[�ڍ׃y�[�W�̋Ɩ����e���o��
'�@�@�@�FDspCondition				�F���l�[�ڍ׃y�[�W�̋Ζ��������o��
'�@�@�@�FDspNeedCondition			�F���l�[�ڍ׃y�[�W�̕K�v�������o��
'�@�@�@�FDspHowToEntry				�F���l�[�ڍ׃y�[�W�̉�������o��
'�@�@�@�FDspContact					�F���l�[�ڍ׃y�[�W�̒S���ҘA������o��
'�@�@�@�FDspElderInterview			�F���l�[�ڍ׃y�[�W�̐�y�C���^�r���[���o��
'�@�@�@�FDspConsultantComment		�F���X�̈Č��S���ҁA�R���T���������o��
'�@�@�@�FDspNewMail					�F���l�[�ڍ׃y�[�W�̍ŐV�̑��M�ς݃��[�����o��
'�@�@�@�FGetWorkingType				�F���l�[�ڍ׃y�[�W�̋Ζ��`�ԕ���
'�@�@�@�FGetJobType					�F���l�[�ڍ׃y�[�W�̐E�핔��
'�@�@�@�FGetWorkingTime				�F���l�[�ڍ׃y�[�W�̋Ζ��`�ԕ���
'�@�@�@�FGetNearbyStation			�F���l�[�ڍ׃y�[�W�̍Ŋ�w����
'�@�@�@�FGetNearbyRailway			�F���l�[�ڍ׃y�[�W�̍Ŋ񉈐�����
'�@�@�@�FGetSkill					�F���l�[�ڍ׃y�[�W�̃X�L������
'�@�@�@�FGetLicense					�F���l�[�ڍ׃y�[�W�̎��i����
'�@�@�@�FGetOrderNote				�F���l�[�ڍ׃y�[�W�̎��i����
'�@�@�@�FGetOrderTitle				�F���l�[�ڍ׃y�[�W�̃^�C�g���ƃf�B�X�N���v�V�������擾
'�@�@�@�FGetSkillList				�F���l�[�ڍ׃y�[�W�̃X�L���̊e���ڕ\��
'�@�@�@�F
'�@�@�@�F�������@���R�����h�@������
'�@�@�@�FDspRecommendOrderList		�F���R�����h���d�����ꗗ�o��
'�@�@�@�FGetRecommendValues			�F���R�����h�̋��l�[�ꗗ�́A���l�[���̊e���ځi�E��A��Ɩ��Ȃǁj���擾
'�@�@�@�F
'�@�@�@�F�������@���l�[�ڍ׃y�[�W�`�F�b�N�p�@������
'�@�@�@�FChkMyOrder					�F���Ћ��l�[���ۂ����`�F�b�N���� ["0"]���Ћ��l�[�ȊO ["1"]���Ћ��l�[
'�@�@�@�F
'�@�@�@�F�������@�f�ڏ�ԕύX�E���l�[�폜�@������
'�@�@�@�FUpdMyOrderPublicFlag		�F���Ћ��l�[�̌f�ڏ�Ԃ�ύX����
'�@�@�@�FDelMyOrder					�F���Ћ��l�[���폜����
'�@�@�@�F
'�@�@�@�F�������@���ʗ��p�@������
'�@�@�@�FGetImgOrderSpeciality		�F���l�[�̓���
'�@�@�@�F
'�@�@�@�F�������@���������Ƃ����ƃi�r�ŕ\�����قȂ镔���p�@������
'�@�@�@�FDspTopRegButton			�F�����ƃi�r�̋��l�[�ڍ׃y�[�W�̏㕔�ɒu���A���O�C���U���{�^���B
'�@�@�@�FDspTopRegButtonResume		�F���������̋��l�[�ڍ׃y�[�W�̏㕔�ɒu���A���O�C���U���{�^���B
'�@�@�@�FDspBottomRegButton			�F�����ƃi�r�̋��l�[�ڍ׃y�[�W�̉����ɒu���A���O�C���U���{�^���B
'�@�@�@�FDspBottomRegButton_OldPlan		�F�����ƃi�r�̋��l�[�ڍ׃y�[�W�̉����ɒu���A���O�C���U���{�^���B
'�@�@�@�FDspBottomRegButtonResume	�F���������̋��l�[�ڍ׃y�[�W�̉����ɒu���A���O�C���U���{�^���B
'�@�@�@�F
'�@�@�@�F�������@���l�[�ڍ׃A�N�Z�X���̐���@������
'�@�@�@�FMailMagazineAccess			�F�V�����l���[������̃A�N�Z�X���̃��O��������
'�@�@�@�FMailMagazineDelivery		�F���l�����}�K����̃A�N�Z�X���̃��O��������
'�@�@�@�FAccessHistoryOrder			�F���Ճ��O�̏�������
'�@�@�@�FAccessCountUp				�F�A�N�Z�X�񐔂̃J�E���g�A�b�v
'�@�@�@�FPVCountUp					�F���l�[�̓��ʂo�u�̃J�E���g�A�b�v
'**********************************************************************************************************************

'******************************************************************************
'�T�@�v�F���l�[�ꗗ�y�[�W�̊e���l�[���ڂ�\��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_SearchOrder or ���l�[�ڍ׌���SQL �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvMyOrder		�F���p�����[�U�̎��Ћ��l�[���ۂ� ["1"]���Ћ��l�[ ["0"]���Ћ��l�[�łȂ�
'�g�p���Forder/order_list_entity.asp
'���@�l�F
'���@���F2006/05/13 LIS K.Kokubo �쐬
'�@�@�@�F2007/11/22 LIS K.Kokubo up_SearchOrder��K�v�ŏ����̂��̂���������Ă���悤�ɂ������Ƃɂ��ύX�Bup_DtlOrder����f�[�^���擾�B
'�@�@�@�F2008/03/04 LIS K.Kokubo �f�ڏI������[RiyoToDate]��[DspPublicLimitDay]�ɕύX
'�@�@�@�F2008/03/11 LIS K.Kokubo �g�b�v�C���^�r���[�ւ̃����N���o��
'�@�@�@�F2008/08/01 LIS K.Kokubo �v�o�����[�̃����N���o��
'�@�@�@�F2008/08/19 LIS �� �����t���O�̒ǉ��ƃt���b�N�X�ړ�
'�@�@�@�F2008/10/20 LIS K.Kokubo �Ζ��n�������ɂ��C��
'�@�@�@�F2010/01/28 LIS K.Kokubo ���^�̋L�ڂ������ꍇ�͋��^�̍��ڂ�\�����Ȃ��iFC�ESOHO�Ή��j
'******************************************************************************
Function DspOrderListDetail(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder)
	Const PICSIZEW = 240
	Const PICSIZEH = 180
	Const PICSIZESUBW = 72
	Const PICSIZESUBH = 56

	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim dbOrderCode			'���R�[�h
	Dim dbCompanyCode		'��ƃR�[�h
	Dim sOrderType			'�󒍎��
	Dim sPlanType			'���C�Z���X�v�������
	Dim iImageLimit			'�ʐ^�f�ڐ�������
	Dim sTitleJobName		'�E��
	Dim sTitleCompanyName	'��Ж�
	Dim sImgMail			'���M�ς݃��[���摜
	Dim sImgOrderState		'��ԉ摜 HOT,�V��,���o��OK,UI�^�[��,��w,�x��120��,�t���b�N�X
	Dim sCatchCopy			'�L���b�`�R�s�[
	Dim flgImg				'�摜�̗L���t���O(�摜�̗L���Ń��C�A�E�g���ω�) [True]�L [False]��
	Dim sImgMain			'�傫���摜
	Dim sImgSub				'�������摜
	Dim sImg1,sImg2,sImg3,sImg4	'�摜URL
	Dim sBusinessDetail		'�S���Ɩ�
	Dim sWorkingType		'�Ζ��`��
	Dim sWorkingPlace		'�Ζ��n �s���{��+�s��S
	Dim sProgress			'���l�[�R����
	Dim sPublicDay			'�f�ړ�
	Dim sPublicListDsp		'�f�ڔ�f�� ���X�g�{�b�N�X�\���X�^�C�� [style="display:none;"]
	Dim sPublicFlag1		'�f�� selected
	Dim sPublicFlag0		'��f�� selected
	Dim sCompanyPictureFlag	'��Ǝʐ^�t���O ["1"]�L ["0"]��
	Dim sRegistDay			'�o�^��
	Dim sPublishLimitStr	'���l�[�f�ڏI����
	Dim sStationName		'�w��
	Dim sYearlyIncomeMin	'�N������
	Dim sYearlyIncomeMax	'�N�����
	Dim sMonthlyIncomeMin	'��������
	Dim sMonthlyIncomeMax	'�������
	Dim sDailyIncomeMin		'��������
	Dim sDailyIncomeMax		'�������
	Dim sHourlyIncomeMin	'��������
	Dim sHourlyIncomeMax	'�������
	Dim dbTopInterviewFlag	'�g�b�v�C���^�r���[�t���O
	Dim dbWValueURL			'�v�o�����[�̂t�q�k

	Dim sYearlyIncome		'�N���\���p
	Dim sDailyIncome		'�����\���p
	Dim sMonthlyIncome		'�����\���p
	Dim sHourlyIncome		'�����\���p
	'��]�Ζ��`�ԁE��]�Ζ��n�A�C�R���@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��
	Dim sWorkingCode
	Dim sWorkingName
	Dim dbWorkingPlacePrefectureCode
	Dim dbWorkingPlacePrefectureName
	Dim dbWorkingPlaceCity
	Dim sBiz
	Dim sBizName1
	Dim sBizName2
	Dim sBizName3
	Dim sBizName4
	Dim sBizPercentage1
	Dim sBizPercentage2
	Dim sBizPercentage3
	Dim sBizPercentage4
	Dim flgBusiness
	Dim idx
	'HTTP�N���X�ύX�p�ϐ�
	Dim HtmlClassName
	Dim HtmlWorkingType
    Dim HimlOiwai

	
	
	
	
	
	Dim flgAddWatchList
	Dim iMailTemplateCnt	'���[���e���v���[�g�̌���
	Dim sAncMT				'���[���e���v���[�g�ւ̃����N
	Dim sOrderCode
	
	
	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	flgAddWatchList = False
	sSQL = "EXEC up_ChkWatchListExists_Staff '" & vUserID & "', '" & sOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("ExistsFlag") = "1" Then flgAddWatchList = True
	End If
	Call RSClose(oRS)
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************
	
	
	
	
	
	
	
'	Dim qsOrderCode				'�I�[�_�[�R�[�h(�󒍕\�ԍ�)
'	Dim iDetail				'���l�[�ڍׂ���̃t���O
'	
'	qsOrderCode = GetForm("ordercode", 2)
'	iDetail = GetForm("Detail", 2)
'
'	'******************************************************************************
'	'��ƃR�[�h start
'	'------------------------------------------------------------------------------
'	flgAddWatchList = False
'	
'	sSQL = "EXEC up_ChkWatchListExists_Staff '" & G_USERID & "', '" & qsOrderCode & "';"
'	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
'	If GetRSState(oRS) = True Then
'		If oRS.Collect("ExistsFlag") = "1" Then flgAddWatchList = True
'	End If
'	Call RSClose(oRS)
'	'------------------------------------------------------------------------------
'	'��ƃR�[�h end
'	'******************************************************************************
	
	
	
	
	
	
	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")

	DspOrderListDetail = False

	If G_USEFLAG = "0" And vMyOrder = "1" And G_OLDAPPLICATIONCODE <> "" Then
		sSQL = "EXEC up_DtlOrder_NEO '" & rRS.Collect("OrderCode") & "', '" & G_OLDAPPLICATIONCODE & "';"
	Else
		sSQL = "EXEC up_DtlOrder_NEO '" & rRS.Collect("OrderCode") & "', '';"
	End If

	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	dbCompanyCode = oRS.Collect("CompanyCode")
	sOrderType = ChkStr(oRS.Collect("OrderType"))
	sPlanType = ChkStr(oRS.Collect("PlanTypeName"))
	iImageLimit = oRS.Collect("ImageLimit")
    '���j�����ݒ�
    HimlOiwai = oRS.Collect("CongratulationPrice")

	'**************************************************************************
	'�E��^��Ж� start
	'--------------------------------------------------------------------------
	sTitleCompanyName = ""
	'STEP1�F��̓I�E�햼�擾
	If oRS.Collect("JobTypeDetail") <> "" Then
		If Len(oRS.Collect("JobTypeDetail")) >= 50 Then
			sTitleJobName = Left(oRS.Collect("JobTypeDetail"), 50)
		Else
			sTitleJobName = oRS.Collect("JobTypeDetail")
		End If
	End If

	'STEP2�F��̓I�E�햼������΁^��ǉ�
	'If sTitleCompanyName <> "" Then sTitleCompanyName = sTitleCompanyName & "�^"
	'STEP3�F��Ɩ��擾
	If oRS.Collect("CompanySpeciality") <>"" THEN 
			sTitleCompanyName = sTitleCompanyName & oRS.Collect("CompanySpeciality")
	Else
		If oRS.Collect("Companykbn") ="4" Then
			sTitleCompanyName = sTitleCompanyName & oRS.Collect("CompanyName")
		ElseIf oRS.Collect("OrderType") > "0" then
				sTitleCompanyName = sTitleCompanyName & "���X�������"
		Else
				sTitleCompanyName = sTitleCompanyName & oRS.Collect("CompanyName")
		End If
	End If
	'--------------------------------------------------------------------------
	'�E��^��Ж� end
	'**************************************************************************

	'******************************************************************************
	'���^ start�@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��
	'------------------------------------------------------------------------------
	'�N��
	sYearlyIncomeMin = ChkStr(oRS.Collect("YearlyIncomeMin"))
	sYearlyIncomeMax = ChkStr(oRS.Collect("YearlyIncomeMax"))
	If sYearlyIncomeMin = "0" Then sYearlyIncomeMin = ""
	If sYearlyIncomeMax = "0" Then sYearlyIncomeMax = ""
	If sYearlyIncomeMin <> "" Then sYearlyIncomeMin = GetJapaneseYen(sYearlyIncomeMin)
	If sYearlyIncomeMax <> "" Then sYearlyIncomeMax = GetJapaneseYen(sYearlyIncomeMax)
	If sYearlyIncomeMin & sYearlyIncomeMax <> "" Then
		If sYearlyIncomeMin <> "" Then sYearlyIncome = sYearlyIncome & sYearlyIncomeMin
		sYearlyIncome = sYearlyIncome & "&nbsp;�`&nbsp;"
		If sYearlyIncomeMax <> "" Then sYearlyIncome = sYearlyIncome & sYearlyIncomeMax
	End If
	'����
	sMonthlyIncomeMin = ChkStr(oRS.Collect("MonthlyIncomeMin"))
	sMonthlyIncomeMax = ChkStr(oRS.Collect("MonthlyIncomeMax"))
	If sMonthlyIncomeMin = "0" Then sMonthlyIncomeMin = ""
	If sMonthlyIncomeMax = "0" Then sMonthlyIncomeMax = ""
	If sMonthlyIncomeMin <> "" Then sMonthlyIncomeMin = GetJapaneseYen(sMonthlyIncomeMin)
	If sMonthlyIncomeMax <> "" Then sMonthlyIncomeMax = GetJapaneseYen(sMonthlyIncomeMax)
	If sMonthlyIncomeMin & sMonthlyIncomeMax <> "" Then
		If sMonthlyIncomeMin <> "" Then sMonthlyIncome = sMonthlyIncome & sMonthlyIncomeMin
		sMonthlyIncome = sMonthlyIncome & "&nbsp;�`&nbsp;"
		If sMonthlyIncomeMax <> "" Then sMonthlyIncome = sMonthlyIncome & sMonthlyIncomeMax
	End If
	'����
	sDailyIncomeMin = ChkStr(oRS.Collect("DailyIncomeMin"))
	sDailyIncomeMax = ChkStr(oRS.Collect("DailyIncomeMax"))
	If sDailyIncomeMin = "0" Then sDailyIncomeMin = ""
	If sDailyIncomeMax = "0" Then sDailyIncomeMax = ""
	If sDailyIncomeMin <> "" Then sDailyIncomeMin = GetJapaneseYen(sDailyIncomeMin)
	If sDailyIncomeMax <> "" Then sDailyIncomeMax = GetJapaneseYen(sDailyIncomeMax)
	If sDailyIncomeMin & sDailyIncomeMax <> "" Then
		If sDailyIncomeMin <> "" Then sDailyIncome = sDailyIncome & sDailyIncomeMin
		sDailyIncome = sDailyIncome & "&nbsp;�`&nbsp;"
		If sDailyIncomeMax <> "" Then sDailyIncome = sDailyIncome & sDailyIncomeMax
	End If
	'����
	sHourlyIncomeMin = ChkStr(oRS.Collect("HourlyIncomeMin"))
	sHourlyIncomeMax = ChkStr(oRS.Collect("HourlyIncomeMax"))
	If sHourlyIncomeMin = "0" Then sHourlyIncomeMin = ""
	If sHourlyIncomeMax = "0" Then sHourlyIncomeMax = ""
	If sHourlyIncomeMin <> "" Then sHourlyIncomeMin = GetJapaneseYen(sHourlyIncomeMin)
	If sHourlyIncomeMax <> "" Then sHourlyIncomeMax = GetJapaneseYen(sHourlyIncomeMax)
	If sHourlyIncomeMin & sHourlyIncomeMax <> "" Then
		If sHourlyIncomeMin <> "" Then sHourlyIncome = sHourlyIncome & sHourlyIncomeMin
		sHourlyIncome = sHourlyIncome & "&nbsp;�`&nbsp;"
		If sHourlyIncomeMax <> "" Then sHourlyIncome = sHourlyIncome & sHourlyIncomeMax
	End If

	'------------------------------------------------------------------------------
	'���^ end
	'******************************************************************************

	'******************************************************************************
	'�Ŋ�w start�@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��
	'2008/10/22 LIS K.Kokubo �Ζ��n�������ɂ��\���ʂ������鋰�ꂪ���邽�߂ɔ�\���ɁB
	'------------------------------------------------------------------------------
	'sStationName = ""
	'sSQL = "sp_GetDataNearbyStation '" & dbOrderCode & "'"
	'flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	'If GetRSState(oRS2) = True Then
	'	sStationName ="�y" & sStationName & GetStrNearbyStation(oRS2.Collect("StationName"), "", "") & "�z"
	'End If
	'------------------------------------------------------------------------------
	'�Ŋ�w end
	'******************************************************************************

	'**************************************************************************
	'���[�����M�ς݊m�F start
	'--------------------------------------------------------------------------
	If vUserType = "staff" Then
		sSQL = "up_DtlMailHistory_Order '" & vUserID & "', '" & dbOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			sImgMail = "<img src=""/img/s_contact.gif"" alt=""���[�����M�ς�"">"
		End If
		Call RSClose(oRS2)
	End If

	'�u���l�[�����[�����M�v�̃����N�ɂԂ���Ȃ��悤�ɐE�햼�����(2007/08/01 T.Sotome�ǉ�)
	If LenByte(sTitleCompanyName) > 72 Then
		sTitleCompanyName = LeftByte(sTitleCompanyName, 70) & "..."
	End If
	'�u�E�H�b�`���X�g�֕ۑ��v�̃����N�ɂԂ���Ȃ��悤�ɐE�햼�����(2007/06/26 T.Sotome�ǉ�)
	If sImgMail = "" Then
		If LenByte(sTitleJobName) > 46 Then
			sTitleJobName = LeftByte(sTitleJobName, 44) & "..."
		End If
	Else
		If LenByte(sTitleJobName) > 36 Then
			sTitleJobName = LeftByte(sTitleJobName, 34) & "..."
		End If
	End If

	'--------------------------------------------------------------------------
	'���[�����M�ς݊m�F end
	'**************************************************************************

	'**************************************************************************
	'���img start
	'--------------------------------------------------------------------------
	sImgOrderState = GetImgOrderSpeciality(rDB, oRS)
	'--------------------------------------------------------------------------
	'���img end
	'**************************************************************************

	'**************************************************************************
	'�L���b�`�R�s�[ start
	'--------------------------------------------------------------------------
	sCatchCopy = ""
	sCatchCopy = chkstr(oRS.Collect("CatchCopy"))
	'--------------------------------------------------------------------------
	'�L���b�`�R�s�[ end
	'**************************************************************************

	'**************************************************************************
	'�摜 start
	'--------------------------------------------------------------------------
	flgImg = False
	If sOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS2,sSQL,sError)
		If GetRSState(oRS2) = True Then
			If ChkStr(oRS2.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS2.Collect("PicNo1")
			End If
			If ChkStr(oRS2.Collect("PicNo2")) <> "" Then
				sImg2 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS2.Collect("PicNo2")
			End If
			If ChkStr(oRS2.Collect("PicNo3")) <> "" Then
				sImg3 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS2.Collect("PicNo3")
			End If
			If ChkStr(oRS2.Collect("PicNo4")) <> "" Then
				sImg4 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS2.Collect("PicNo4")
			End If
		End If
		Call RSClose(oRS2)
	ElseIf iImageLimit > 0 Then
		sCompanyPictureFlag = ChkStr(oRS.Collect("CompanyPictureFlag"))

		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & oRS.Collect("OrderCode") & "', 'orderpicture'"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			If ChkStr(oRS2.Collect("OptionNo1")) <> "" Or (sOrderType = "0" And sCompanyPictureFlag = "1") Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo1")
			End If

			If sPlanType = "platinum" Or sPlanType = "old" Or iImageLimit > 1 Then
				If ChkStr(oRS2.Collect("OptionNo2")) <> "" Then
					sImg2 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo2")
				End If
				If ChkStr(oRS2.Collect("OptionNo3")) <> "" Then
					sImg3 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo3")
				End If
				If ChkStr(oRS2.Collect("OptionNo4")) <> "" Then
					sImg4 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo4")
				End If
			End If
		Else
			If sCompanyPictureFlag = "1" And sOrderType = "0" Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If

		Call RSClose(oRS2)
	End If

	If sImg1 & sImg2 & sImg3 & sImg4 <> "" Then flgImg = True

	If sImg1 <> "" Then
		sImgMain = "background:url("& sImg1 &") no-repeat;"
	End If

	If sImg2 <> "" Then
		sImgSub = sImgSub & "<div class=""sub_img"">" & _
			"<img src=""" & sImg2 & """><br>"
		sImgSub = sImgSub & "</div>"
		flgImg = True
	End If
	If sImg3 <> "" Then
		sImgSub = sImgSub & "<div class=""sub_img"" style=""margin-top: 5px;"">" & _
			"<img src=""" & sImg3 & """><br>"
		sImgSub = sImgSub & "</div>"
		flgImg = True
	End If
	'If sImg4 <> "" Then
	'	sImgSub = sImgSub & "<div class=""sub_img"">" & _
	'		"<img src=""" & sImg4 & """><br>"
	'	sImgSub = sImgSub & "</div>"
	'End If

	If sImgSub <> "" Then sImgSub =  sImgSub 
	'--------------------------------------------------------------------------
	'�摜 end
	'**************************************************************************

	'**************************************************************************
	'�S���Ɩ� start
	'--------------------------------------------------------------------------
	If flgImg = True Then
		'�摜���L��ꍇ�͕��͂�Z�߂ɃJ�b�g
		sBusinessDetail = Left(oRS.Collect("BusinessDetail"),300) & "&nbsp;"
		If Len(sBusinessDetail) > 300 Then sBusinessDetail = sBusinessDetail & "..."
	Else
		'�摜�������ꍇ�͕��͂𒷂߂ɃJ�b�g
		sBusinessDetail = Left(oRS.Collect("BusinessDetail"),465) & "&nbsp;"
		If Len(sBusinessDetail) > 465 Then sBusinessDetail = sBusinessDetail & "..."
	End If
	'--------------------------------------------------------------------------
	'�S���Ɩ� end
	'**************************************************************************

	'**************************************************************************
	'�Ζ��`�� start
	'--------------------------------------------------------------------------
	sWorkingType = ""
	sSQL = "sp_GetDataWorkingType '" & oRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Dim Counter
	Counter = 1
	Do While GetRSState(oRS2) = True
		if Counter = 1 Then
			HtmlWorkingType = oRS2.Collect("WorkingTypeCode")
			Counter = Counter + 1
		End If
		sWorkingType = sWorkingType & oRS2.Collect("WorkingTypeName")
		If (oRS.Collect("OrderType") ="0" And oRS.Collect("Companykbn") = "2") Or oRS.Collect("OrderType") ="1" Or oRS.Collect("OrderType") ="2" Or oRS.Collect("OrderType") ="3" Then
			Select Case oRS2.Collect("WorkingTypeCode")
				Case "001": sWorkingType = sWorkingType & "<span class=""smartNone"">�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/koyoukeitai_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")' class=""haken_tr"">�h���Ƃ�</a>�z</span>" 
				Case "002","003": sWorkingType = sWorkingType & "<span class=""smartNone"">�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")' class=""shokai_tr"">�l�ޏЉ�Ƃ�</a>�z</span>" 
				Case "004": sWorkingType = sWorkingType & "<span class=""smartNone"">�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/syoukaiyotei_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>�Љ�\��h���Ƃ�</a>�z</span>" 
			End Select
		End If
		sWorkingType = sWorkingType & "<br>"
		oRS2.MoveNext
	Loop
	Call RSClose(oRS2)
	'--------------------------------------------------------------------------
	'�Ζ��`�� end
	'**************************************************************************

	'**************************************************************************
	'�Ζ��n start
	'--------------------------------------------------------------------------
	idx = 0
	sWorkingPlace = ""
	sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True And idx < 20
		dbWorkingPlacePrefectureCode = ChkStr(oRS2.Collect("WorkingPlacePrefectureCode"))
		dbWorkingPlacePrefectureName = ChkStr(oRS2.Collect("WorkingPlacePrefectureName"))
		dbWorkingPlaceCity = ChkStr(oRS2.Collect("WorkingPlaceCity"))
		'<�Ζ��n�A�C�R��>
		If InStr(sImgOrderState, "/icon_p" & dbWorkingPlacePrefectureCode & ".gif") = 0 Then
			'�����s���{���A�C�R���͏o���Ȃ��I
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_p" & dbWorkingPlacePrefectureCode & ".gif"" alt=""" & dbWorkingPlacePrefectureName & """ width=""50"" height=""15"">&nbsp;"
		End If
		'</�Ζ��n�A�C�R��>

		'<�Ζ��n>
		If sWorkingPlace <> "" Then sWorkingPlace = sWorkingPlace & "/"
		sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
		'</�Ζ��n>

		oRS2.MoveNext
		idx = idx + 1
	Loop
	If oRS2.RecordCount > 20 Then sWorkingPlace = sWorkingPlace & "&nbsp;...��"
	Call RSClose(oRS2)
	'--------------------------------------------------------------------------
	'�Ζ��n end
	'**************************************************************************

	'**************************************************************************
	'�f�ڏ�ԃ��X�g�{�b�N�X start
	'--------------------------------------------------------------------------
	sPublicFlag1 = ""
	sPublicFlag0 = ""
	If oRS.Collect("PublicFlag") = "1" Then
		sPublicFlag1 = " selected"
	Else
		sPublicFlag0 = " selected"
	End If
	'--------------------------------------------------------------------------
	'�f�ڏ�ԃ��X�g�{�b�N�X start
	'**************************************************************************

	'**************************************************************************
	'�R���̐i�� start
	'--------------------------------------------------------------------------
	sProgress = ""
	sPublicListDsp = ""
	sPublicDay = ""

	'�R����
	If oRS.Collect("PermitFlag") = "0" Then
		'���X���R��
		sProgress = "���X�R����"
		sPublicListDsp = "style=""display:none;"""
	ElseIf oRS.Collect("PermitFlag") = "1" Then
		'���X����
		If oRS.Collect("PublicFlag") = "0" Then
			sProgress = "���X����(��f��)"
		Else
			sProgress = "�f�ڒ�"
		End If
	Else
		'�ȊO
		If oRS.Collect("PublicFlag") = "1" And oRS.Collect("PermitFlag") = "1" Then
			sProgress = "�f��"
		Else
			sProgress = "��f��"
		End If
		sPublicListDsp = "style=""display:none;"""
	End If

	'�f�ړ�
	sPublicDay = GetDateStr(oRS.Collect("PublicDay"), "/")
	If oRS.Collect("PermitFlag") = "1" And oRS.Collect("PublicDay") > Date Then
		sPublicDay = "<span style=""color:#ff0000;"">��(" & sPublicDay & ")</span>"
		sPublicListDsp = "style=""display:none;"""
	End If
	'--------------------------------------------------------------------------
	'�R���̐i�� end
	'**************************************************************************

	'**************************************************************************
	'�o�^�� start
	'--------------------------------------------------------------------------
	sRegistDay = GetDateStr(oRS.Collect("RegistDay"), "/")
	'--------------------------------------------------------------------------
	'�o�^�� end
	'**************************************************************************

	'******************************************************************************
	'���l�[�f�ڊ��� start
	'------------------------------------------------------------------------------
	'��ƃ��O�C�����ȊO�̂Ƃ��Ɍf�ڊ�����\��
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(oRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(oRS.Collect("PublicLimitDay"))
	End If

	If sPublishLimitStr = "" Then
		If oRS.Collect("NowPublicFlag") = "0" Then
			'���C�Z���X�؂�̂Ƃ���"�f�ڏI��"�ƕ\��
			sPublishLimitStr = "�f�ڏI��"
		Else
			sPublishLimitStr = "�펞��W��"
		End If
	End If

    '<�����������@�\�Ή�>
    '2016/04/01 �r�c���C
    If sPublishLimitStr = "9999/12/31" Then
        '�������̏ꍇ�́A�f�ڊ����Ɍ������w��B�X�V���Ɍ������w��B
        sPublishLimitStr = DateSerial(Year(Date()), Month(Date()) + 1, 0)
        'sUpdateDay       = DateSerial(Year(Date()), Month(Date()), 1)
    End If
    '</�����������@�\�Ή�>


	sPublishLimitStr = sPublishLimitStr & "&nbsp;"
	'------------------------------------------------------------------------------
	'���l�[�f�ڊ��� end
	'******************************************************************************

	'******************************************************************************
	'�d���̊��� start�@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��
	'------------------------------------------------------------------------------
	If sPlanType = "platinum" Or sPlanType = "gold" Or sPlanType = "old" Then
		sBiz = ""
		sBizName1 = ""
		sBizName2 = ""
		sBizName3 = ""
		sBizName4 = ""
		sBizPercentage1 = ""
		sBizPercentage2 = ""
		sBizPercentage3 = ""
		sBizPercentage4 = ""

		sBizName1 = ChkStr(oRS.Collect("BizName1"))
		sBizName2 = ChkStr(oRS.Collect("BizName2"))
		sBizName3 = ChkStr(oRS.Collect("BizName3"))
		sBizName4 = ChkStr(oRS.Collect("BizName4"))
		sBizPercentage1 = ChkStr(oRS.Collect("BizPercentage1"))
		sBizPercentage2 = ChkStr(oRS.Collect("BizPercentage2"))
		sBizPercentage3 = ChkStr(oRS.Collect("BizPercentage3"))
		sBizPercentage4 = ChkStr(oRS.Collect("BizPercentage4"))
		If sBizPercentage1 = "" Then sBizPercentage1 = "0"
		If sBizPercentage2 = "" Then sBizPercentage2 = "0"
		If sBizPercentage3 = "" Then sBizPercentage3 = "0"
		If sBizPercentage4 = "" Then sBizPercentage4 = "0"

		If Len(sBizName1) >= 17 Then sBizName1 = Left(sBizName1,17) & "..."
		If Len(sBizName2) >= 17 Then sBizName2 = Left(sBizName2,17) & "..."
		If Len(sBizName3) >= 17 Then sBizName3 = Left(sBizName3,17) & "..."
		If Len(sBizName4) >= 17 Then sBizName4 = Left(sBizName4,17) & "..."

		If sBizName1 & sBizName2 & sBizName3 & sBizName4 <> "" Then
			If sBizName1 <> "" And sBizPercentage1 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName1 & "</td><td class=""biz2"">" & sBizPercentage1 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage1) * 3 & """ height=""20""></td></tr>"
			If sBizName2 <> "" And sBizPercentage2 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName2 & "</td><td class=""biz2"">" & sBizPercentage2 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage2) * 3 & """ height=""20""></td></tr>"
			If sBizName3 <> "" And sBizPercentage3 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName3 & "</td><td class=""biz2"">" & sBizPercentage3 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage3) * 3 & """ height=""20""></td></tr>"
			If sBizName4 <> "" And sBizPercentage4 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName4 & "</td><td class=""biz2"">" & sBizPercentage4 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage4) * 3 & """ height=""20""></td></tr>"
			sBiz = "<table>" & sBiz & "</table>"
			flgBusiness = True
		End If
	End If
	'------------------------------------------------------------------------------
	'�d���̊��� end
	'******************************************************************************

	'******************************************************************************
	'�g�b�v�C���^�r���[ start
	'------------------------------------------------------------------------------
	dbTopInterviewFlag = oRS.Collect("TopInterviewFlag")
	'------------------------------------------------------------------------------
	'�g�b�v�C���^�r���[ end
	'******************************************************************************

	'******************************************************************************
	'�v�o�����[�t�q�k start
	'------------------------------------------------------------------------------
	dbWValueURL = ChkStr(oRS.Collect("WValueURL"))
	'------------------------------------------------------------------------------
	'�v�o�����[�t�q�k end
	'******************************************************************************

	Response.Write "<input type=""hidden"" name=""CONF_OrderCodes"" value=""" & oRS.Collect("OrderCode") & """>"
	
    '�N���X�ύX
    If sOrderType = "2" Then
    '�Љ�̎� 
        HtmlClassName = "neo_shokai"
    Elseif sOrderType = "1" Then
    '�h���̎�
        HtmlClassName = "neo_haken"
    Elseif sOrderType = "3" Then
    '�Љ�\��h���̎�
        HtmlClassName = "neo_ttp"
    Elseif sOrderType = "0" Then
    '�L���̂Ƃ�
        if HtmlWorkingType = "005" Then
            HtmlClassName = "neo_beit"
		Elseif HtmlWorkingType = "006" Then
			 HtmlClassName = "neo_soho"
        Else
            HtmlClassName = "neo_shain"
        End if
    End if


	'�L���ꗗ

	If oRS.Collect("CompanyCode") = vUserID And vMyOrder = "1" And G_USEFLAG = "1" Then

		%>
<div class="my_order">
<div>
<span>���R�[�h</span>(<%= oRS.Collect("OrderCode") %>)
</div>
<div>
<span>��� </span><%= sProgress %>
</div>
<div>
<select name="CONF_PublicFlags" <%= sPublicListDsp %>>
<%		If oRS.Collect("PublicFlag") = "1" Then		%>
			<option value="1" selected>�f��</option>
			<option value="0">��f��</option>
<%		Else	%>
            <option value="1">�f��</option>
            <option value="0" selected>��f��</option>
<%		End If	%>
</select>
</div>
<div>
<span>�f�ړ�</span>	<%= sPublicDay %>
</div>
<div>
<span>�o�^��</span> <%= sRegistDay %>
</div>
<div>
<input type="checkbox" name="CONF_DeleteFlags" value="<%= oRS.Collect("OrderCode") %>">�폜
</div>
</div>
<br clear="both">
<%	End If	%>

 <% if Replace(sPublishLimitStr, "/", " ") >= Replace(Date, "/", " ") Then  %>   
    
	<table border="0" class="old <%= HtmlClassName %>">
 <% else %>   
    <table border"0" class="old motto_old <%= HtmlClassName %>">
 <% end if %>
    
	<tbody>
	<tr>
	<td class="old11" valign="middle" colspan="2">
    
    <%

	If vUserType = "" Or vUserType = "staff" Then
		'�񃍃O�C�����A�X�^�b�t���O�C����

		'�E���l�[�t�q�k�����[�����M
		'�E�E�H�b�`���X�g�֕ۑ�
		%>
		<div class="order_titele">
			<%= sTitleCompanyName %>
            <h3><a href="<%= HTTPS_CURRENTURL %>order/order_detail.asp?OrderCode=<%= oRS.Collect("OrderCode") %>">
			<% if sCatchCopy = "" then
			 response.write "�ڍׂ͂�����"
			  else
			  response.write sCatchCopy
			  end if %>
            </a><%= sImgMail %></h3>
		</div><!--/order_titele-->
		      
        <%
        	ElseIf vUserType = "company" Then
		'��ƃ��O�C����
		%>
		<div class="order_titele">
			<%= sTitleCompanyName %>
            <h3><a href="<%= HTTPS_CURRENTURL %>order/order_detail.asp?OrderCode=<%= oRS.Collect("OrderCode") %>">
			<% if sCatchCopy = "" then
			 response.write "�ڍׂ͂�����"
			  else
			  response.write sCatchCopy
			  end if %>
            </a><%= sImgMail %></h3>
		</div><!--/order_titele-->
<%	End If %>
	<div class="support_type oiwai_<%= HimlOiwai %>">
    	
    </div>
	</td><!--/old11-->
	</tr>
	
    <tr>
	<td class="old12" colspan="2">
    <div class="order_state"><%= sImgOrderState %></div>
    <div class="publish_limit">�f�ڊ����F<%= sPublishLimitStr %>�@<span class="jSpan">���R�[�h</span>(<%= oRS.Collect("OrderCode") %>)</div> 
    <div class="arrow_img"></div>
    </td>
    </tr>
   
   <% If flgImg = True Then '�摜���L��ꍇ�̃��C�A�E�g %>
    <tr>
    <td class="old27 td_point">
    <div>
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_typ.png">
    <p><%= sTitleJobName %></p>
    </div><!--�E��-->
    
    <div>
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_emp.png">
    <p><%= sWorkingType %></p>
     </div>
     
     <div>
     <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_sal.png">
      <p id="salary">
    <% If sYearlyIncome <> "" Then %>
	<span>�N��</span>
	<span><%= sYearlyIncome %></span><br>
    <% End If %>
    
    <% If sMonthlyIncome <> "" Then %>
	<span>����</span>
	<span><%= sMonthlyIncome %></span><br>
    <% End If %>
    
    <% If sHourlyIncome <> "" Then %>
	<span>����</span>
	<span><%= sHourlyIncome %></span>
    <% End If %>
     </p>
     </div>
    
    <div id="kinmuchi">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_loc.png">
    <p><%= sWorkingPlace %><%= sStationName%></p>
    </div>
    

    
    </td><!--/old27-->
    	<% if Replace(sPublishLimitStr, "/", " ") >= Replace(Date, "/", " ") Then  %> 
    		<td class="td_img">
    	<% else %>
    		<td class="td_img" style="background:#f5f5fa;">
    	<% end if %>
    
    <div id="pics_out">
        <div id="pics_main_out">
        	<a href="<%= HTTP_NAVI_CURRENTURL %>order/order_detail.asp?OrderCode=<%= oRS.Collect("OrderCode") %>" title="<%= sTitleCompanyName %>" style="width: 316px; height: 237px; display: block;<%= sImgMain %> background-size: contain; background-position: 50%;"></a>
        </div>
	<%= sImgSub %>
    </div>
	
    </td><!--/td_img-->
    </tr>
    <tr>
    	<td colspan="2" class="b_detail">
        	<b>�y�S���Ɩ��̐����z</b><br>
    		<%= sBusinessDetail %>
        </td>
    </tr>
    <tr>
    <td class="old28 td_point" colspan="2">
	<a href="<%= HTTPS_CURRENTURL %>order/order_detail.asp?OrderCode=<%= oRS.Collect("OrderCode") %>" class="neo_reg" target="_self">�ڍׂ��݂�</a>
    
    
	
    
    <% if Replace(sPublishLimitStr, "/", " ") >= Replace(Date, "/", " ") Then 
    

			If flgAddWatchList = True Then
				Response.Write "<div class=""order_button""><span class=""m0 kentoZumi"">���̋��l�[�͊��ɂ��C�ɓ��胊�X�g�ɓo�^�ς݂ł�</span></div>"
			Else
			
				If vUserType = "staff" Then
					%>
                
                    <div class="order_button">
                        <!--<a href="#" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;">		
                            �E�H�b�`���X�g�ɒǉ�</a>-->
                        <form id="<%= sOrderCode %>frmSendMailJobOfferAddress">
                        	<%
								If sOrderType <> "0"  Then
							%>
                            	<input type="button" value="���l�ւ̖⍇��" class="qmail" onclick="contactCompany('1',this);" name="<%= sOrderCode %>">
                            <%
								end if
							%>
                        
                            <input type="button" value="���C�ɓ��胊�X�g" onclick="kento(this)" class="kento" name="<%= sOrderCode %>">
                            <input type="hidden" name="CONF_OrderCode" value="<%= sOrderCode %>">
                        </form>
                    </div><!--/order_button-->
                <%
					
					else
					%>
                    	<div class="order_button">
							<form>
                            <%
							If sOrderType <> "0"  Then
							%>
                            	<input type="button" value="���l�ւ̖⍇��" class="qmail" onclick="forRegQ(this)" name="<%= sOrderCode %>">
                            <%
							end if
							%>
                            
                            	<input type="button" value="���C�ɓ��胊�X�g" onclick="forReg(this)" class="kento" name="<%= sOrderCode %>">
                            </form>
                        </div>
					<%
					
				end if
			
				
			End If
		
		

	end if %>
    
    </td><!--/old28-->
    </tr>
		
        	

    

<% Else '�摜�������ꍇ�̃��C�A�E�g %>    
    
    <tr>
    <td class="old21 td_point">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_typ.png">
    <p><%= sTitleJobName %></p>
    </td><!--/old21-->
        <td class="old22 td_point">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_emp.png">
    <p><%= sWorkingType %></p>
    </td><!--/old22-->
    </tr> 
 
     <tr>
    <td class="old24 td_point">
   <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_sal.png">
    <p id="salary">
    <% If sYearlyIncome <> "" Then %>
	<span>�N��</span>
	<span><%= sYearlyIncome %></span><br>
    <% End If %>
    
    <% If sMonthlyIncome <> "" Then %>
	<span>����</span>
	<span><%= sMonthlyIncome %></span><br>
    <% End If %>
    
    <% If sHourlyIncome <> "" Then %>
	<span>����</span>
	<span><%= sHourlyIncome %></span>
    <% End If %>
    </p>
    </td><!--/old24-->
    <td class="old23 td_point" id="kinmuchi2">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_loc.png">
    <p><%= sWorkingPlace %><%= sStationName%></p>
    </td><!--/old23-->
    </tr> 
    
      
    <tr>
    <td class="old25" colspan="2">
    <b>�y�S���Ɩ��̐����z</b><br>
    <%= sBusinessDetail %>
    </td><!--/old23-->
    </tr>
    
    <tr>
    <td class="old26" colspan="2">
	<a href="<%= HTTPS_CURRENTURL %>order/order_detail.asp?OrderCode=<%= oRS.Collect("OrderCode") %>" class="neo_reg" target="_self">�ڍׂ��݂�</a>

	
	
   <% if Replace(sPublishLimitStr, "/", " ") >= Replace(Date, "/", " ") Then 
    

			If flgAddWatchList = True Then
				Response.Write "<div class=""order_button""><span class=""m0 kentoZumi"">���̋��l�[�͊��ɂ��C�ɓ��胊�X�g�ɓo�^�ς݂ł�</span></div>"
			Else
				If vUserType = "staff" Then
					%>
                
                    <div class="order_button">
                        <!--<a href="#" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;">		
                            �E�H�b�`���X�g�ɒǉ�</a>-->

                        <form id="<%= sOrderCode %>frmSendMailJobOfferAddress" method="post" action="../staff/watchlist_register.asp" onSubmit="return Submit();">
                        
                        	<%
								If sOrderType <> "0"  Then
							%>
                            	<input type="button" value="���l�ւ̖⍇��" class="qmail" onclick="contactCompany('1',this);" name="<%= sOrderCode %>">
                            <%
								end if
							%>
                        
                            <input type="button" value="���C�ɓ��胊�X�g" onclick="kento(this)" class="kento" name="<%= sOrderCode %>">
                            <input type="hidden" name="CONF_OrderCode" value="<%= sOrderCode %>">
                        </form>
                    </div><!--/order_button-->
                <%
					
					else
					%>
						<div class="order_button">
							<form>
                            
                            <%
							If sOrderType <> "0"  Then
							%>
                            	<input type="button" value="���l�ւ̖⍇��" class="qmail" onclick="forRegQ(this)" name="<%= sOrderCode %>">
                            <%
							end if
							%>
                            
                            	<input type="button" value="���C�ɓ��胊�X�g" onclick="forReg(this)" class="kento" name="<%= sOrderCode %>">
                            </form>
                        </div>
					<%
					
				end if
			End If
			

	end if %>

    </td><!--/old26-->
    </tr>
      
    
<% End If %>  

    </tbody>
    </table>




  
<%




	DspOrderListDetail = True
End Function

'******************************************************************************
'����m�F�y�[�W
'******************************************************************************
Function DspOrderListDetail4(ByRef rDB, ByVal vUserType, ByVal vUserID, ByVal vMyOrder)
	Const PICSIZEW = 240
	Const PICSIZEH = 180
	Const PICSIZESUBW = 72
	Const PICSIZESUBH = 56

	Dim sSQL
	Dim oRS
	Dim oRS2
    Dim oRS3                '����ω摜�\���Ɏg�p
	Dim flgQE
	Dim sError

	Dim dbOrderCode			'���R�[�h
	Dim dbCompanyCode		'��ƃR�[�h
	Dim sOrderType			'�󒍎��
	Dim sPlanType			'���C�Z���X�v�������
	Dim iImageLimit			'�ʐ^�f�ڐ�������
	Dim sTitleJobName		'�E��
	Dim sTitleCompanyName	'��Ж�
	Dim sImgMail			'���M�ς݃��[���摜
	Dim sImgOrderState		'��ԉ摜 HOT,�V��,���o��OK,UI�^�[��,��w,�x��120��,�t���b�N�X
	Dim sCatchCopy			'�L���b�`�R�s�[
	Dim flgImg				'�摜�̗L���t���O(�摜�̗L���Ń��C�A�E�g���ω�) [True]�L [False]��
	Dim sImgMain			'�傫���摜
	Dim sImgSub				'�������摜
	Dim sImg1,sImg2,sImg3,sImg4	'�摜URL
	Dim sBusinessDetail		'�S���Ɩ�
	Dim sWorkingType		'�Ζ��`��
	Dim sWorkingPlace		'�Ζ��n �s���{��+�s��S
	Dim sProgress			'���l�[�R����
	Dim sPublicDay			'�f�ړ�
	Dim sPublicListDsp		'�f�ڔ�f�� ���X�g�{�b�N�X�\���X�^�C�� [style="display:none;"]
	Dim sPublicFlag1		'�f�� selected
	Dim sPublicFlag0		'��f�� selected
	Dim sCompanyPictureFlag	'��Ǝʐ^�t���O ["1"]�L ["0"]��
	Dim sRegistDay			'�o�^��
	Dim sPublishLimitStr	'���l�[�f�ڏI����
	Dim sStationName		'�w��
	Dim sYearlyIncomeMin	'�N������
	Dim sYearlyIncomeMax	'�N�����
	Dim sMonthlyIncomeMin	'��������
	Dim sMonthlyIncomeMax	'�������
	Dim sDailyIncomeMin		'��������
	Dim sDailyIncomeMax		'�������
	Dim sHourlyIncomeMin	'��������
	Dim sHourlyIncomeMax	'�������
	Dim dbTopInterviewFlag	'�g�b�v�C���^�r���[�t���O
	Dim dbWValueURL			'�v�o�����[�̂t�q�k

	Dim sYearlyIncome		'�N���\���p
	Dim sDailyIncome		'�����\���p
	Dim sMonthlyIncome		'�����\���p
	Dim sHourlyIncome		'�����\���p
	'��]�Ζ��`�ԁE��]�Ζ��n�A�C�R���@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��
	Dim sWorkingCode
	Dim sWorkingName
	Dim dbWorkingPlacePrefectureCode
	Dim dbWorkingPlacePrefectureName
	Dim dbWorkingPlaceCity
	Dim sBiz
	Dim sBizName1
	Dim sBizName2
	Dim sBizName3
	Dim sBizName4
	Dim sBizPercentage1
	Dim sBizPercentage2
	Dim sBizPercentage3
	Dim sBizPercentage4
	Dim flgBusiness
	Dim idx
	'HTTP�N���X�ύX�p�ϐ�
	Dim HtmlClassName
	Dim HtmlWorkingType
    Dim HimlOiwai

    
	'If GetRSState(rRS) = False Then Exit Function
	dbOrderCode = vMyOrder
    'response.Write vMyOrder
	DspOrderListDetail4 = False

	If G_USEFLAG = "0" And vMyOrder = "1" And G_OLDAPPLICATIONCODE <> "" Then
		sSQL = "EXEC up_DtlOrder_NEO '" & vMyOrder & "', '" & G_OLDAPPLICATIONCODE & "';"
	Else
		sSQL = "EXEC up_DtlOrder_NEO '" & vMyOrder & "', '';"
	End If

	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
'�f�o�b�N�R�[�h
'Response.write "sSQL:" & sSQL & "<br>"

	If GetRSState(oRS) = False Then Exit Function	'���l�[�������폜����Ă���ꍇ�A����f�[�^������̂ɁA���l�f�[�^��\���ł��Ȃ����ݓI�o�O����̂��߃R�[�h�ǉ� 2014/07/25 �r�c

	dbCompanyCode = ChkStr(oRS.Collect("CompanyCode"))
	sOrderType    = ChkStr(oRS.Collect("OrderType"))
	sPlanType     = ChkStr(oRS.Collect("PlanTypeName"))
	iImageLimit   = oRS.Collect("ImageLimit")
	'���j�����ݒ�
	HimlOiwai = oRS.Collect("CongratulationPrice")

	'**************************************************************************
	'�E��^��Ж� start
	'--------------------------------------------------------------------------
	sTitleCompanyName = ""
	'STEP1�F��̓I�E�햼�擾
	If oRS.Collect("JobTypeDetail") <> "" Then
		If Len(oRS.Collect("JobTypeDetail")) >= 50 Then
			sTitleJobName = Left(oRS.Collect("JobTypeDetail"), 50)
		Else
			sTitleJobName = oRS.Collect("JobTypeDetail")
		End If
	End If

	'STEP2�F��̓I�E�햼������΁^��ǉ�
	'If sTitleCompanyName <> "" Then sTitleCompanyName = sTitleCompanyName & "�^"
	'STEP3�F��Ɩ��擾
	If oRS.Collect("CompanySpeciality") <>"" THEN 
			sTitleCompanyName = sTitleCompanyName & oRS.Collect("CompanySpeciality")
	Else
		If oRS.Collect("Companykbn") ="4" Then
			sTitleCompanyName = sTitleCompanyName & oRS.Collect("CompanyName")
		ElseIf oRS.Collect("OrderType") > "0" then
				sTitleCompanyName = sTitleCompanyName & "���X�������"
		Else
				sTitleCompanyName = sTitleCompanyName & oRS.Collect("CompanyName")
		End If
	End If
	'--------------------------------------------------------------------------
	'�E��^��Ж� end
	'**************************************************************************

	'******************************************************************************
	'���^ start�@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��
	'------------------------------------------------------------------------------
	'�N��
	sYearlyIncomeMin = ChkStr(oRS.Collect("YearlyIncomeMin"))
	sYearlyIncomeMax = ChkStr(oRS.Collect("YearlyIncomeMax"))
	If sYearlyIncomeMin = "0" Then sYearlyIncomeMin = ""
	If sYearlyIncomeMax = "0" Then sYearlyIncomeMax = ""
	If sYearlyIncomeMin <> "" Then sYearlyIncomeMin = GetJapaneseYen(sYearlyIncomeMin)
	If sYearlyIncomeMax <> "" Then sYearlyIncomeMax = GetJapaneseYen(sYearlyIncomeMax)
	If sYearlyIncomeMin & sYearlyIncomeMax <> "" Then
		If sYearlyIncomeMin <> "" Then sYearlyIncome = sYearlyIncome & sYearlyIncomeMin
		sYearlyIncome = sYearlyIncome & "&nbsp;�`&nbsp;"
		If sYearlyIncomeMax <> "" Then sYearlyIncome = sYearlyIncome & sYearlyIncomeMax
	End If
	'����
	sMonthlyIncomeMin = ChkStr(oRS.Collect("MonthlyIncomeMin"))
	sMonthlyIncomeMax = ChkStr(oRS.Collect("MonthlyIncomeMax"))
	If sMonthlyIncomeMin = "0" Then sMonthlyIncomeMin = ""
	If sMonthlyIncomeMax = "0" Then sMonthlyIncomeMax = ""
	If sMonthlyIncomeMin <> "" Then sMonthlyIncomeMin = GetJapaneseYen(sMonthlyIncomeMin)
	If sMonthlyIncomeMax <> "" Then sMonthlyIncomeMax = GetJapaneseYen(sMonthlyIncomeMax)
	If sMonthlyIncomeMin & sMonthlyIncomeMax <> "" Then
		If sMonthlyIncomeMin <> "" Then sMonthlyIncome = sMonthlyIncome & sMonthlyIncomeMin
		sMonthlyIncome = sMonthlyIncome & "&nbsp;�`&nbsp;"
		If sMonthlyIncomeMax <> "" Then sMonthlyIncome = sMonthlyIncome & sMonthlyIncomeMax
	End If
	'����
	sDailyIncomeMin = ChkStr(oRS.Collect("DailyIncomeMin"))
	sDailyIncomeMax = ChkStr(oRS.Collect("DailyIncomeMax"))
	If sDailyIncomeMin = "0" Then sDailyIncomeMin = ""
	If sDailyIncomeMax = "0" Then sDailyIncomeMax = ""
	If sDailyIncomeMin <> "" Then sDailyIncomeMin = GetJapaneseYen(sDailyIncomeMin)
	If sDailyIncomeMax <> "" Then sDailyIncomeMax = GetJapaneseYen(sDailyIncomeMax)
	If sDailyIncomeMin & sDailyIncomeMax <> "" Then
		If sDailyIncomeMin <> "" Then sDailyIncome = sDailyIncome & sDailyIncomeMin
		sDailyIncome = sDailyIncome & "&nbsp;�`&nbsp;"
		If sDailyIncomeMax <> "" Then sDailyIncome = sDailyIncome & sDailyIncomeMax
	End If
	'����
	sHourlyIncomeMin = ChkStr(oRS.Collect("HourlyIncomeMin"))
	sHourlyIncomeMax = ChkStr(oRS.Collect("HourlyIncomeMax"))
	If sHourlyIncomeMin = "0" Then sHourlyIncomeMin = ""
	If sHourlyIncomeMax = "0" Then sHourlyIncomeMax = ""
	If sHourlyIncomeMin <> "" Then sHourlyIncomeMin = GetJapaneseYen(sHourlyIncomeMin)
	If sHourlyIncomeMax <> "" Then sHourlyIncomeMax = GetJapaneseYen(sHourlyIncomeMax)
	If sHourlyIncomeMin & sHourlyIncomeMax <> "" Then
		If sHourlyIncomeMin <> "" Then sHourlyIncome = sHourlyIncome & sHourlyIncomeMin
		sHourlyIncome = sHourlyIncome & "&nbsp;�`&nbsp;"
		If sHourlyIncomeMax <> "" Then sHourlyIncome = sHourlyIncome & sHourlyIncomeMax
	End If

	'------------------------------------------------------------------------------
	'���^ end
	'******************************************************************************

	'******************************************************************************
	'�Ŋ�w start�@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��
	'2008/10/22 LIS K.Kokubo �Ζ��n�������ɂ��\���ʂ������鋰�ꂪ���邽�߂ɔ�\���ɁB
	'------------------------------------------------------------------------------
	'sStationName = ""
	'sSQL = "sp_GetDataNearbyStation '" & dbOrderCode & "'"
	'flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	'If GetRSState(oRS2) = True Then
	'	sStationName ="�y" & sStationName & GetStrNearbyStation(oRS2.Collect("StationName"), "", "") & "�z"
	'End If
	'------------------------------------------------------------------------------
	'�Ŋ�w end
	'******************************************************************************

	'**************************************************************************
	'���[�����M�ς݊m�F start
	'--------------------------------------------------------------------------
	If vUserType = "staff" Then
		sSQL = "up_DtlMailHistory_Order '" & vUserID & "', '" & dbOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			sImgMail = "<img src=""/img/s_contact.gif"" alt=""���[�����M�ς�"">"
		End If
		Call RSClose(oRS2)
	End If

	'�u���l�[�����[�����M�v�̃����N�ɂԂ���Ȃ��悤�ɐE�햼�����(2007/08/01 T.Sotome�ǉ�)
	If LenByte(sTitleCompanyName) > 72 Then
		sTitleCompanyName = LeftByte(sTitleCompanyName, 70) & "..."
	End If
	'�u�E�H�b�`���X�g�֕ۑ��v�̃����N�ɂԂ���Ȃ��悤�ɐE�햼�����(2007/06/26 T.Sotome�ǉ�)
	If sImgMail = "" Then
		If LenByte(sTitleJobName) > 46 Then
			sTitleJobName = LeftByte(sTitleJobName, 44) & "..."
		End If
	Else
		If LenByte(sTitleJobName) > 36 Then
			sTitleJobName = LeftByte(sTitleJobName, 34) & "..."
		End If
	End If

	'--------------------------------------------------------------------------
	'���[�����M�ς݊m�F end
	'**************************************************************************

	'**************************************************************************
	'���img start
	'--------------------------------------------------------------------------
	sImgOrderState = GetImgOrderSpeciality(rDB, oRS)
	'--------------------------------------------------------------------------
	'���img end
	'**************************************************************************

	'**************************************************************************
	'�L���b�`�R�s�[ start
	'--------------------------------------------------------------------------
	sCatchCopy = ""
	sCatchCopy = oRS.Collect("CatchCopy")
	'--------------------------------------------------------------------------
	'�L���b�`�R�s�[ end
	'**************************************************************************

	'**************************************************************************
	'�摜 start
	'--------------------------------------------------------------------------
	flgImg = False
	If sOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS2,sSQL,sError)
		If GetRSState(oRS2) = True Then
			If ChkStr(oRS2.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS2.Collect("PicNo1")
			End If
			If ChkStr(oRS2.Collect("PicNo2")) <> "" Then
				sImg2 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS2.Collect("PicNo2")
			End If
			If ChkStr(oRS2.Collect("PicNo3")) <> "" Then
				sImg3 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS2.Collect("PicNo3")
			End If
			If ChkStr(oRS2.Collect("PicNo4")) <> "" Then
				sImg4 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS2.Collect("PicNo4")
			End If
		End If
		Call RSClose(oRS2)
	ElseIf iImageLimit > 0 Then
		sCompanyPictureFlag = ChkStr(oRS.Collect("CompanyPictureFlag"))

		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & oRS.Collect("OrderCode") & "', 'orderpicture'"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			If ChkStr(oRS2.Collect("OptionNo1")) <> "" Or (sOrderType = "0" And sCompanyPictureFlag = "1") Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo1")
			End If

			If sPlanType = "platinum" Or sPlanType = "old" Or iImageLimit > 1 Then
				If ChkStr(oRS2.Collect("OptionNo2")) <> "" Then
					sImg2 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo2")
				End If
				If ChkStr(oRS2.Collect("OptionNo3")) <> "" Then
					sImg3 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo3")
				End If
				If ChkStr(oRS2.Collect("OptionNo4")) <> "" Then
					sImg4 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo4")
				End If
			End If
		Else
			If sCompanyPictureFlag = "1" And sOrderType = "0" Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If

		Call RSClose(oRS2)
	End If

	If sImg1 & sImg2 & sImg3 & sImg4 <> "" Then flgImg = True

	If sImg1 <> "" Then
		sImgMain = "<img src=""" & sImg1 & """>"
	End If

	If sImg2 <> "" Then
		sImgSub = sImgSub & "<div class=""sub_img"">" & _
			"<img src=""" & sImg2 & """><br>"
		sImgSub = sImgSub & "</div>"
		flgImg = True
	End If
	If sImg3 <> "" Then
		sImgSub = sImgSub & "<div class=""sub_img"" style=""margin-top: 5px;"">" & _
			"<img src=""" & sImg3 & """><br>"
		sImgSub = sImgSub & "</div>"
		flgImg = True
	End If
	'If sImg4 <> "" Then
	'	sImgSub = sImgSub & "<div class=""sub_img"">" & _
	'		"<img src=""" & sImg4 & """><br>"
	'	sImgSub = sImgSub & "</div>"
	'End If

	If sImgSub <> "" Then sImgSub =  sImgSub 
	'--------------------------------------------------------------------------
	'�摜 end
	'**************************************************************************

	'**************************************************************************
	'�S���Ɩ� start
	'--------------------------------------------------------------------------
	If flgImg = True Then
		'�摜���L��ꍇ�͕��͂�Z�߂ɃJ�b�g
		sBusinessDetail = Left(oRS.Collect("BusinessDetail"),300) & "&nbsp;"
		If Len(sBusinessDetail) > 300 Then sBusinessDetail = sBusinessDetail & "..."
	Else
		'�摜�������ꍇ�͕��͂𒷂߂ɃJ�b�g
		sBusinessDetail = Left(oRS.Collect("BusinessDetail"),465) & "&nbsp;"
		If Len(sBusinessDetail) > 465 Then sBusinessDetail = sBusinessDetail & "..."
	End If
	'--------------------------------------------------------------------------
	'�S���Ɩ� end
	'**************************************************************************

	'**************************************************************************
	'�Ζ��`�� start
	'--------------------------------------------------------------------------
	sWorkingType = ""
	sSQL = "sp_GetDataWorkingType '" & oRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Dim Counter
	Counter = 1
	Do While GetRSState(oRS2) = True
		if Counter = 1 Then
			HtmlWorkingType = oRS2.Collect("WorkingTypeCode")
			Counter = Counter + 1
		End If
		sWorkingType = sWorkingType & oRS2.Collect("WorkingTypeName")
		If (oRS.Collect("OrderType") ="0" And oRS.Collect("Companykbn") = "2") Or oRS.Collect("OrderType") ="1" Or oRS.Collect("OrderType") ="2" Or oRS.Collect("OrderType") ="3" Then
			Select Case oRS2.Collect("WorkingTypeCode")
				Case "001": sWorkingType = sWorkingType & "<span class=""smartNone"">�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/koyoukeitai_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")' class=""haken_tr"">�h���Ƃ�</a>�z</span>" 
				Case "002","003": sWorkingType = sWorkingType & "<span class=""smartNone"">�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")' class=""shokai_tr"">�l�ޏЉ�Ƃ�</a>�z</span>" 
				Case "004": sWorkingType = sWorkingType & "<span class=""smartNone"">�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/syoukaiyotei_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>�Љ�\��h���Ƃ�</a>�z</span>" 
			End Select
		End If
		sWorkingType = sWorkingType & "<br>"
		oRS2.MoveNext
	Loop
	Call RSClose(oRS2)
	'--------------------------------------------------------------------------
	'�Ζ��`�� end
	'**************************************************************************

	'**************************************************************************
	'�Ζ��n start
	'--------------------------------------------------------------------------
	idx = 0
	sWorkingPlace = ""
	sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True And idx < 20
		dbWorkingPlacePrefectureCode = ChkStr(oRS2.Collect("WorkingPlacePrefectureCode"))
		dbWorkingPlacePrefectureName = ChkStr(oRS2.Collect("WorkingPlacePrefectureName"))
		dbWorkingPlaceCity = ChkStr(oRS2.Collect("WorkingPlaceCity"))
		'<�Ζ��n�A�C�R��>
		If InStr(sImgOrderState, "/icon_p" & dbWorkingPlacePrefectureCode & ".gif") = 0 Then
			'�����s���{���A�C�R���͏o���Ȃ��I
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_p" & dbWorkingPlacePrefectureCode & ".gif"" alt=""" & dbWorkingPlacePrefectureName & """ width=""50"" height=""15"">&nbsp;"
		End If
		'</�Ζ��n�A�C�R��>

		'<�Ζ��n>
		If sWorkingPlace <> "" Then sWorkingPlace = sWorkingPlace & "/"
		sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
		'</�Ζ��n>

		oRS2.MoveNext
		idx = idx + 1
	Loop
	If oRS2.RecordCount > 20 Then sWorkingPlace = sWorkingPlace & "&nbsp;...��"
	Call RSClose(oRS2)
	'--------------------------------------------------------------------------
	'�Ζ��n end
	'**************************************************************************

	'**************************************************************************
	'�f�ڏ�ԃ��X�g�{�b�N�X start
	'--------------------------------------------------------------------------
	sPublicFlag1 = ""
	sPublicFlag0 = ""
	If oRS.Collect("PublicFlag") = "1" Then
		sPublicFlag1 = " selected"
	Else
		sPublicFlag0 = " selected"
	End If
	'--------------------------------------------------------------------------
	'�f�ڏ�ԃ��X�g�{�b�N�X start
	'**************************************************************************

	'**************************************************************************
	'�R���̐i�� start
	'--------------------------------------------------------------------------
	sProgress = ""
	sPublicListDsp = ""
	sPublicDay = ""

	'�R����
	If oRS.Collect("PermitFlag") = "0" Then
		'���X���R��
		sProgress = "���X�R����"
		sPublicListDsp = "style=""display:none;"""
	ElseIf oRS.Collect("PermitFlag") = "1" Then
		'���X����
		If oRS.Collect("PublicFlag") = "0" Then
			sProgress = "���X����(��f��)"
		Else
			sProgress = "�f�ڒ�"
		End If
	Else
		'�ȊO
		If oRS.Collect("PublicFlag") = "1" And oRS.Collect("PermitFlag") = "1" Then
			sProgress = "�f��"
		Else
			sProgress = "��f��"
		End If
		sPublicListDsp = "style=""display:none;"""
	End If

	'�f�ړ�
	sPublicDay = GetDateStr(oRS.Collect("PublicDay"), "/")
	If oRS.Collect("PermitFlag") = "1" And oRS.Collect("PublicDay") > Date Then
		sPublicDay = "<span style=""color:#ff0000;"">��(" & sPublicDay & ")</span>"
		sPublicListDsp = "style=""display:none;"""
	End If
	'--------------------------------------------------------------------------
	'�R���̐i�� end
	'**************************************************************************

	'**************************************************************************
	'�o�^�� start
	'--------------------------------------------------------------------------
	sRegistDay = GetDateStr(oRS.Collect("RegistDay"), "/")
	'--------------------------------------------------------------------------
	'�o�^�� end
	'**************************************************************************

	'******************************************************************************
	'���l�[�f�ڊ��� start
	'------------------------------------------------------------------------------
	'��ƃ��O�C�����ȊO�̂Ƃ��Ɍf�ڊ�����\��
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(oRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(oRS.Collect("PublicLimitDay"))
	End If

	If sPublishLimitStr = "" Then
		If oRS.Collect("NowPublicFlag") = "0" Then
			'���C�Z���X�؂�̂Ƃ���"�f�ڏI��"�ƕ\��
			sPublishLimitStr = "�f�ڏI��"
		Else
			sPublishLimitStr = "�펞��W��"
		End If
	End If

	sPublishLimitStr = sPublishLimitStr & "&nbsp;"
	'------------------------------------------------------------------------------
	'���l�[�f�ڊ��� end
	'******************************************************************************

	'******************************************************************************
	'�d���̊��� start�@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��
	'------------------------------------------------------------------------------
	If sPlanType = "platinum" Or sPlanType = "gold" Or sPlanType = "old" Then
		sBiz = ""
		sBizName1 = ""
		sBizName2 = ""
		sBizName3 = ""
		sBizName4 = ""
		sBizPercentage1 = ""
		sBizPercentage2 = ""
		sBizPercentage3 = ""
		sBizPercentage4 = ""

		sBizName1 = ChkStr(oRS.Collect("BizName1"))
		sBizName2 = ChkStr(oRS.Collect("BizName2"))
		sBizName3 = ChkStr(oRS.Collect("BizName3"))
		sBizName4 = ChkStr(oRS.Collect("BizName4"))
		sBizPercentage1 = ChkStr(oRS.Collect("BizPercentage1"))
		sBizPercentage2 = ChkStr(oRS.Collect("BizPercentage2"))
		sBizPercentage3 = ChkStr(oRS.Collect("BizPercentage3"))
		sBizPercentage4 = ChkStr(oRS.Collect("BizPercentage4"))
		If sBizPercentage1 = "" Then sBizPercentage1 = "0"
		If sBizPercentage2 = "" Then sBizPercentage2 = "0"
		If sBizPercentage3 = "" Then sBizPercentage3 = "0"
		If sBizPercentage4 = "" Then sBizPercentage4 = "0"

		If Len(sBizName1) >= 17 Then sBizName1 = Left(sBizName1,17) & "..."
		If Len(sBizName2) >= 17 Then sBizName2 = Left(sBizName2,17) & "..."
		If Len(sBizName3) >= 17 Then sBizName3 = Left(sBizName3,17) & "..."
		If Len(sBizName4) >= 17 Then sBizName4 = Left(sBizName4,17) & "..."

		If sBizName1 & sBizName2 & sBizName3 & sBizName4 <> "" Then
			If sBizName1 <> "" And sBizPercentage1 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName1 & "</td><td class=""biz2"">" & sBizPercentage1 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage1) * 3 & """ height=""20""></td></tr>"
			If sBizName2 <> "" And sBizPercentage2 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName2 & "</td><td class=""biz2"">" & sBizPercentage2 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage2) * 3 & """ height=""20""></td></tr>"
			If sBizName3 <> "" And sBizPercentage3 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName3 & "</td><td class=""biz2"">" & sBizPercentage3 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage3) * 3 & """ height=""20""></td></tr>"
			If sBizName4 <> "" And sBizPercentage4 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName4 & "</td><td class=""biz2"">" & sBizPercentage4 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage4) * 3 & """ height=""20""></td></tr>"
			sBiz = "<table>" & sBiz & "</table>"
			flgBusiness = True
		End If
	End If
	'------------------------------------------------------------------------------
	'�d���̊��� end
	'******************************************************************************

	'******************************************************************************
	'�g�b�v�C���^�r���[ start
	'------------------------------------------------------------------------------
	dbTopInterviewFlag = oRS.Collect("TopInterviewFlag")
	'------------------------------------------------------------------------------
	'�g�b�v�C���^�r���[ end
	'******************************************************************************

	'******************************************************************************
	'�v�o�����[�t�q�k start
	'------------------------------------------------------------------------------
	dbWValueURL = ChkStr(oRS.Collect("WValueURL"))
	'------------------------------------------------------------------------------
	'�v�o�����[�t�q�k end
	'******************************************************************************

	Response.Write "<input type=""hidden"" name=""CONF_OrderCodes"" value=""" & oRS.Collect("OrderCode") & """>"
	
    '�N���X�ύX
    If sOrderType = "2" Then
    '�Љ�̎� 
        HtmlClassName = "neo_shokai"
    Elseif sOrderType = "1" Then
    '�h���̎�
        HtmlClassName = "neo_haken"
    Elseif sOrderType = "3" Then
    '�Љ�\��h���̎�
        HtmlClassName = "neo_ttp"
    Elseif sOrderType = "0" Then
    '�L���̂Ƃ�
        if HtmlWorkingType = "005" Then
            HtmlClassName = "neo_beit"
		Elseif HtmlWorkingType = "006" Then
			HtmlClassName = "neo_soho"
        Else
            HtmlClassName = "neo_shain"
        End if
    End if


    '******************************************************************************
    '2014/05/22 ikeda �ǉ�
    '�u����ρv�Ɖ摜�\������
    '�Ώۏ����F
    '   ���ɉ���ς݂̋��l
    '   ������Ƃ̑����l�ɉ��債�Ă���ꍇ
    '------------------------------------------------------------------------------
    Dim iExists_C_Flag  '����Ƃւ̉���f�[�^�����݂���ꍇ
    Dim iExists_O_Flag  '�����l�ւ́@�@�@�@�h
                        '0: �Ȃ� 1:���݂���
	sSQL = "EXEC up_ChkAdoptionExists '" & dbCompanyCode & "','" & vMyOrder & "','" & vUserID & "';"
	flgQE = QUERYEXE(rDB, oRS3, sSQL, sError)

	iExists_C_Flag = oRS3.Collect("Exists_C_Flag")
	iExists_O_Flag = oRS3.Collect("Exists_O_Flag")

    Call RSClose(oRS3)

    '------------------------------------------------------------------------------
    '2014/05/22 ikeda �����܂�
    '******************************************************************************


	'�L���ꗗ

	If oRS.Collect("CompanyCode") = vUserID And vMyOrder = "1" And G_USEFLAG = "1" Then

		%>
<div class="my_order">
<div>
<span>���R�[�h</span>(<%= oRS.Collect("OrderCode") %>)
</div>
<div>
<span>��� </span><%= sProgress %>
</div>
<div>
<select name="CONF_PublicFlags" <%= sPublicListDsp %>>
<%		If oRS.Collect("PublicFlag") = "1" Then		%>
			<option value="1" selected>�f��</option>
			<option value="0">��f��</option>
<%		Else	%>
            <option value="1">�f��</option>
            <option value="0" selected>��f��</option>
<%		End If	%>
</select>
</div>
<div>
<span>�f�ړ�</span>	<%= sPublicDay %>
</div>
<div>
<span>�o�^��</span> <%= sRegistDay %>
</div>
<div>
<input type="checkbox" name="CONF_DeleteFlags" value="<%= oRS.Collect("OrderCode") %>">�폜
</div>
</div>
<br clear="both">
<%	End If	%>

 <% if Replace(sPublishLimitStr, "/", " ") >= Replace(Date, "/", " ") Then  %>   
    
	<table border="0" class="old delSmart <%= HtmlClassName %>">
 <% else %>   
    <table border"0" class="old motto_old delSmart <%= HtmlClassName %>">
 <% end if %>
    
	<tbody>
	<tr>
	<td class="old11" valign="middle" colspan="2">

		<div class="order_titele">
			<%= sTitleCompanyName %>
            <h3><%= sCatchCopy %><%= sImgMail %></h3>
		</div><!--/order_titele-->

	<div class="support_type oiwai_<%= HimlOiwai %>">
    	
    </div>
	</td><!--/old11-->
	</tr>
	
    <tr>
	<td class="old12" colspan="2">
    <div class="order_state"><%= sImgOrderState %></div>
    <div class="publish_limit">�f�ڊ����F<%= sPublishLimitStr %></div> 
    <div class="arrow_img"></div>

    <% '����ς݉摜�\��(����Ƃ̑����l)
       If iExists_O_Flag = 1 Then %>
        <div class="adoption_img"></div>
    <% End If %>

    </td>
    </tr>
   

    <tr>
    <td class="old21 td_point">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_typ.png">
    <p><%= sTitleJobName %></p>
    </td><!--/old21-->
        <td class="old22 td_point">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_emp.png">
    <p><%= sWorkingType %></p>
    </td><!--/old22-->
    </tr> 
 
     <tr>
    <td class="old24 td_point">
   <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_sal.png">
    <p id="salary">
    <% If sYearlyIncome <> "" Then %>
	<span>�N��</span>
	<span><%= sYearlyIncome %></span><br>
    <% End If %>
    
    <% If sMonthlyIncome <> "" Then %>
	<span>����</span>
	<span><%= sMonthlyIncome %></span><br>
    <% End If %>
    
    <% If sHourlyIncome <> "" Then %>
	<span>����</span>
	<span><%= sHourlyIncome %></span>
    <% End If %>
    </p>
    </td><!--/old24-->
    <td class="old23 td_point" id="kinmuchi2">
    <img src="<%= HTTPS_NAVI_CURRENTURL %>img/order/list_loc.png">
    <p><%= sWorkingPlace %><%= sStationName%></p>
    </td><!--/old23-->
    </tr> 
    
      
    <tr>
    <td class="old25" colspan="2">
    <b>�y�S���Ɩ��̐����z</b><br>
    <%= sBusinessDetail %>
    </td><!--/old23-->
    </tr>


    </tbody>
    </table>




  
<%




	DspOrderListDetail4 = True
End Function



'******************************************************************************
'�T�@�v�F�������w�肵�Č���������
'���@���FrDB		�FDB�ڑ��I�u�W�F�N�g
'�@�@�@�FrRS		�F���l�[�ꗗ�̃��R�[�h�Z�b�g
'�@�@�@�FvOrderCode	�F���݂̗�
'�߂�l�F
'���@�l�F
'���@���FLIS K.NIINA
'�@�@�@�F2008/10/20 LIS K.Kokubo �Ζ��n�������ɂ��C��
'******************************************************************************
Function Retrieval(byref rDB, byref rRS, ByVal vOrderCode)
	Dim oRS
	Dim sSQL
	Dim sError
	Dim sWT
	Dim sAC2
	Dim sJT2

	Dim dbWorkingPlacePrefectureCode

	'<�Ζ��`��>
	sSQL = "EXEC sp_GetDataWorkingType '" & vOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = true Then
		sWT = oRS.Collect("WorkingTypeCode")
	End If
	Call RSClose(oRS)
	'</�Ζ��`��>

	'<�Ζ��n>
	sAC2 = ""
	sSQL = "EXEC up_LstC_WorkingPlace '" & vOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbWorkingPlacePrefectureCode = oRS.Collect("WorkingPlacePrefectureCode")

		If sAC2 <> "" Then sAC2 = sAC2 & ","
		sAC2 = sAC2 & dbWorkingPlacePrefectureCode
		oRS.MoveNext
	End If
	Call RSClose(oRS)
	'</�Ζ��n>

	'<�E��>
	sSQL = "sp_GetDataJobType '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = true Then
		sJT2 = oRS.Collect("JobTypeCode")
	End If
	Call RSClose(oRS)
	'</�E��>

	Retrieval = "<div align=""right""><a href=""/order/order_list.asp?wt=" & sWT & "&amp;ac2=" & sAC2 & "&amp;jt2=" & sJT2 & """><img src=""/img/order_detail_icon/serchimage.gif"" border=""0"" style=""vertical-align:bottom;"">�������w�肵�Č�����������</a></div>"
End Function

'******************************************************************************
'�T�@�v�F���l�[�ꗗ�̉����уo�[�W����
'���@���FrDB		�FDB�ڑ��I�u�W�F�N�g
'�@�@�@�FrRS		�F���l�[�ꗗ�̃��R�[�h�Z�b�g
'�@�@�@�FvCols		�F���݂̗�
'�@�@�@�FvMaxCols	�F��ő吔
'�߂�l�F
'�쐬���F2007/05/23
'�쐬�ҁFLis Kokubo
'���@�l�F
'���@���F
'******************************************************************************
Function DspOrderListDetail2(ByRef rDB, ByRef rRS, ByVal vCols, ByVal vMaxCols)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim sOrderCode			'���R�[�h
	Dim sOrderType			'�󒍋敪
	Dim sCompanyKbn			'��Ћ敪
	Dim sCompanyName		'��Ɩ�
	Dim sCompanyNameF		'��Ɩ��J�i
	Dim sCompanySpeciality	'��Ɩ��i�����j
	Dim sJobTypeDetail		'��̓I�E�햼(alt��title�ŏo�͂���)
	Dim sViewJobTypeDetail	'���E�҂Ɍ������̓I�E�햼(����������̓J�b�g�����)
	Dim sBusinessDetail		'�S���Ɩ�
	Dim sYearlyIncome		'�N��
	Dim sYearlyIncomeMin	'�N������
	Dim sYearlyIncomeMax	'�N�����
	Dim sMonthlyIncome		'����
	Dim sMonthlyIncomeMin	'��������
	Dim sMonthlyIncomeMax	'�������
	Dim sDailyIncome		'����
	Dim sDailyIncomeMin		'��������
	Dim sDailyIncomeMax		'�������
	Dim sHourlyIncome		'����
	Dim sHourlyIncomeMin	'��������
	Dim sHourlyIncomeMax	'�������
	Dim sWorkingTypeIcon	'�Ζ��`�ԃA�C�R������
	Dim sImg				'�摜URL

	Dim sURL				'���l�[�ڍׂ�URL
	Dim sAlign				'�g�� [vCols = 1]left [vCols = vMaxCols]right [����ȊO]center

	If GetRSState(rRS) = False Then Exit Function

	sURL = HTTPS_CURRENTURL & "order/order_detail.asp"

	If vCols = 1 Then
		sAlign = "left"
	ElseIf vCols = vMaxCols Then
		sAlign = "right"
	Else
		sAlign = "center"
	End If

	sSQL = "up_DtlOrder '" & rRS.Collect("OrderCode") & "', ''"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'���R�[�h
	sOrderCode = ChkStr(oRS.Collect("OrderCode"))
	'�󒍋敪
	sOrderType = ChkStr(oRS.Collect("OrderType"))
	'��Ƌ敪
	sCompanyKbn = ChkStr(oRS.Collect("CompanyKbn"))
	'��Ɩ�, ��Ɩ��J�i
	sCompanyName = ChkStr(oRS.Collect("CompanyName"))
	sCompanyNameF = ChkStr(oRS.Collect("CompanyName_F"))
	sCompanySpeciality = ChkStr(oRS.Collect("CompanySpeciality"))
	Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
	'��̓I�E�햼
	sJobTypeDetail = ChkStr(oRS.Collect("JobTypeDetail"))
	sViewJobTypeDetail = sJobTypeDetail
	If Len(sViewJobTypeDetail) > 14 Then sViewJobTypeDetail = Left(sViewJobTypeDetail, 14) & ".."
	'�S���Ɩ�
	sBusinessDetail = ChkStr(oRS.Collect("BusinessDetail"))

	'******************************************************************************
	'���^ start
	'------------------------------------------------------------------------------
	'�N��
	sYearlyIncomeMin = ChkStr(oRS.Collect("YearlyIncomeMin"))
	sYearlyIncomeMax = ChkStr(oRS.Collect("YearlyIncomeMax"))
	If sYearlyIncomeMin = "0" Then sYearlyIncomeMin = ""
	If sYearlyIncomeMax = "0" Then sYearlyIncomeMax = ""
	If sYearlyIncomeMin <> "" Then sYearlyIncomeMin = GetJapaneseYen(sYearlyIncomeMin)
	If sYearlyIncomeMax <> "" Then sYearlyIncomeMax = GetJapaneseYen(sYearlyIncomeMax)
	If sYearlyIncomeMin & sYearlyIncomeMax <> "" Then
		If sYearlyIncomeMin <> "" Then sYearlyIncome = sYearlyIncome & sYearlyIncomeMin
		sYearlyIncome = sYearlyIncome & "&nbsp;�`&nbsp;"
		If sYearlyIncomeMax <> "" Then sYearlyIncome = sYearlyIncome & sYearlyIncomeMax
	End If
	'����
	sMonthlyIncomeMin = ChkStr(oRS.Collect("MonthlyIncomeMin"))
	sMonthlyIncomeMax = ChkStr(oRS.Collect("MonthlyIncomeMax"))
	If sMonthlyIncomeMin = "0" Then sMonthlyIncomeMin = ""
	If sMonthlyIncomeMax = "0" Then sMonthlyIncomeMax = ""
	If sMonthlyIncomeMin <> "" Then sMonthlyIncomeMin = GetJapaneseYen(sMonthlyIncomeMin)
	If sMonthlyIncomeMax <> "" Then sMonthlyIncomeMax = GetJapaneseYen(sMonthlyIncomeMax)
	If sMonthlyIncomeMin & sMonthlyIncomeMax <> "" Then
		If sMonthlyIncomeMin <> "" Then sMonthlyIncome = sMonthlyIncome & sMonthlyIncomeMin
		sMonthlyIncome = sMonthlyIncome & "&nbsp;�`&nbsp;"
		If sMonthlyIncomeMax <> "" Then sMonthlyIncome = sMonthlyIncome & sMonthlyIncomeMax
	End If
	'����
	sDailyIncomeMin = ChkStr(oRS.Collect("DailyIncomeMin"))
	sDailyIncomeMax = ChkStr(oRS.Collect("DailyIncomeMax"))
	If sDailyIncomeMin = "0" Then sDailyIncomeMin = ""
	If sDailyIncomeMax = "0" Then sDailyIncomeMax = ""
	If sDailyIncomeMin <> "" Then sDailyIncomeMin = GetJapaneseYen(sDailyIncomeMin)
	If sDailyIncomeMax <> "" Then sDailyIncomeMax = GetJapaneseYen(sDailyIncomeMax)
	If sDailyIncomeMin & sDailyIncomeMax <> "" Then
		If sDailyIncomeMin <> "" Then sDailyIncome = sDailyIncome & sDailyIncomeMin
		sDailyIncome = sDailyIncome & "&nbsp;�`&nbsp;"
		If sDailyIncomeMax <> "" Then sDailyIncome = sDailyIncome & sDailyIncomeMax
	End If
	'����
	sHourlyIncomeMin = ChkStr(oRS.Collect("HourlyIncomeMin"))
	sHourlyIncomeMax = ChkStr(oRS.Collect("HourlyIncomeMax"))
	If sHourlyIncomeMin = "0" Then sHourlyIncomeMin = ""
	If sHourlyIncomeMax = "0" Then sHourlyIncomeMax = ""
	If sHourlyIncomeMin <> "" Then sHourlyIncomeMin = GetJapaneseYen(sHourlyIncomeMin)
	If sHourlyIncomeMax <> "" Then sHourlyIncomeMax = GetJapaneseYen(sHourlyIncomeMax)
	If sHourlyIncomeMin & sHourlyIncomeMax <> "" Then
		If sHourlyIncomeMin <> "" Then sHourlyIncome = sHourlyIncome & sHourlyIncomeMin
		sHourlyIncome = sHourlyIncome & "&nbsp;�`&nbsp;"
		If sHourlyIncomeMax <> "" Then sHourlyIncome = sHourlyIncome & sHourlyIncomeMax
	End If
	'------------------------------------------------------------------------------
	'���^ end
	'******************************************************************************

	'******************************************************************************
	'�Ζ��`�ԃA�C�R�� start
	'------------------------------------------------------------------------------
	sWorkingTypeIcon = ""
	sSQL = "sp_GetListWorkingType '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True
		Select Case ChkStr(oRS2.Collect("WorkingTypeCode"))
			Case "001": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/haken.gif"" alt=""�h��"">&nbsp;"
			Case "002": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/seishain.gif"" alt=""���Ј�"">&nbsp;"
			Case "003": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/keiyaku.gif"" alt=""�_��Ј�"">&nbsp;"
			Case "004": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/syoha.gif"" alt=""�Љ�\��h��"">&nbsp;"
			Case "005": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/arbeit.gif"" alt=""�A���o�C�g�E�p�[�g"">&nbsp;"
			Case "006": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/soho.gif"" alt=""SOHO"">&nbsp;"
			Case "007": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/fc.gif"" alt=""FC"">&nbsp;"
		End Select
		oRS2.MoveNext
	Loop
	Call RSClose(oRS2)
	'------------------------------------------------------------------------------
	'�Ζ��`�ԃA�C�R�� end
	'******************************************************************************

	'******************************************************************************
	'�摜 start
	'------------------------------------------------------------------------------
	sImg = ""
	sSQL = "up_GetListOrderPictureNow '" & sCompanyCode & "', '" & sOrderCode & "', 'orderpicture'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	If GetRSState(oRS2) = True Then
		If sImg = "" And ChkStr(oRS2.Collect("OptionNo1")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo1")
		If sImg = "" And ChkStr(oRS2.Collect("OptionNo2")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo2")
		If sImg = "" And ChkStr(oRS2.Collect("OptionNo3")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo3")
		If sImg = "" And ChkStr(oRS2.Collect("OptionNo4")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & oRS2.Collect("OptionNo4")
	End If

	If sImg = "" And sOrderType = "0" Then
		sSQL = "sp_GetDataPicture '" & sCompanyCode & "', '1'"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			sImg = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=1"
		End If
	End If

	If sImg = "" Then sImg = "/img/nopicture180.gif"
	'------------------------------------------------------------------------------
	'�摜 end
	'******************************************************************************
%>
<div align="<%= sAlign %>" style="float:left; width:200px;">
	<table class="pattern1" border="0" style="width:195px;">
		<thead>
		<tr>
			<th colspan="2" valign="top" style="width:183px;">
				<div style="float:left; width:64px;"><img src="<%= sImg %>" alt="<%= sJobTypeDetail %>" width="64" height="48"></div>
				<div style="float:left; width:114px; margin-left:5px;"><a href="<%= sURL %>?ordercode=<%= sOrderCode %>"><%= sViewJobTypeDetail %></a></div>
				<br clear="all">
			</th>
		</tr>
		</thead>
		<tbody>
<!--
		<tr>
			<td colspan="2" align="center">
				<a href="<%= sURL %>?ordercode=<%= sOrderCode %>" title="<%= sJobTypeDetail %>">
					<img src="<%= sImg %>" alt="<%= sJobTypeDetail %>" border="1" width="180" height="135" style="border-color:#999999;">
				</a>
			</td>
		</tr>
-->
		<tr>
			<th style="width:63px;">��Ж�</th>
			<td style="width:109px;"><%= sCompanyName %></td>
		</tr>
		<tr>
			<th>�Ζ��`��</th>
			<td><%= sWorkingTypeIcon %></td>
		</tr>
<!--
		<tr>
			<th>�S���Ɩ�</th>
			<td><%= sBusinessDetail %></td>
		</tr>
-->
<%
			If sYearlyIncome <> "" Then
%>
		<tr>
			<th>�N��</th>
			<td><%= sYearlyIncome %></td>
		</tr>
<%
			End If

			If sMonthlyIncome <> "" Then
%>
		<tr>
			<th>����</th>
			<td><%= sMonthlyIncome %></td>
		</tr>
<%
			End If

			If sDailyIncome <> "" Then
%>
		<tr>
			<th>����</th>
			<td><%= sDailyIncome %></td>
		</tr>
<%
			End If

			If sHourlyIncome <> "" Then
%>
		<tr>
			<th>����</th>
			<td><%= sHourlyIncome %></td>
		</tr>
<%
			End If
%>
		</tbody>
	</table>
</div>
<%
End Function

'******************************************************************************
'�T�@�v�F���l�[�ꗗ�����уo�[�W����2
'���@���FrDB		�FDB�ڑ��I�u�W�F�N�g
'�@�@�@�FrRS		�F���d���������ʂ�ێ�����̃��R�[�h�Z�b�g
'�@�@�@�FvPageSize	�F�P�y�[�W������̋��l�[����
'�@�@�@�FvPage		�F���ݕ\�����̃y�[�W
'�@�@�@�FvRCMD		�F���R�����h��� ["1"]����Ȃ��d���������Ă܂� ["2"]�߂������̂��d����� ["3"]���i
'�߂�l�F
'�쐬���F2007/05/31
'�쐬�ҁFLis Kokubo
'���@�l�F
'���@���F
'******************************************************************************
Function DspOrderListDetail3(ByRef rDB, ByRef rRS, ByVal vPageSize, ByVal vPage, ByVal vRCMD)
	Const MAXCOLS = 3

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sTitle
	Dim iRecordCnt	'���R�[�h����
	Dim idx			'���[�v�J�E���g�A�b�v�ϐ�
	Dim iCols		'��
	Dim aPadding(2)	'�e��̃p�f�B���O
	Dim aJobTypeDetail()
	Dim aCompanyName()
	Dim aImg()
	Dim aWorkingTypeIcon()
	Dim aWorkingPlace()
	Dim aStation()
	Dim aYearlyIncome()
	Dim aMonthlyIncome()
	Dim aDailyIncome()
	Dim aHourlyIncome()

	If GetRSState(rRS) = False Then Exit Function
	If IsNumeric(vPageSize) = False Then Exit Function

	If IsNumeric(vPage) = False Then vPage = 1
	rRS.PageSize = vPageSize
	rRS.AbsolutePage = vPage

	If GetRSState(rRS) = False Then Exit Function

	iRecordCnt = 0
	idx = 0
	Do While GetRSState(rRS) = True And idx < vPageSize
		ReDim Preserve aJobTypeDetail(idx)
		ReDim Preserve aCompanyName(idx)
		ReDim Preserve aImg(idx)
		ReDim Preserve aWorkingTypeIcon(idx)
		ReDim Preserve aWorkingPlace(idx)
		ReDim Preserve aStation(idx)
		ReDim Preserve aYearlyIncome(idx)
		ReDim Preserve aMonthlyIncome(idx)
		ReDim Preserve aDailyIncome(idx)
		ReDim Preserve aHourlyIncome(idx)

		Call GetRecommendValues(rDB, rRS, vRCMD, aJobTypeDetail(idx), aCompanyName(idx), aImg(idx), aWorkingTypeIcon(idx), aWorkingPlace(idx), aStation(idx), aYearlyIncome(idx), aMonthlyIncome(idx), aDailyIncome(idx), aHourlyIncome(idx))
		idx = idx + 1
		iRecordCnt = iRecordCnt + 1
		rRS.MoveNext
	Loop

	'�e��̃p�f�B���O
	aPadding(0) = "padding:0px 4px 0px 0px;"
	aPadding(1) = "padding:0px 2px 0px 2px;"
	aPadding(2) = "padding:0px 0px 0px 4px;"

	idx = 0
	Do While idx < iRecordCnt
		For iCols = 0 To MAXCOLS - 1
			If idx >= iRecordCnt Then Exit For

			Response.Write "<div class=""sec_div"" style=""float:left; width:235px; margin:0 2px;""><div class=""thr_div"" style=""line-height:16px; " & aPadding(iCols) & """>"

			Response.Write "<div class=""jovimg"">" & aImg(idx) & "</div>"
			Response.Write "<div class=""jovtext"">"
			If aJobTypeDetail(idx) <> "" Then Response.Write "�y�E��z" & aJobTypeDetail(idx) & "<br>" & vbCrLf
			'If aCompanyName(idx) <> "" Then Response.Write "�y��Ɓz" & aCompanyName(idx) & "<br>" & vbCrLf
			If aWorkingTypeIcon(idx) <> "" Then Response.Write "�y�`�ԁz" & aWorkingTypeIcon(idx)  & "<br>"& vbCrLf
			If aWorkingPlace(idx) <> "" Then Response.Write "�y�ꏊ�z" & aWorkingPlace(idx) & "<br>" & vbCrLf
			If aStation(idx) <> "" Then Response.Write "�y�Ŋ�z" & Replace(aStation(idx), "<br>", "�A") & "<br>" & vbCrLf
			Response.Write "�y���^�z"
			If aYearlyIncome(idx) <> "" Then Response.Write "[�N��]" & aYearlyIncome(idx)
			If aMonthlyIncome(idx) <> "" Then Response.Write "[����]" & aMonthlyIncome(idx)
			If aDailyIncome(idx) <> "" Then Response.Write "[����]" & aDailyIncome(idx)
			If aHourlyIncome(idx) <> "" Then Response.Write "[����]" & aHourlyIncome(idx)

			idx = idx + 1
			Response.Write "</div></div></div>"
		Next

		Response.Write "<div style=""padding-bottom:15px; clear:both;""></div>"
	Loop
End Function

'******************************************************************************
'�T�@�v�F���l�[�ꗗ�y�[�W�̃y�[�W�R���g���[��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_SearchOrder or ���l�[�ڍ׌���SQL �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvPageSize		�F�P�y�[�W������̕\������
'�@�@�@�FvPage			�F�\�����y�[�W
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/order_list_entity.asp
'�@�@�@�F�����ƃi�r/order/company_order.asp
'******************************************************************************
Function DspPageControl(ByRef rDB, ByRef rRS, ByVal vPageSize, ByVal vPage)
	Dim iMaxPage
	Dim iLine
	Dim S_Page
	Dim E_Page
	Dim Sort
	Dim idx

	If GetRSState(rRS) = False Then Exit Function

	If vPage <> "" Then vPage = CInt(vPage)

	'�y�[�W������̕\������
	rRS.PageSize = vPageSize

	iMaxPage = rRS.PageCount
	If vPage > iMaxPage Then vPage = iMaxPage
	rRS.AbsolutePage = vPage

	'��ʏ�ɕ\������J�n�E�I���y�[�W�ԍ���ݒ�
	'�\���J�n�y�[�W�ԍ����w��
	S_Page = vPage - 5
	If S_Page < 1 Then
		S_Page = 1
	End If

	'�\���I���y�[�W�ԍ����w��
	E_Page = vPage + 4
	If E_Page < 10 Then E_Page = 10
	If E_Page > iMaxPage Then
		E_Page = iMaxPage
	End If
	If S_Page > iMaxPage - 9 And iMaxPage - 9 > 0 Then S_Page = iMaxPage - 9
%>
<table style="margin:10px 0px;" class="smartPager">
	<tbody>
	<tr>
		<td style="width:88px; padding:5px; border-width:1px 0px 1px 1px; text-align:center;">
<%
	If vPage > 1 Then Response.Write "<a href='javascript:ChgPage(" & vPage - 1 & ");'>�O�̃y�[�W</a>"
%>
		</td>
		<td style="padding:5px; border-width:1px 0px 1px 0px; text-align:center;">
<%
	If S_Page <> 1 Then Response.Write "�c"
	For idx = S_Page To E_Page	'�y�[�W�ԍ���\��
		Response.write "�@"
		If idx = vPage Then		'�w��y�[�W�̕\��
			Response.Write "[" & idx & "]"
		Else
			Response.Write "<a href='javascript:ChgPage(" & idx & ");'>" & idx & "</a>"
		End If
	Next
	If E_Page < iMaxPage Then Response.Write "�@�c"
%>
		</td>
		<td style="width:89px; padding:5px; border-width:1px 1px 1px 0px; text-align:center;">
<%
	If vPage < iMaxPage Then Response.Write "<a href='javascript:ChgPage(" & vPage + 1 & ");'>���̃y�[�W</a>"
%>
		</td>
	</tr>
	</tbody>
</table>
<%
End Function

'******************************************************************************
'�T�@�v�F��Ə��̊�{�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvOrderCode		�F
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�����ƃi�r/order/company_order.asp
'���@�l�F
'���@���F2007/02/11 LIS K.Kokubo �쐬
'�@�@�@�F2008/06/25 LIS K.Kokubo �Ŋ�w�ǉ�
'******************************************************************************
Function DspCompanyInfo(ByRef rDB, ByRef rRS, ByVal vOrderCode, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbPlanTypeName		'���C�Z���X�v�����^�C�v
	Dim dbImageLimit		'�ő�摜�f�ڐ�
	Dim sCompanyCode		'��ƃR�[�h
	Dim sCompanyName		'��Ɩ���
	Dim sCompanyNameF		'��Ɩ��̃J�i
	Dim sOrderType			'���l��� ["0"]�����ƃi�r��� ["1"]�h�� ["2"]�Љ� ["3"]
	Dim sCompanyPictureFlag	'��Ǝʐ^�t���O ["1"]�L ["0"]��
	Dim sCompanyKbn			'��Ƌ敪
	Dim sCompanySpeciality	'��Ɠ���
	Dim sEstablishYear		'�ݗ��N�x
	Dim sCapitalAmount		'���{�z
	Dim sListClass			'�������J
	Dim sEmployeeNum		'�Ј���
	Dim sIndustryType		'�Ǝ�
	Dim sAddress			'�{�ЏZ��
	Dim sHomePage			'�z�[���y�[�W
	Dim dbStationName1		'�Ŋ�w�P
	Dim dbToStation1		'�Ŋ�w�P�����Ђ܂ł̏��v����
	Dim dbToStationRemark1	'�Ŋ�w�P�܂ł̌�ʎ�i
	Dim dbStationName2		'�Ŋ�w�Q
	Dim dbToStation2		'�Ŋ�w�Q�����Ђ܂ł̏��v����
	Dim dbToStationRemark2	'�Ŋ�w�Q�܂ł̌�ʎ�i

	Dim sNearbyStation		'�Ŋ�w
	Dim sClass				'�g�p����X�^�C���V�[�g�̃N���X�@�摜�̗L���ŕω�
	Dim sLineClass			'
	Dim flgLine				'�������t���O
	Dim sAddTitle			'�h����Ƃ̏��̏ꍇ�́u�h���v�����ږ��ɕt����

	If GetRSState(rRS) = False Then Exit Function

	dbPlanTypeName = rRS.Collect("PlanTypeName")
	dbImageLimit = rRS.Collect("ImageLimit")

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sCompanyCode = rRS.Collect("CompanyCode")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'******************************************************************************
	'��Ж� start
	'------------------------------------------------------------------------------
	sCompanyName = rRS.Collect("CompanyName")
	sCompanyNameF = rRS.Collect("CompanyName_F")
	sOrderType = rRS.Collect("OrderType")
	sCompanyPictureFlag = rRS.Collect("CompanyPictureFlag")
	sCompanyKbn = rRS.Collect("CompanyKbn")
	sCompanySpeciality = rRS.Collect("CompanySpeciality")

	If sOrderType = "0" And sCompanyKbn = "4" Then sAddTitle = "�h����Ƃ�"

	Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
	'------------------------------------------------------------------------------
	'��Ж� end
	'******************************************************************************

	'******************************************************************************
	'�ݗ��N�x start
	'------------------------------------------------------------------------------
	sEstablishYear = ""
	sEstablishYear = rRS.Collect("EstablishYear")
	If sEstablishYear <> "" Then sEstablishYear = sEstablishYear & "�N"
	'------------------------------------------------------------------------------
	'�ݗ��N�x end
	'******************************************************************************

	'******************************************************************************
	'���{�z start
	'------------------------------------------------------------------------------
	sCapitalAmount = ""
	sCapitalAmount = rRS.Collect("CapitalAmount")
	If IsNumeric(sCapitalAmount) = True Then sCapitalAmount = GetJapaneseYen(sCapitalAmount)
	'------------------------------------------------------------------------------
	'���{�z end
	'******************************************************************************

	'******************************************************************************
	'�������J start
	'------------------------------------------------------------------------------
	sListClass = ""
	sListClass = rRS.Collect("ListClass")
	'------------------------------------------------------------------------------
	'�������J end
	'******************************************************************************

	'******************************************************************************
	'�Ј��� start
	'------------------------------------------------------------------------------
	sEmployeeNum = ""
	If ChkStr(rRS.Collect("ManEmployeeNum")) <> "" Or ChkStr(rRS.Collect("WomanEmployeeNum")) <> "" Then
		If rRS.Collect("ManEmployeeNum") <> "" Then
			sEmployeeNum = sEmployeeNum & "�j��" & rRS.Collect("ManEmployeeNum") & "�l"
		End If
		If ChkStr(rRS.Collect("WomanEmployeeNum")) <> "" Then
			If sEmployeeNum <> "" Then sEmployeeNum = sEmployeeNum & "�@"
			sEmployeeNum = sEmployeeNum & "����" & rRS.Collect("WomanEmployeeNum") & "�l"
		End If
		sEmployeeNum = "(" & sEmployeeNum & ")"
	End If
	If ChkStr(rRS.Collect("AllEmployeeNum")) <> "" Then
		sEmployeeNum = rRS.Collect("AllEmployeeNum") & "�l" & sEmployeeNum
	End If
	'------------------------------------------------------------------------------
	'�Ј��� end
	'******************************************************************************

	'******************************************************************************
	'�Ǝ� start
	'------------------------------------------------------------------------------
	sIndustryType = ""
	sIndustryType = rRS.Collect("IndustryTypeName")
	'------------------------------------------------------------------------------
	'�������J end
	'******************************************************************************

	'******************************************************************************
	'�{�ЏZ�� start
	'------------------------------------------------------------------------------
	sAddress = ""
	If rRS.Collect("Post_U") & rRS.Collect("Post_L") <> "" Then
		sAddress = "��" & rRS.Collect("Post_U") & "-" & rRS.Collect("Post_L") & "<br>"
	End If
	sAddress = sAddress & rRS.Collect("Address")
	'------------------------------------------------------------------------------
	'�{�ЏZ�� end
	'******************************************************************************

	'******************************************************************************
	'�z�[���y�[�W start
	'------------------------------------------------------------------------------
	sHomePage = ""
	If rRS.Collect("HomepageAddress") <> "" And sOrderType = "0" Then
		sHomePage = rRS.Collect("HomePageAddress")
	End If
	'------------------------------------------------------------------------------
	'�z�[���y�[�W end
	'******************************************************************************

	'******************************************************************************
	'�Ŋ�w start
	'------------------------------------------------------------------------------
	sSQL = "/* �i�r�F��Ə��y�[�W�̍Ŋ�w�擾 */"
	sSQL = sSQL & "EXEC sp_GetDetailCompany '" & sCompanyCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sNearbyStation = ""
		dbStationName1 = ChkStr(oRS.Collect("StationName1"))
		dbStationName2 = ChkStr(oRS.Collect("StationName2"))
		If dbStationName1 & dbStationName2 <> "" And sOrderType = "0" Then
			dbToStation1 = ChkStr(oRS.Collect("WorkOrBus1"))
			dbToStationRemark1 = ChkStr(oRS.Collect("CompanySyudan1_1"))
			dbToStation2 = ChkStr(oRS.Collect("WorkOrBus2"))
			dbToStationRemark2 = ChkStr(oRS.Collect("CompanySyudan2_1"))

			If dbStationName1 <> "" Then
				If sNearbyStation <> "" Then sNearbyStation = sNearbyStation & "<br>"

				sNearbyStation = sNearbyStation & dbStationName1 & "�w"
				If dbToStation1 <> "" Then sNearbyStation = sNearbyStation & "(" & dbToStationRemark1 & dbToStation1 & "��)"
			End If

			If dbStationName2 <> "" Then
				If sNearbyStation <> "" Then sNearbyStation = sNearbyStation & "<br>"

				sNearbyStation = sNearbyStation & dbStationName2 & "�w"
				If dbToStation2 <> "" Then sNearbyStation = sNearbyStation & "(" & dbToStationRemark2 & dbToStation2 & "��)"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'�Ŋ�w end
	'******************************************************************************

	If sCompanyPictureFlag = "1" And dbImageLimit > 0 Then
		sClass = "value1"
		sLineClass = "odline2"
	Else
		sClass = "value2"
		sLineClass = "odline1"
	End If

	flgLine = False
	Response.Write "<div class=""companyblock"">"
	Response.Write "<h3>" & sAddTitle & "��Ə��</h3>"
	'If sCompanyPictureFlag = "1" And dbImageLimit > 0 Then
'	Response.Write "<div id=""imgcompany""><img id=""imgcompany"" src=""" & HTTPS_NAVI_CURRENTURL & "company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=1"" alt=""�C���[�W�ʐ^""></div>"
'	End If
		
		%>
        <table id="company_code">
            <tbody>
            	<tr>
                	<th>��ƃR�[�h</th>
                    <td><%= sCompanyCode %></td>
            	</tr>
                <tr>
                	<th>�ݗ��N�x</th>
                    <td><%= sEstablishYear %></td>
                </tr>
                <tr>
                	<th>���{�z</th>
                    <td><%= sCapitalAmount %></td>
                </tr>
                <tr>
                	<th>�������J</th>
                    <td><%= sListClass %></td>
                </tr>
                <tr>
                	<th>�Ј���</th>
                    <td><%= sEmployeeNum %></td>
                </tr>
                <tr>
                	<th>�Ǝ�</th>
                    <td><%= sIndustryType %></td>
                </tr>
                <tr>
                	<th>�{�ЏZ��</th>
                    <td><%= sAddress %></td>
                </tr>
                <tr>
                	<th>�{�ЍŊ�w</th>
                    <td><%= sNearbyStation %></td>
                </tr>

                <!-- 2014/04/22 �r�c ���HP���牞�債�Ă��܂��\��������̂ŁA�R�����g�A�E�g
                <tr>
                	<th>�R�[�|���[�g�T�C�g</th>
                    <td>
                    <%
                        if replace(sHomePage," ","") <> "" then
                        %>
                            <a href="<%= sHomePage %>" target="_blank" rel="nofollow">���̊�Ƃ̃z�[���y�[�W</a>
                        <%
                        else
                            
                        end if
                    %>
                    </td>
                </tr>
                -->
                 
            </tbody>
        </table>
		<%


	Response.Write "</div>"

End Function

'******************************************************************************
'�T�@�v�F��Ə��̊�{�����o��(NEO��p)
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvOrderCode		�F
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�����ƃi�r/order/company_order.asp
'���@�l�F
'���@���F2007/02/11 LIS K.Kokubo �쐬
'�@�@�@�F2008/06/25 LIS K.Kokubo �Ŋ�w�ǉ�
'******************************************************************************
Function DspCompanyInfoNEO(ByRef rDB, ByRef rRS, ByVal vOrderCode, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbPlanTypeName		'���C�Z���X�v�����^�C�v
	Dim dbImageLimit		'�ő�摜�f�ڐ�
	Dim sCompanyCode		'��ƃR�[�h
	Dim sCompanyName		'��Ɩ���
	Dim sCompanyNameF		'��Ɩ��̃J�i
	Dim sOrderType			'���l��� ["0"]�����ƃi�r��� ["1"]�h�� ["2"]�Љ� ["3"]
	Dim sCompanyPictureFlag	'��Ǝʐ^�t���O ["1"]�L ["0"]��
	Dim sCompanyKbn			'��Ƌ敪
	Dim sCompanySpeciality	'��Ɠ���
	Dim sEstablishYear		'�ݗ��N�x
	Dim sCapitalAmount		'���{�z
	Dim sListClass			'�������J
	Dim sEmployeeNum		'�Ј���
	Dim sIndustryType		'�Ǝ�
	Dim sAddress			'�{�ЏZ��
	Dim sHomePage			'�z�[���y�[�W
	Dim dbStationName1		'�Ŋ�w�P
	Dim dbToStation1		'�Ŋ�w�P�����Ђ܂ł̏��v����
	Dim dbToStationRemark1	'�Ŋ�w�P�܂ł̌�ʎ�i
	Dim dbStationName2		'�Ŋ�w�Q
	Dim dbToStation2		'�Ŋ�w�Q�����Ђ܂ł̏��v����
	Dim dbToStationRemark2	'�Ŋ�w�Q�܂ł̌�ʎ�i

	Dim sNearbyStation		'�Ŋ�w
	Dim sClass				'�g�p����X�^�C���V�[�g�̃N���X�@�摜�̗L���ŕω�
	Dim sLineClass			'
	Dim flgLine				'�������t���O
	Dim sAddTitle			'�h����Ƃ̏��̏ꍇ�́u�h���v�����ږ��ɕt����

	If GetRSState(rRS) = False Then Exit Function

	dbPlanTypeName = rRS.Collect("PlanTypeName")
	dbImageLimit = rRS.Collect("ImageLimit")

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sCompanyCode = rRS.Collect("CompanyCode")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'******************************************************************************
	'��Ж� start
	'------------------------------------------------------------------------------
	sCompanyName = rRS.Collect("CompanyName")
	sCompanyNameF = rRS.Collect("CompanyName_F")
	sOrderType = rRS.Collect("OrderType")
	sCompanyPictureFlag = rRS.Collect("CompanyPictureFlag")
	sCompanyKbn = rRS.Collect("CompanyKbn")
	sCompanySpeciality = rRS.Collect("CompanySpeciality")

	If sOrderType = "0" And sCompanyKbn = "4" Then sAddTitle = "�h����Ƃ�"

	Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
	'------------------------------------------------------------------------------
	'��Ж� end
	'******************************************************************************

	'******************************************************************************
	'�ݗ��N�x start
	'------------------------------------------------------------------------------
	sEstablishYear = ""
	sEstablishYear = rRS.Collect("EstablishYear")
	If sEstablishYear <> "" Then sEstablishYear = sEstablishYear & "�N"
	'------------------------------------------------------------------------------
	'�ݗ��N�x end
	'******************************************************************************

	'******************************************************************************
	'���{�z start
	'------------------------------------------------------------------------------
	sCapitalAmount = ""
	sCapitalAmount = rRS.Collect("CapitalAmount")
	If IsNumeric(sCapitalAmount) = True Then sCapitalAmount = GetJapaneseYen(sCapitalAmount)
	'------------------------------------------------------------------------------
	'���{�z end
	'******************************************************************************

	'******************************************************************************
	'�������J start
	'------------------------------------------------------------------------------
	sListClass = ""
	sListClass = rRS.Collect("ListClass")
	'------------------------------------------------------------------------------
	'�������J end
	'******************************************************************************

	'******************************************************************************
	'�Ј��� start
	'------------------------------------------------------------------------------
	sEmployeeNum = ""
	If ChkStr(rRS.Collect("ManEmployeeNum")) <> "" Or ChkStr(rRS.Collect("WomanEmployeeNum")) <> "" Then
		If rRS.Collect("ManEmployeeNum") <> "" Then
			sEmployeeNum = sEmployeeNum & "�j��" & rRS.Collect("ManEmployeeNum") & "�l"
		End If
		If ChkStr(rRS.Collect("WomanEmployeeNum")) <> "" Then
			If sEmployeeNum <> "" Then sEmployeeNum = sEmployeeNum & "�@"
			sEmployeeNum = sEmployeeNum & "����" & rRS.Collect("WomanEmployeeNum") & "�l"
		End If
		sEmployeeNum = "(" & sEmployeeNum & ")"
	End If
	If ChkStr(rRS.Collect("AllEmployeeNum")) <> "" Then
		sEmployeeNum = rRS.Collect("AllEmployeeNum") & "�l" & sEmployeeNum
	End If
	'------------------------------------------------------------------------------
	'�Ј��� end
	'******************************************************************************

	'******************************************************************************
	'�Ǝ� start
	'------------------------------------------------------------------------------
	sIndustryType = ""
	sIndustryType = rRS.Collect("IndustryTypeName")
	'------------------------------------------------------------------------------
	'�������J end
	'******************************************************************************

	'******************************************************************************
	'�{�ЏZ�� start
	'------------------------------------------------------------------------------
	sAddress = ""
	If rRS.Collect("Post_U") & rRS.Collect("Post_L") <> "" Then
		sAddress = "��" & rRS.Collect("Post_U") & "-" & rRS.Collect("Post_L") & "<br>"
	End If
	sAddress = sAddress & rRS.Collect("Address")
	'------------------------------------------------------------------------------
	'�{�ЏZ�� end
	'******************************************************************************

	'******************************************************************************
	'�z�[���y�[�W start
	'------------------------------------------------------------------------------
	sHomePage = ""
	If rRS.Collect("HomepageAddress") <> "" And sOrderType = "0" Then
		sHomePage = rRS.Collect("HomePageAddress")
	End If
	'------------------------------------------------------------------------------
	'�z�[���y�[�W end
	'******************************************************************************

	'******************************************************************************
	'�Ŋ�w start
	'------------------------------------------------------------------------------
	sSQL = "/* �i�r�F��Ə��y�[�W�̍Ŋ�w�擾 */"
	sSQL = sSQL & "EXEC sp_GetDetailCompany '" & sCompanyCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sNearbyStation = ""
		dbStationName1 = ChkStr(oRS.Collect("StationName1"))
		dbStationName2 = ChkStr(oRS.Collect("StationName2"))
		If dbStationName1 & dbStationName2 <> "" And sOrderType = "0" Then
			dbToStation1 = ChkStr(oRS.Collect("WorkOrBus1"))
			dbToStationRemark1 = ChkStr(oRS.Collect("CompanySyudan1_1"))
			dbToStation2 = ChkStr(oRS.Collect("WorkOrBus2"))
			dbToStationRemark2 = ChkStr(oRS.Collect("CompanySyudan2_1"))

			If dbStationName1 <> "" Then
				If sNearbyStation <> "" Then sNearbyStation = sNearbyStation & "<br>"

				sNearbyStation = sNearbyStation & dbStationName1 & "�w"
				If dbToStation1 <> "" Then sNearbyStation = sNearbyStation & "(" & dbToStationRemark1 & dbToStation1 & "��)"
			End If

			If dbStationName2 <> "" Then
				If sNearbyStation <> "" Then sNearbyStation = sNearbyStation & "<br>"

				sNearbyStation = sNearbyStation & dbStationName2 & "�w"
				If dbToStation2 <> "" Then sNearbyStation = sNearbyStation & "(" & dbToStationRemark2 & dbToStation2 & "��)"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'�Ŋ�w end
	'******************************************************************************

	If sCompanyPictureFlag = "1" And dbImageLimit > 0 Then
		sClass = "value1"
		sLineClass = "odline2"
	Else
		sClass = "value2"
		sLineClass = "odline1"
	End If

	flgLine = False
	Response.Write "<div class=""companyblock"">"
	Response.Write "<h3>" & sAddTitle & "��Ə��</h3>"
	'If sCompanyPictureFlag = "1" And dbImageLimit > 0 Then
'	Response.Write "<div id=""imgcompany""><img id=""imgcompany"" src=""" & HTTPS_NAVI_CURRENTURL & "company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=1"" alt=""�C���[�W�ʐ^""></div>"
'	End If
		
		%>
        <table id="company_code">
            <tbody>
            	<tr>
                	<th>��Ɩ�</th>
                    <td><%= sCompanyName %></td>
            	</tr>
                <tr>
                	<th>�ݗ��N�x</th>
                    <td><%= sEstablishYear %></td>
                </tr>
                <tr>
                	<th>���{�z</th>
                    <td><%= sCapitalAmount %></td>
                </tr>
                <tr>
                	<th>�������J</th>
                    <td><%= sListClass %></td>
                </tr>
                <tr>
                	<th>�Ј���</th>
                    <td><%= sEmployeeNum %></td>
                </tr>
                <tr>
                	<th>�Ǝ�</th>
                    <td><%= sIndustryType %></td>
                </tr>
                <tr>
                	<th>�{�ЏZ��</th>
                    <td><%= sAddress %></td>
                </tr>
                <tr>
                	<th>�{�ЍŊ�w</th>
                    <td><%= sNearbyStation %></td>
                </tr>

                <!-- 2014/04/22 �r�c ���HP���牞�債�Ă��܂��\��������̂ŁA�R�����g�A�E�g
                <tr>
                	<th>�R�[�|���[�g�T�C�g</th>
                    <td>
                    <%
                        if replace(sHomePage," ","") <> "" then
                        %>
                            <a href="<%= sHomePage %>" target="_blank" rel="nofollow">���̊�Ƃ̃z�[���y�[�W</a>
                        <%
                        else
                            
                        end if
                    %>
                    </td>
                </tr>
                -->
                 
            </tbody>
        </table>
		<%


	Response.Write "</div>"

End Function

'******************************************************************************
'�T�@�v�F��Ə��̂o�q�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvOrderCode		�F
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�p���F�����ƃi�r/order/company_order.asp
'���@�l�F
'���@���F2007/02/11 LIS K.Kokubo �쐬
'�@�@�@�F2009/01/06 LIS K.Kokubo �����������l�ǉ�
'******************************************************************************
Function DspCompanyPR(ByRef rDB, ByRef rRS, ByVal vOrderCode, ByVal vUserType, ByVal vUserID)
	Const WELFARECOL = "3"	'���������̂P�s������̗�

	Dim sOrderType			'�󒍎��
	Dim sCompanyKbn			'��Ƌ敪
	Dim sBusiness			'���Ɠ��e
	Dim sPR					'��ƏЉ�
	Dim sAtmosphere			'��Ђ̕��͋C
	Dim sWelfare			'��������
	Dim iWelfare			'���������J�E���g
	Dim sWelfareProgramRemark'�����������l
	Dim idx
	Dim flgPR
	Dim flgLine				'�������t���O
	Dim sClass
	Dim sAddTitle			'�h����Ƃ̏��̏ꍇ�́u�h����Ƃ́v�����ږ��ɕt����

	If GetRSState(rRS) = False Then Exit Function

	sOrderType = rRS.Collect("OrderType")
	sCompanyKbn = rRS.Collect("CompanyKbn")

	If sOrderType = "0" And sCompanyKbn = "4" Then sAddTitle = "�h����Ƃ�"

	'<���Ɠ��e>
	sBusiness = ""
	sBusiness = Replace(ChkStr(rRS.Collect("BusinessContents")), vbCrLf, "<br>")
	sBusiness = Replace(sBusiness, vbCr, "<br>")
	sBusiness = Replace(sBusiness, vbLf, "<br>")
	'</���Ɠ��e>

	'<��ЏЉ�>
	sPR = ""
	sPR = Replace(ChkStr(rRS.Collect("CompanyPR")), vbCrLf, "<br>")
	sPR = Replace(sPR, vbCr, "<br>")
	sPR = Replace(sPR, vbLf, "<br>")
	'</��ЏЉ�>

	'<��Ђ̕��͋C>
	sAtmosphere = ""
	sAtmosphere = Replace(ChkStr(rRS.Collect("Atmosphere")), vbCrLf, "<br>")
	sAtmosphere = Replace(sAtmosphere, vbCr, "<br>")
	sAtmosphere = Replace(sAtmosphere, vbLf, "<br>")
	'</��Ђ̕��͋C>

	'<��������>
	sWelfare = ""
	iWelfare = 0

	If ChkStr(rRS.Collect("SocietyInsuranceFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�Љ�ی�����</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("SanatoriumFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�ۗ{��</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("EnterprisePensionFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">��ƔN��</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("WealthShapeFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">���`���~</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("StockOptionFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�������x(�X�g�b�N�I�v�V����)</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("RetirementPayFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�ސE�����x</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("ResidencePayFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�Z��蓖</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("FamilyPayFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�Ƒ��蓖</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("EmployeeDormitoryFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�Ј���</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("CompanyHouseFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�Б�</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("NewEmployeeTrainingFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�V���Ј����C</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("OverseasTrainingFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�C�O���C</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	If ChkStr(rRS.Collect("OtherTrainingFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�e�팤�C</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

	'**TOP 08/08/19 Lis�� DEL
	'If ChkStr(rRS.Collect("FlexTimeFlag")) = "1" Then
	'	iWelfare = iWelfare + 1
	'	If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
	'	sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�t���b�N�X�^�C��</p></td>"
	'	If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	'End If
	'**BTM 08/08/19 Lis�� DEL

	'���r���[�ŏI������ꍇ�̒���
	If iWelfare Mod WELFARECOL <> 0 Then
		For idx = 1 To (WELFARECOL - iWelfare Mod WELFARECOL)
			sWelfare = sWelfare & "<td class=""welfare""></td>"
		Next
		sWelfare = sWelfare & "</tr>"
	End If

	If sWelfare <> "" Then
		sWelfare = "<table class=""welfare"">" & sWelfare & "</table>"
	End If

	sWelfareProgramRemark = Replace(ChkStr(rRS.Collect("WelfareProgramRemark")),VbCrLf,"<br>")
	'<��������>
	
	%>
    <div class="companyblock">
        <h3><%= sAddTitle %>PR���</h3>
        <table id="company_code">
            <tbody>
                <tr>
                    <th>���Ɠ��e</th>
                    <td><%= sBusiness %></td>
                </tr>
                <tr>
                    <th>���PR</th>
                    <td><%= sPR %></td>
            	</tr>
                <tr>
                    <th>��Ђ̕��͋C</th>
                    <td><%= sAtmosphere %></td>
            	</tr>
                <tr>
                    <th>��������</th>
                    <td><%= sWelfare & sWelfareProgramRemark %></td>
            	</tr>
            </tbody>
        </table>
	</div>
	<%

End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̃��X�̏Љ��E�h�����Ə����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'******************************************************************************
Function DspLisOrderCompanyInfo(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sCompanyCode		'��ƃR�[�h
	Dim sOrderType			'�󒍋敪
	Dim sListClass			'�������J
	Dim sIndustryType		'�Ǝ�
	Dim sCapitalAmount		'���{�z		'**TOP 08/08/21 Lis�� ADD
	Dim sEmployeeNum		'�Ј���
	Dim sAccountingPeriod1	'���Z��1
	Dim sSalesAmount1		'���㍂1
	Dim sOrdinaryProfit1	'�o�험�v1
	Dim sAccountingPeriod2	'���Z��2
	Dim sSalesAmount2		'���㍂2
	Dim sOrdinaryProfit2	'�o�험�v2
	Dim sAccountingPeriod3	'���Z��3
	Dim sSalesAmount3		'���㍂3
	Dim sOrdinaryProfit3	'�o�험�v3
	Dim sImportantNotice	'���L����
	Dim sflgAct							'**BTM 08/08/21 Lis�� ADD
	Dim sPR					'���Ɠ��e
	Dim sImgTitle			'�^�C�g���C���[�W
	Dim sIntrDisp			'�h�� or �Љ��
	Dim flgDsp
	Dim flgLine				'�������t���O

	DspLisOrderCompanyInfo = False

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'��ƃR�[�h start
		'------------------------------------------------------------------------------
		sCompanyCode = rRS.Collect("CompanyCode")
		sOrderType = rRS.Collect("OrderType")
		If sOrderType = "2" Then
			sImgTitle = "/img/order/lisorderpr2.gif"
			
		Else
			sImgTitle = "/img/order/lisorderpr1.gif"
			
		End If
		'------------------------------------------------------------------------------
		'��ƃR�[�h end
		'******************************************************************************

		'******************************************************************************
		'�������J start
		'------------------------------------------------------------------------------
		sListClass = ""
		sListClass = rRS.Collect("ListClass")
		'------------------------------------------------------------------------------
		'�������J end
		'******************************************************************************

		'******************************************************************************
		'�Ǝ� start
		'------------------------------------------------------------------------------
		sIndustryType = ""
		sIndustryType = ChkStr(rRS.Collect("IndustryTypeName"))
		'------------------------------------------------------------------------------
		'�Ǝ� end
		'******************************************************************************

		'******************************************************************************
		'��ЏЉ� start
		'------------------------------------------------------------------------------
		sPR = ""
		sPR = Replace(ChkStr(rRS.Collect("BusinessContents")), vbCrLf, "<br>")
		sPR = Replace(sPR, vbCr, "<br>")
		sPR = Replace(sPR, vbLf, "<br>")
		'------------------------------------------------------------------------------
		'��ЏЉ� end
		'******************************************************************************
		'**TOP 08/08/21 Lis�� ADD
		'******************************************************************************
		'���{�z start
		'------------------------------------------------------------------------------
		sCapitalAmount = ""
		sCapitalAmount = ChkStr(rRS.Collect("CapitalAmount"))
		if IsNumeric(sCapitalAmount) = True then
			sCapitalAmount = GetJapaneseYen(sCapitalAmount)
		elseif sCapitalAmount <> "" then
			if InStr(sCapitalAmount,"�~") > 0 then		'"�~"�������Ă����炻�̂܂�
			else
				sCapitalAmount = sCapitalAmount & "�~"
			end if
		end if
		'------------------------------------------------------------------------------
		'���{�z end
		'******************************************************************************

		'******************************************************************************
		'�Ј��� start
		'------------------------------------------------------------------------------
		sEmployeeNum = ""
		sEmployeeNum = ChkStr(rRS.Collect("AllEmployeeNum"))
		If sEmployeeNum <> "" Then sEmployeeNum = sEmployeeNum & "�l"
		'------------------------------------------------------------------------------
		'�Ј��� end
		'******************************************************************************
		
		'******************************************************************************
		'���Z���E���㍂�E�o�험�v start
		'------------------------------------------------------------------------------
		sAccountingPeriod1 = ""
		sSalesAmount1 = ""
		sOrdinaryProfit1 = ""
		sAccountingPeriod2 = ""
		sSalesAmount2 = ""
		sOrdinaryProfit2 = ""
		sAccountingPeriod3 = ""
		sSalesAmount3 = ""
		sOrdinaryProfit3 = ""
		sImportantNotice = ""
		sAccountingPeriod1 = ChkStr(rRS.Collect("AccountingPeriod1"))
		sSalesAmount1 = ChkStr(rRS.Collect("SalesAmount1"))
		'if sSalesAmount1 <> "" and InStr(sSalesAmount1,"�~") <= 0 then sSalesAmount1 = sSalesAmount1 & "�~"
		sOrdinaryProfit1 = ChkStr(rRS.Collect("OrdinaryProfit1"))
		'if sOrdinaryProfit1 <> "" and InStr(sOrdinaryProfit1,"�~") <= 0 then sOrdinaryProfit1 = sOrdinaryProfit1 & "�~"
		sAccountingPeriod2 = ChkStr(rRS.Collect("AccountingPeriod2"))
		sSalesAmount2 = ChkStr(rRS.Collect("SalesAmount2"))
		'if sSalesAmount2 <> "" and InStr(sSalesAmount2,"�~") <= 0 then sSalesAmount2 = sSalesAmount2 & "�~"
		sOrdinaryProfit2 = ChkStr(rRS.Collect("OrdinaryProfit2"))
		'if sOrdinaryProfit2 <> "" and InStr(sOrdinaryProfit2,"�~") <= 0 then sOrdinaryProfit2 = sOrdinaryProfit2 & "�~"
		sAccountingPeriod3 = ChkStr(rRS.Collect("AccountingPeriod3"))
		sSalesAmount3 = ChkStr(rRS.Collect("SalesAmount3"))
		'if sSalesAmount3 <> "" and InStr(sSalesAmount3,"�~") <= 0 then sSalesAmount3 = sSalesAmount3 & "�~"
		sOrdinaryProfit3 = ChkStr(rRS.Collect("OrdinaryProfit3"))
		'if sOrdinaryProfit3 <> "" and InStr(sOrdinaryProfit3,"�~") <= 0 then sOrdinaryProfit3 = sOrdinaryProfit3 & "�~"
		sImportantNotice = ChkStr(rRS.Collect("ImportantNotice"))
		'------------------------------------------------------------------------------
		'���Z���E���㍂�E�o�험�v end
		'******************************************************************************
		'**BTM 08/08/21 Lis�� ADD
	End If

	flgLine = False

	'**TOP 08/08/21 Lis�� REP
	'If sListClass & sIndustryType & sPR <> "" Then
	If sListClass & sIndustryType & sPR & sCapitalAmount & sEmployeeNum <> "" or _
		(InStr(sImportantNotice,"����J") <= 0 and _
		((sAccountingPeriod1 <> "" and sSalesAmount1 <> "" and InStr(sAccountingPeriod1 & sSalesAmount1,"����J") <= 0) or _
		 (sAccountingPeriod2 <> "" and sSalesAmount2 <> "" and InStr(sAccountingPeriod2 & sSalesAmount2,"����J") <= 0) or _
		 (sAccountingPeriod3 <> "" and sSalesAmount3 <> "" and InStr(sAccountingPeriod3 & sSalesAmount3,"����J") <= 0))) Then
	'**BTM 08/08/21 Lis�� REP
		DspLisOrderCompanyInfo = True
%>
	<img src="/img/order/tab_detail_sk.png" class="tab_img">
    <table class="detail_table">
	<tbody>
    <tr>
    <th class="fast_th"><%= sIntrDisp %></th>
    <td>
    </td>
    </tr>
    <% If sListClass <> "" Then %>
    <tr>
    <th class="dborder_bottom">�������J</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sListClass %></p>
    </td>
    <% End If %>
    
	<% If sIndustryType <> "" Then %>
    <tr>
    <th class="dborder_bottom">�Ǝ�</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sIndustryType %></p>
    </td>
	</tr>     
   	<% End If %>
    
   	<% If sPR <> "" Then %>
    <tr>
    <th class="dborder_bottom">���Ɠ��e</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sPR %></p>
    </td>
	</tr>     
   	<% End If %>
    
   	<% If sCapitalAmount <> "" Then %>
    <tr>
    <th class="dborder_bottom">���{��</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sCapitalAmount %></p>
    </td>
	</tr>     
   	<% End If %>    

   	<% If sEmployeeNum <> "" Then %>
    <tr>
    <th class="dborder_bottom">�Ј���</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sEmployeeNum %></p>
    </td>
	</tr>     
   	<% End If %> 
    
   	<% sflgAct = ""
		If InStr(sImportantNotice,"����J") <= 0 and _
		((sAccountingPeriod1 <> "" and sSalesAmount1 <> "" and InStr(sAccountingPeriod1 & sSalesAmount1,"����J") <= 0) or _
		 (sAccountingPeriod2 <> "" and sSalesAmount2 <> "" and InStr(sAccountingPeriod2 & sSalesAmount2,"����J") <= 0) or _
		 (sAccountingPeriod3 <> "" and sSalesAmount3 <> "" and InStr(sAccountingPeriod3 & sSalesAmount3,"����J") <= 0)) then %>
    <tr>
    <th class="dborder_bottom">�������</th>
    <td class="dborder_bottom">
    <p class="m0"><% 			'���㍂�P�E�o�험�v�P�E���Z���P
			if sAccountingPeriod1 <> "" and sSalesAmount1 <> "" and InStr(sAccountingPeriod1 & sSalesAmount1,"����J") <= 0 then
				if sSalesAmount1 <> "" and InStr(sSalesAmount1,"����J") <= 0 then
					response.write "���㍂�F" & sSalesAmount1 & "�@"
				end if
				if sOrdinaryProfit1 <> "" and InStr(sOrdinaryProfit1,"����J") <= 0 then
					response.write "�o�험�v�F" & sOrdinaryProfit1
				end if
				if sAccountingPeriod1 <> "" and InStr(sAccountingPeriod1,"����J") <= 0 then
					response.write "�i���Z���F" & sAccountingPeriod1 & "�j<br>"
				end if
				sflgAct = "1"
			end if
			'���㍂�Q�E�o�험�v�Q�E���Z���Q
			if sAccountingPeriod2 <> "" and sSalesAmount2 <> "" and InStr(sAccountingPeriod2 & sSalesAmount2,"����J") <= 0 then
				if sSalesAmount2 <> "" and InStr(sSalesAmount2,"����J") <= 0 then
					response.write "���㍂�F" & sSalesAmount2 & "�@"
				end if
				if sOrdinaryProfit2 <> "" and InStr(sOrdinaryProfit2,"����J") <= 0 then
					response.write "�o�험�v�F" & sOrdinaryProfit2
				end if
				if sAccountingPeriod2 <> "" and InStr(sAccountingPeriod2,"����J") <= 0 then
					response.write "�i���Z���F" & sAccountingPeriod2 & "�j<br>"
				end if
				sflgAct = "1"
			end if
			'���㍂�R�E�o�험�v�R�E���Z���R
			if sAccountingPeriod3 <> "" and sSalesAmount3 <> "" and InStr(sAccountingPeriod3 & sSalesAmount3,"����J") <= 0 then
				if sSalesAmount3 <> "" and InStr(sSalesAmount3,"����J") <= 0 then
					response.write "���㍂�F" & sSalesAmount3 & "�@"
				end if
				if sOrdinaryProfit3 <> "" and InStr(sOrdinaryProfit3,"����J") <= 0 then
					response.write "�o�험�v�F" & sOrdinaryProfit3
				end if
				if sAccountingPeriod3 <> "" and InStr(sAccountingPeriod3,"����J") <= 0 then
					response.write "�i���Z���F" & sAccountingPeriod3 & "�j<br>"
				end if
				sflgAct = "1"
			end if
			'���L����
			If sflgAct = "1" and sImportantNotice <> "" and InStr(sImportantNotice,"����J") <= 0 then
				response.write "�i"
				if InStr(sImportantNotice,"��") <= 0 then response.write "��"
				response.write  sImportantNotice & "�j<br>"
			End If
			 %></p>
    </td>
	</tr>     
   	<% End If %>     
       
    </tbody>
    </table>
<p class="m0">���l��<% If sOrderType = "2" Then response.write "�Љ�" Else response.write "�h��" End If %>�ł��ē����邨�d���̂��߁A�ڂ�����Џ��͉��̃{�^���₨�d�b�ȂǂŒ��ڂ��⍇�����������B</p>
<div class="to_top"><a class="stext_middle" href="#pagetop">���y�[�WTOP��</a></div> 

<%

	End If
End Function

'******************************************************************************
'�T�@�v�F�h����Ƃ̔h�����Ə����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvMyOrder		�F���Ћ��l�[�t���O
'���@�l�F
'�g�p���F�����ƃi�r/order/company_order.asp
'���@���F2007/02/11 LIS K.Kokubo �쐬
'******************************************************************************
Function DspTempOrderCompanyInfo(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder)
	Dim dbOrderCode			'���R�[�h
	Dim dbTempCompanyName
	Dim dbTempCompanyName_F
	Dim dbTempEstablishYear
	Dim dbTempIndustryTypeName
	Dim dbTempCapitalAmount
	Dim dbTempForeinCapital
	Dim dbTempListClass
	Dim dbTempAllEmployeeNumber
	Dim dbTempHomepageAddress
	Dim dbTempPost_U
	Dim dbTempPost_L
	Dim dbTempPrefectureCode
	Dim dbTempCity
	Dim dbTempCity_F
	Dim dbTempTown
	Dim dbTempAddress
	Dim dbTempTelephoneNumber

	Dim sClearSolid
	Dim flgLine				'�������t���O
	Dim flgData
	Dim sCapital
	Dim sTempAllEmployeeNumber

	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function

	flgData = False

	'<�h�����Ə��擾>
	dbOrderCode = ChkStr(rRS.Collect("OrderCode"))
	'dbTempCompanyName = ChkStr(rRS.Collect("TempCompanyName"))
	'dbTempCompanyName_F = ChkStr(rRS.Collect("TempCompanyName_F"))
	dbTempEstablishYear = ChkStr(rRS.Collect("TempEstablishYear"))
	dbTempIndustryTypeName = ChkStr(rRS.Collect("TempIndustryTypeName"))
	dbTempCapitalAmount = ChkStr(rRS.Collect("TempCapitalAmount"))
	dbTempForeinCapital = ChkStr(rRS.Collect("TempForeinCapital"))
	dbTempListClass = ChkStr(rRS.Collect("TempListClass"))
	dbTempAllEmployeeNumber = ChkStr(rRS.Collect("TempAllEmployeeNumber"))
	'dbTempHomepageAddress = ChkStr(rRS.Collect("TempHomepageAddress"))
	'dbTempPost_U = ChkStr(rRS.Collect("TempPost_U"))
	'dbTempPost_L = ChkStr(rRS.Collect("TempPost_L"))
	'dbTempPrefectureCode = ChkStr(rRS.Collect("TempPrefectureCode"))
	'dbTempCity = ChkStr(rRS.Collect("TempCity"))
	'dbTempCity_F = ChkStr(rRS.Collect("TempCity_F"))
	'dbTempTown = ChkStr(rRS.Collect("TempTown"))
	'dbTempAddress = ChkStr(rRS.Collect("TempAddress"))
	'dbTempTelephoneNumber = ChkStr(rRS.Collect("TempTelephoneNumber"))
	'</�h�����Ə��擾>

	'<�ݗ��N�x>
	If dbTempEstablishYear <> "" Then
		dbTempEstablishYear = dbTempEstablishYear & "�N"
		flgData = True
	End If
	'</�ݗ��N�x>

	'<�Ǝ�>
	If dbTempIndustryTypeName <> "" Then
		flgData = True
	End If
	'</�Ǝ�>

	'<���{>
	sCapital = ""
	If dbTempCapitalAmount & dbTempForeinCapital <> "" Then
		If dbTempCapitalAmount <> "" Then
			sCapital = sCapital & GetJapaneseYen(dbTempCapitalAmount)
		End If

		If dbTempForeinCapital <> "" Then
			sCapital = sCapital & "&nbsp;�i�O���F" & dbTempForeinCapital & "�j"
		End If

		flgData = True
	End If
	'</���{>

	'<����>
	If dbTempListClass <> "" Then
		flgData = True
	End If
	'</����>

	'<�Ј���>
	If dbTempAllEmployeeNumber <> "" Then
		sTempAllEmployeeNumber = dbTempAllEmployeeNumber & "�l"
		flgData = True
	End If
	'</�Ј���>

	flgLine = False

	If flgData = True Then
		sHTML = sHTML & "<h3>�h�����Ə��</h3>" & vbCrLf

		If dbTempEstablishYear <> "" Then
			If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True

			sHTML = sHTML & "<div class=""category1""><h4>�ݗ��N�x</h4></div>"
			sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbTempEstablishYear & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If dbTempIndustryTypeName <> "" Then
			If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True

			sHTML = sHTML & "<div class=""category1""><h4>�Ǝ�</h4></div>"
			sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbTempIndustryTypeName & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sCapital <> "" Then
			If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True

			sHTML = sHTML & "<div class=""category1""><h4>���{��</h4></div>"
			sHTML = sHTML & "<div class=""value1"">"
			sHTML = sHTML & "<p class=""m0"">" & sCapital & "</p>"
			sHTML = sHTML & "</div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If dbTempListClass <> "" Then
			If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True

			sHTML = sHTML & "<div class=""category1""><h4>����</h4></div>"
			sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbTempListClass & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sTempAllEmployeeNumber <> "" Then
			If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True

			sHTML = sHTML & "<div class=""category1""><h4>�Ј���</h4></div>"
			sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sTempAllEmployeeNumber & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		End If

		sHTML = sHTML & "<br>" & vbCrLf
	End If

	Response.Write sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[�R���g���[���{�^��
'���@���FrDB				�F�ڑ�����DBConnection
'�@�@�@�FrRS				�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType			�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID			�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvMyOrder			�F���Ћ��l�[���ۂ� ["1"]���Ћ��l�[ ["0"]���Ћ��l�[�łȂ�
'�@�@�@�FvJobTypeLimitFlag	�F�E�퐔���������z���Ă��Ȃ��� ["1"]OK ["0"]NO
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'���@���F2007/02/11 LIS K.Kokubo �쐬
'�@�@�@�F2009/03/11 LIS K.Kokubo �ύX �C���^�r���[�ҏW�{�^���̕\�����@�ύX(�i�r�������Ή�)
'******************************************************************************
Function DspOrderControlButton(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder, ByVal vJobTypeLimitFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbPlanTypeName		'���C�Z���X�v�������
	Dim sOrderCode
	Dim sCompanyCode		'��ƃR�[�h
	Dim sOrderType			'�󒍎��
	Dim sPermitFlag			'�f�ڋ��t���O
	Dim sPublicFlag			'�f�ڃt���O
	Dim sRiyoFlag			'�f�ڊJ�n��
	Dim sHakouFlag			'���p�J�n���i���C�Z���X�������j
	Dim dbSearchName		'�ۑ�����������
	Dim dbSearchParam		'�ۑ����������p�����[�^

	Dim flgAddWatchList
	Dim iMailTemplateCnt	'���[���e���v���[�g�̌���
	Dim sAncMT				'���[���e���v���[�g�ւ̃����N

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	dbPlanTypeName = rRS.Collect("PlanTypeName")
	sOrderCode = rRS.Collect("OrderCode")
	sCompanyCode = rRS.Collect("CompanyCode")
	sOrderType = rRS.Collect("OrderType")
	sPermitFlag = rRS.Collect("PermitFlag")
	sPublicFlag = rRS.Collect("PublicFlag")
	sRiyoFlag = rRS.Collect("RiyoFlag")
	sHakouFlag = rRS.Collect("HakouFlag")
	iMailTemplateCnt = rRS.Collect("MailTemplateCnt")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	flgAddWatchList = False
	sSQL = "EXEC up_ChkWatchListExists_Staff '" & vUserID & "', '" & sOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("ExistsFlag") = "1" Then flgAddWatchList = True
	End If
	Call RSClose(oRS)
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	
End Function

'******************************************************************************
'�T�@�v�F���l�[�R���g���[���{�^��2
'���@���FrDB				�F�ڑ�����DBConnection
'�@�@�@�FrRS				�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType			�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID			�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvMyOrder			�F���Ћ��l�[���ۂ� ["1"]���Ћ��l�[ ["0"]���Ћ��l�[�łȂ�
'�@�@�@�FvJobTypeLimitFlag	�F�E�퐔���������z���Ă��Ȃ��� ["1"]OK ["0"]NO
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'���@���F����{�^���̂�
'******************************************************************************
Function DspOrderControlButton2(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder, ByVal vJobTypeLimitFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbPlanTypeName		'���C�Z���X�v�������
	Dim sOrderCode
	Dim sCompanyCode		'��ƃR�[�h
	Dim sOrderType			'�󒍎��
	Dim sPermitFlag			'�f�ڋ��t���O
	Dim sPublicFlag			'�f�ڃt���O
	Dim sRiyoFlag			'�f�ڊJ�n��
	Dim sHakouFlag			'���p�J�n���i���C�Z���X�������j
	Dim dbSearchName		'�ۑ�����������
	Dim dbSearchParam		'�ۑ����������p�����[�^

	Dim flgAddWatchList
	Dim iMailTemplateCnt	'���[���e���v���[�g�̌���
	Dim sAncMT				'���[���e���v���[�g�ւ̃����N

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	dbPlanTypeName = rRS.Collect("PlanTypeName")
	sOrderCode = rRS.Collect("OrderCode")
	sCompanyCode = rRS.Collect("CompanyCode")
	sOrderType = rRS.Collect("OrderType")
	sPermitFlag = rRS.Collect("PermitFlag")
	sPublicFlag = rRS.Collect("PublicFlag")
	sRiyoFlag = rRS.Collect("RiyoFlag")
	sHakouFlag = rRS.Collect("HakouFlag")
	iMailTemplateCnt = rRS.Collect("MailTemplateCnt")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	flgAddWatchList = False
	sSQL = "EXEC up_ChkWatchListExists_Staff '" & vUserID & "', '" & sOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("ExistsFlag") = "1" Then flgAddWatchList = True
	End If
	Call RSClose(oRS)
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************
	
	Dim qsOrderCode				'�I�[�_�[�R�[�h(�󒍕\�ԍ�)
	Dim iDetail				'���l�[�ڍׂ���̃t���O
	
	qsOrderCode = GetForm("ordercode", 2)
	iDetail = GetForm("Detail", 2)
	
	
	
	
	
	
	
	
	If vUserType = "staff" Then
	
		'******************************************************************************
		'���O�C�����E�҂̏ꍇ start
		'------------------------------------------------------------------------------
'		Response.Write "<div id=""login_watch"">"
'		If rRS.Collect("NowPublicFlag") = "1" Then
'
'            if sOrderType = "0" Then
'			    Response.Write "<input type=""button"" value=""���̋��l�ɉ��僁�[���𑗐M����"" onclick=""contactCompanyAdv('');"" class=""obo"">"
'            Else
'                Response.Write "<input type=""button"" value=""���̋��l�ɉ��僁�[���𑗐M����"" onclick=""contactCompanyLis('');"" class=""obo"">"
'			    Response.Write "<input type=""button"" value=""���̋��l�̎��⃁�[���𑗐M����"" onclick=""contactCompany('1');"" class=""qmail"">"
'            End If
'
'
'			If flgAddWatchList = True Then
'				Response.Write "<span class=""m0 kentoZumi"">���̋��l�[�͊��ɂ��C�ɓ��胊�X�g�ɓo�^�ς݂ł�</span>"
'			Else
'				Response.Write "<input type=""button"" value=""���C�ɓ��胊�X�g"" onclick=""ListAdd()"" class=""kento"">"
'			End If
'			Response.Write "<br clear=""both"">"
'
'		Else
'			Response.Write "<div class=""description"" align=""center""><b>���̋��l�[�͌f�ڂ��I�����Ă��܂��B���[�����M�͂ł��܂���B</b></div>"
'		End If
'		Response.Write "</div>"
		'------------------------------------------------------------------------------
		'���O�C�����E�҂̏ꍇ end
		'******************************************************************************

		'******************************************************************************
		'���O�C�����E�҂̏ꍇ start
		'------------------------------------------------------------------------------
		Response.Write "<div id=""login_watch"">"
		Response.Write "<ul>"
		If rRS.Collect("NowPublicFlag") = "1" Then

            if sOrderType = "0" Then
			    Response.Write "<li><input type=""button"" value=""���̋��l�ɉ��僁�[���𑗐M����"" onclick=""contactCompanyAdv('');"" class=""obo""></li>"
            Else
                Response.Write "<li><input type=""button"" value=""���̋��l�ɉ��僁�[���𑗐M����"" onclick=""contactCompanyLis('');"" class=""obo""></li>"
			    Response.Write "<li><input type=""button"" value=""���̋��l�̎��⃁�[���𑗐M����"" onclick=""contactCompany('1');"" class=""qmail""></li>"
            End If


			If flgAddWatchList = True Then
				Response.Write "<li><span class=""m0 kentoZumi"">���̋��l�[�͊��ɂ��C�ɓ��胊�X�g�ɓo�^�ς݂ł�</span></li>"
			Else
				response.write "<li>"
				'2017/04/04 �ؑ��F���C�ɓ���ǉ���IE�̂݋@�\���Ȃ��̂��C���@��onSubmit="return Submit();"�͓��̓`�F�b�N�pJS�Ȃ̂ŊO���܂���
				'response.write "<form id=""frmSendMailJobOfferAddress"" method=""post"" action=""../staff/watchlist_register.asp"" onSubmit=""return Submit();"">"
				'Response.Write "<input type=""button"" value=""���C�ɓ��胊�X�g"" onclick=""document.forms.frmSendMailJobOfferAddress.submit();return false;"" class=""kento"">"
				'response.write "<input type=""hidden"" name=""CONF_OrderCode"" value='"& qsOrderCode &"'>"

				response.write "<form id=""frmSendMailJobOfferAddress"" method=""post"" action=""../staff/watchlist_register.asp"" onSubmit=""document.forms.frmSendMailJobOfferAddress.submit();return false;"">"
				Response.Write "<input type=""submit"" value=""���C�ɓ��胊�X�g"" class=""kento"">"
				response.write "<input type=""hidden"" name=""CONF_OrderCode"" value='"& qsOrderCode &"'>"

				response.write "</form></li>"
			End If
			

		Else
			Response.Write "<li id=""finKokoku""><div class=""description"" align=""center""><b>���̋��l�[�͌f�ڂ��I�����Ă��܂��B���[�����M�͂ł��܂���B</b></div></li>"
		End If
		Response.Write "<br clear=""both"">"
		Response.Write "</ul></div>"
		'------------------------------------------------------------------------------
		'���O�C�����E�҂̏ꍇ end
		'******************************************************************************

	End If

End Function



'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̃R���g���[���{�^���ŗ��p����javascript�̏o��
'�@�@�@�F���Ћ��l�[ or ���O�C�����̋��E�҂̏ꍇ�́A�ҏW�{�^�� or ���[�����M�{�^������������
'�@�@�@�Fjavascript���o��
'���@���FrDB				�F�ڑ�����DBConnection
'�@�@�@�FrRS				�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType			�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID			�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvMyOrder			�F���Ћ��l�[���ۂ� ["1"]���Ћ��l�[ ["0"]���Ћ��l�[�łȂ�
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'******************************************************************************
Function JSOrderControlButton(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sOrderCode

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'���R�[�h
		sOrderCode = rRS.Collect("OrderCode")
	End If

	If vUserType = "staff" Then
		'******************************************************************************
		'���O�C�����E�҂̏ꍇ start
		'------------------------------------------------------------------------------
		If rRS.Collect("PublicFlag") = "1" Then
%>
<script type="text/javascript">
function contactCompany(vflg) {
	var sQ = '';
	if(vflg){
		if(vflg.length > 0)sQ = 'q=1&';
	}
	MailWin = window.open('<%= HTTPS_NAVI_CURRENTURL %>staff/mailtocompany.asp?' + sQ + 'ordercode=<%= sOrderCode %>','_blank');
}
function contactCompanyAdv(vflg) {
    var sQ = '';
    if (vflg) {
        if (vflg.length > 0) sQ = 'q=1&';
    }

//  �ʃE�B���h�E�\�����瓯�E�B���h�E���y�[�W�J�ڂ֏C���@�r�c 2014/04/07    
//  MailWin = window.open('<%= HTTPS_NAVI_CURRENTURL %>staff/mailtocompanyAdv.asp?' + sQ + 'ordercode=<%= sOrderCode %>', '_blank');
    document.location = '<%= HTTPS_NAVI_CURRENTURL %>staff/mailtocompanyAdv.asp?' + sQ + 'ordercode=<%= sOrderCode %>';
    
}
function contactCompanyLis(vflg) {
    var sQ = '';
    if (vflg) {
        if (vflg.length > 0) sQ = 'q=1&';
    }
    MailWin = window.open('<%= HTTPS_NAVI_CURRENTURL %>staff/mailtocompanyLis.asp?' + sQ + 'ordercode=<%= sOrderCode %>', '_blank');
}
function ListAdd() {
    MailWin2 = window.open('<%= HTTPS_NAVI_CURRENTURL %>order/sendmail_jobofferaddress.asp?ordercode=<%= sOrderCode %>', 'list', 'menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=470');
}
</script>
<%
		End If
		'------------------------------------------------------------------------------
		'���O�C�����E�҂̏ꍇ end
		'******************************************************************************
	End If
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̃R���g���[���{�^���Ŏg�p����FORM�f�[�^���o��
'���@���FrDB				�F�ڑ�����DBConnection
'�@�@�@�FrRS				�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType			�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID			�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvMyOrder			�F���Ћ��l�[���ۂ� ["1"]���Ћ��l�[ ["0"]���Ћ��l�[�łȂ�
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'******************************************************************************
Function FrmOrderControlButton(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sOrderCode
	Dim sCompanyCode		'��ƃR�[�h
	Dim sOrderType

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'��ƃR�[�h start
		'------------------------------------------------------------------------------
		sOrderCode = rRS.Collect("OrderCode")
		sCompanyCode = rRS.Collect("CompanyCode")
		sOrderType = rRS.Collect("OrderType")
		'------------------------------------------------------------------------------
		'��ƃR�[�h end
		'******************************************************************************
	End If
%>
	<form id="frmMain" action="./" method="post">
	<input type="hidden" name="CONF_OrderCode" value="<%= sOrderCode %>">
	<input type="hidden" name="CONF_CompanyCode" value="<%= sCompanyCode %>">
	<input type="hidden" name="CONF_OrderType" value="<%= sOrderType %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="CONF_SearchMode" value="">
	</form>
<%
End Function

'******************************************************************************
'�T�@�v�F���l�[�̊�Ɩ��̂��o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'���@�l�F
'���@���F2007/02/11 LIS K.Kokubo �쐬
'�@�@�@�F2008/03/04 LIS K.Kokubo �f�ڏI������[RiyoToDate]��[DspPublicLimitDay]�ɕύX
'�@�@�@�F2009/03/18 LIS K.Kokubo vReplaceFlag�ǉ�
'******************************************************************************
Function DspOrderCompanyName(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vReplaceFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType
	Dim dbSecretFlag		'�V�[�N���b�g�t���O
	Dim sCompanyCode		'��ƃR�[�h
	Dim sCompanyName		'��Ɩ���
	Dim sCompanyNameF		'��Ɩ��̃J�i
	Dim sCompanyKbn			'��Ƌ敪
	Dim sCompanySpeciality	'��Ɠ���
	Dim sPublishLimitStr	'�f�ڊ����\���p������
	Dim sCautionStr			'�f�ڊ����\�����ӕ���������
	Dim dbTempOrderFlag		'�h���Č��t���O
	Dim dbTTPOrderFlag		'�Љ�\��h���Č��t���O
	Dim flgNowPublic		'���݌f�ڒ��̋��l�[���� '[True]�f�ڒ� [False]��f��

	Dim sUpdateDay
	Dim vAccessCount
	
	If GetRSState(rRS) = False Then Exit Function

	dbSecretFlag = rRS.Collect("SecretFlag")

	'******************************************************************************
	'��Ж� start
	'------------------------------------------------------------------------------
	sCompanyName = rRS.Collect("CompanyName")
	sCompanyNameF = rRS.Collect("CompanyName_F")
	sCompanyKbn = rRS.Collect("CompanyKbn")
	sCompanySpeciality = rRS.Collect("CompanySpeciality")
	sOrderType = rRS.Collect("OrderType")
	dbTempOrderFlag = rRS.Collect("TempOrderFlag")
	dbTTPOrderFlag = rRS.Collect("TTPOrderFlag")

	If vReplaceFlag = "1" Then
		Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
	End If
	'------------------------------------------------------------------------------
	'��Ж� end
	'******************************************************************************

	'******************************************************************************
	'���l�[�f�ڊ��� start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">�������O�Ɍf�ڏI������ꍇ������܂��B</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'�f�ڒ� or ��f��
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'�ЊO�Č��Ȃ�DspPublicLimitDay���A�Г��Č��Ȃ�PublicLimitDay��\��
	'�ЊO�Č� OrderType = 0
	'�Г��Č� OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
			'���C�Z���X�؂�̂Ƃ���"�f�ڏI��"�ƕ\��
			sPublishLimitStr = "�f�ڏI��"
			sCautionStr = ""
		Else
			sPublishLimitStr = "�펞��W��"
		End If
	End If
	'------------------------------------------------------------------------------
	'���l�[�f�ڊ��� end
	'******************************************************************************


%>
<!--
<div style="margin-bottom:10px;">
<%
	'���X�Љ�Č�,�l�މ�ЏЉ�Č��̏ꍇ�́u�]�E���k�Č��v�C���[�W��\��
	If sOrderType = "2" Or (sCompanyKbn = "2" And dbTempOrderFlag = "0" And dbTTPOrderFlag = "0") Then
		Response.Write "<div style=""text-align:left;padding-bottom:10px;font-size:8pt; color:#666666;"">"
		Response.Write "<img src=""/img/order/counselable_order.gif"" width=""150"" height=""25"" alt=""�]�E�x�����󂯂ĉ��傷�鋁�l�ł�"">"
		Response.Write "&nbsp;&nbsp;�m�l�މ�Ђ̓]�E�x�����󂯂ĉ���ł��鋁�l�n&nbsp;&nbsp;&nbsp;"
	else
		Response.Write "<div style=""font-size:8pt;color:#666666;text-align:right"">"
	end if


	'Twitter�t�H���[�{�^��
	if Request.ServerVariables("HTTPS") <> "on" then
		Response.Write "<div style=""float:right;width:160px"">"
		
		Select Case rRS.Collect("BranchCode")
			Case "OY"
				Response.Write "<a href=""http://"
				Response.Write "twitter.com/navi_okayama"" class=""twitter-follow-button"" data-show-count=""false"" data-lang=""ja"""
				Response.Write ">Follow @navi_okayama</a><script type=""text/javascript"" src="""
				Response.Write "http://platform.twitter.com/widgets.js"" charset=""utf-8""></script>"
			Case Else
				Response.Write "<a href=""http://"
				Response.Write "twitter.com/shigoto_navi"" class=""twitter-follow-button"" data-show-count=""false"" data-lang=""ja"""
				Response.Write ">Follow @shigoto_navi</a><script type=""text/javascript"" src="""
				Response.Write "http://platform.twitter.com/widgets.js"" charset=""utf-8""></script>"
		End Select
		Response.Write "</div>"
	end if
	'FaceBook�{�^��
	Response.write "<iframe src="""
	if Request.ServerVariables("HTTPS") = "on" then
		Response.Write "https://"
	else
		Response.Write "http://"
	end if
	Response.Write "www.facebook.com/plugins/like.php?href=http%3a%2f%2fwww.shigotonavi.co.jp%2forder%2forder_detail.asp%3fOrderCode%3d" & rRS.Collect("OrderCode")
	Response.Write "&amp;layout=button_count&amp;show_faces=true&amp;width=30&amp;action=like&amp;colorscheme=light&amp;height=21"
	Response.Write """ scrolling=""no"" frameborder=""0"" style=""float:right;border:none; overflow:hidden; width:80px; height:21px;"" allowTransparency=""true""></iframe>"

    'Google+1�{�^��
    Response.Write "<div style=""float:right"">"
    Response.Write "<g:plusone size=""medium"" count=""true"" href=""http://www.shigotonavi.co.jp/""></g:plusone>"
    Response.Write "</div>"


	Response.Write "</div>"

	'�V�[�N���b�g���l�̏ꍇ�́u�V�[�N���b�g���l�v�C���[�W��\��
	'If dbSecretFlag = "1" Then Response.Write "<img src=""/img/order/secret_order.gif"" width=""150"" height=""25"" alt=""���̋��l����X�J�E�g���󂯂��l�������{���ł��鋁�l�ł�"">"
	If dbSecretFlag = "1" Then Response.Write "<p class=""m0"" style=""color:#ff9933; font-weight:bold;"">���X�J�E�g���󂯂��l�������{���ł��鋁�l���ł��B</p>"

	If vUserType = "" Or vUserType = "staff" Then
		'�񃍃O�C�����A�X�^�b�t���O�C����

		If G_USERID <> "" And flgNowPublic = True Then
			'�����ƃi�r�Ƀ��O�C�����̏ꍇ�́A��Ɩ��{�f�ڊ����{���l�[�t�q�k���[�����M
%>
	<div style="width:400px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="width:200px; float:left;">
		<div style="float:right; padding:0px;">

<%
	if Request.ServerVariables("HTTPS") = "on" then
		Response.Write "<img src=""https://www.google.com/chart?chs=82x82&cht=qr&chl="
	else
		Response.Write "<img src=""http://chart.apis.google.com/chart?chs=82x82&cht=qr&chl="
	end if
	Response.Write Server.URLEncode("http://m.shigotonavi.jp/detail/JobDetail.asp?OrderCode=" & rRS.Collect("OrderCode") & "&an=QROD")
	Response.Write """ border=""0"" align=""right"" alt=""QRCode"">"
%>
		</div>
		<div style="text-align:right; font-size:11px; padding-top:6px;">
			<a href="<%= HTTPS_NAVI_CURRENTURL %>order/sendmail_jobofferaddress.asp?OrderCode=<% = rRS.Collect("OrderCode") %>&amp;detail=1" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=470');return false;"><img src="<%= HTTP_NAVI_CURRENTURL %>img/staff/mail/mailhei.gif" border="0" align="bottom" alt="���l�[�����[�����M"> ���l�[�����[�����M</a>
		</div>
		<p class="m0" style="text-align:right;padding:4px 0px 0px 0px;">�f�ڊ����F<%= sPublishLimitStr %></p>

		<%= sCautionStr %>
		<div style="clear:both;"></div>
	</div>
	<div style="clear:both;"></div>
<%
		ElseIf flgNowPublic = True Then
			'�����ƃi�r�ɔ񃍃O�C���̏ꍇ�́A��Ɩ��{�f�ڊ����{���l�[�t�q�k���[�����M
%>
	<div style="width:400px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="width:200px; float:left;">

<%
	if Request.ServerVariables("HTTPS") = "on" then
		Response.Write "<img src=""https://www.google.com/chart?chs=82x82&cht=qr&chl="
	else
		Response.Write "<img src=""http://chart.apis.google.com/chart?chs=82x82&cht=qr&chl="
	end if
	Response.Write Server.URLEncode("http://m.shigotonavi.jp/detail/JobDetail.asp?OrderCode=" & rRS.Collect("OrderCode") & "&an=QROD")
	Response.Write """ border=""0"" align=""right"" alt=""QRCode"">"
%>
		<div style="text-align:right; font-size:11px; padding-top:6px;"><a href="<%= HTTPS_NAVI_CURRENTURL %>order/sendmail_jobofferaddress.asp?OrderCode=<% = rRS.Collect("OrderCode") %>&amp;detail=1" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=380');return false;"><img src="<%= HTTP_NAVI_CURRENTURL %>img/staff/mail/mailhei.gif" border="0" align="bottom" alt="���l�[�����[�����M"> ���l�[�����[�����M</a></div>
		<p class="m0" style="text-align:right;padding:4px 0px 0px 0px;">�f�ڊ����F<%= sPublishLimitStr %></p>
		<%= sCautionStr %>
		<div style="clear:both;"></div>
	</div>
	<div style="clear:both;"></div>
<%
		Else
%>
	<div style="width:400px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="width:200px; float:left;">
		<p class="m0" style="text-align:right; padding-top:21px;">�f�ڊ����F<%= sPublishLimitStr %></p>
		<div style="clear:both;"></div>
	</div>
	<div style="clear:both;"></div>
<%
		End If
	Else
%>
	<div style="width:400px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="width:200px; float:left;">

<%
	if Request.ServerVariables("HTTPS") = "on" then
		Response.Write "<img src=""https://www.google.com/chart?chs=82x82&cht=qr&chl="
	else
		Response.Write "<img src=""http://chart.apis.google.com/chart?chs=82x82&cht=qr&chl="
	end if
	Response.Write Server.URLEncode("http://m.shigotonavi.jp/detail/JobDetail.asp?OrderCode=" & rRS.Collect("OrderCode") & "&an=QROD")
	Response.Write """ border=""0"" align=""right"" alt=""QRCode"">"
%>		<p class="m0" style="text-align:right; width:156px; padding-top:21px;">�f�ڊ����F<%= sPublishLimitStr %></p>

		<%= sCautionStr %>
		<div style="clear:both;"></div>
	</div>
	<div style="clear:both;"></div>
<%
	End If
%>
</div>-->


<%
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̉�Џ��E�E����E�C���^�r���[�؂�ւ��{�^���ƎQ�Ɖ񐔂��o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvType			�F�\�������̎�� ["0"]�E���� ["1"]��Џ�� ["2"]�C���^�r���[
'�@�@�@�FvAccessCount	�F�\�������l�[�̃A�N�Z�X��
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'******************************************************************************
Function DspOrderShowTypeSwitch(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vType, ByVal vAccessCount)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sOrderType
	Dim sJobTypeDetail
	Dim sUpdateDay
	Dim dbTopInterviewFlag
	Dim dbPlanType

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	sOrderType = rRS.Collect("OrderType")
	dbPlanType = ChkStr(rRS.Collect("PlanTypeName"))
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'��̓I�E�햼
	sJobTypeDetail = rRS.Collect("JobTypeDetail")
	'�X�V��
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")
	'�g�b�v�C���^�r���[
	dbTopInterviewFlag = rRS.Collect("TopInterviewFlag")

	If sJobTypeDetail <> "" Then sJobTypeDetail = sJobTypeDetail & "�̂��d�����ڍ�"

	Response.Write "<div id=""tab_switch"">"
	Response.Write "<div class=""left"">"
	If vType = "0" Then
		'�d������\�����̏ꍇ
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_orderdetail_on.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></div>"
		If sOrderType = "0" Then
			'��ʂ̋��l�L���̏ꍇ�͉�Џ��ւ̃����N��\��
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./company_order.asp?poc=" & sOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_companyinfo_off.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></a></div>"
		End If

		If sOrderType = "0" And dbTopInterviewFlag = "1" Then
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & sOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_interview_off.gif"" alt=""�C���^�r���["" border=""0"" width=""93"" height=""22""></a></div>"
		End If
	ElseIf vType = "1" Then
		'��Џ���\�����̏ꍇ
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & sOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
		If sOrderType = "0" Then
			'��ʂ̋��l�L���̏ꍇ�͉�Џ���\��
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_companyinfo_on.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></div>"
		End If

		If sOrderType = "0" And dbTopInterviewFlag = "1" Then
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & sOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_interview_off.gif"" alt=""�C���^�r���["" border=""0"" width=""93"" height=""22""></a></div>"
		End If

	ElseIf vType = "2" Then
		'�C���^�r���[��\�����̏ꍇ
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & sOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./company_order.asp?poc=" & sOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_companyinfo_off.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></a></div>"
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_interview_on.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></div>"
	End If

	Response.Write "</div>"


	Response.Write "<br clear=""both""></div>" & vbCrLf
End Function

'******************************************************************************
'�T�@�v�F���l�[�̃L���b�`�R�s�[�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F
'���@���F2007/02/11 LIS K.Kokubo �쐬
'�@�@�@�F2010/05/06 LIS K.Kokubo �Г��Č��p�ʐ^
'******************************************************************************
Function DspOrderCatchCopy(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode

	Dim sOptionNo			'�傫���ʐ^�̔ԍ�
	Dim sCompanyPictureFlag	'��Ǝʐ^�t���O ["1"]�L ["0"]��
	Dim sImg1
	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim vAccessCount
	Dim sPublishLimitStr
	
	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")

		sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'�傫���摜 start
	'------------------------------------------------------------------------------
	dbImageLimit = rRS.Collect("ImageLimit")
	sOptionNo = ""
	sImg1 = ""
	If dbImageLimit > 0 Then
		If dbImageLimit > 1 Then
			sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
					sOptionNo = oRS.Collect("OptionNo1")
					sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & sOptionNo
				End If
			End If
		End If

		If sImg1 = "" And dbOrderType = "0" Then
			sSQL = "sp_GetDataPicture '" & dbCompanyCode & "', '1'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'�傫���摜 end
	'******************************************************************************

	'<�Г��Č��p�ʐ^>
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			End If
		End If
		Call RSClose(oRS)
	End If
	'</�Г��Č��p�ʐ^>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)

	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"">"

		Response.Write "<div class=""main_pics""><div>"
		Response.Write "<img src=""" & sImg1 & """ alt="""" id=""big_pics"">"
		Response.Write "</div></div>"

		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"">" & rRS.Collect("CatchCopy") & "</p><br>"
		Response.Write "<div>"

		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">�y��W�̓����z</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write "</div>"

		%>
       		<div id="lissapo">
			<div><span>�]�E�T�|�[�g�Č�</span><br>
			�l�މ�Ђ̓]�E�x�����󂯂ĉ���ł��鋁�l
			</div>
			<p>�f�ڊ����F<%= sPublishLimitStr %><br>
			���ԉ{���񐔁F<%= vAccessCount %>��<br>
			�X�V���F<%= sUpdateDay %></p>
			<span>�����O�Ɍf�ڏI������ꍇ������܂��B</span>
			</div>
           <br clear="all">
           
           <% If G_USERTYPE = "" Then %> 
            <div id="top_reg_button">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="�������o�^���ĉ���" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="���O�C�����ĉ���" border="0">
            </a>
			</div>
            <% End If %>
            
		<%	

		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
	Else
		Response.Write "<div id=""catchcopy2"">"
		Response.Write "<div id=""in_catch"">"		
		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"" style=""padding-top:20px;"">" & rRS.Collect("CatchCopy") & "</p><br><br>"


		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">�y��W�̓����z</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write"</div>"
		
			%>
       		<div id="lissapo">
			<div><span>�]�E�T�|�[�g�Č�</span><br>
			�l�މ�Ђ̓]�E�x�����󂯂ĉ���ł��鋁�l
			</div>
			<p>�f�ڊ����F<%= sPublishLimitStr %><br>
			���ԉ{���񐔁F<%= vAccessCount %>��<br>
			�X�V���F<%= sUpdateDay %></p>
			<span>�����O�Ɍf�ڏI������ꍇ������܂��B</span>
			</div>
           <br clear="all">
           <% If G_USERTYPE = "" Then %> 
            <div id="top_reg_button">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="�������o�^���ĉ���" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="���O�C�����ĉ���" border="0">
            </a>
			</div>
            <% End If %>

		<%
		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
			
		  If G_USERTYPE = "" Then  %>
			
<div class="center">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="�������o�^���ĉ���" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="���O�C�����ĉ���" border="0">
            </a>
			</div>

		<% End If 

	End If
End Function


'******************************************************************************
'�T�@�v�F���l�[�̃L���b�`�R�s�[�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F2

'******************************************************************************
Function DspOrderCatchCopy2(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vAccessCount)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode
    Dim dbCompanyName

	Dim sOptionNo			'�傫���ʐ^�̔ԍ�
	Dim sCompanyPictureFlag	'��Ǝʐ^�t���O ["1"]�L ["0"]��
	Dim sImg1
	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic
	
	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")

		sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

        if dbOrderType = "0" Then
            dbCompanyName = rRS.Collect("CompanyName")
        else
            dbCompanyName = ""
        End if
        %>
        <div id="c_name"><%= dbCompanyName %></div>
        <%
	'******************************************************************************
	'�傫���摜 start
	'------------------------------------------------------------------------------
	dbImageLimit = rRS.Collect("ImageLimit")
	sOptionNo = ""
	sImg1 = ""
	If dbImageLimit > 0 Then
		If dbImageLimit > 1 Then
			sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
					sOptionNo = oRS.Collect("OptionNo1")
					sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & sOptionNo
				End If
			End If
		End If

		If sImg1 = "" And dbOrderType = "0" Then
			sSQL = "sp_GetDataPicture '" & dbCompanyCode & "', '1'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'�傫���摜 end
	'******************************************************************************

	'�X�V��
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'���l�[�f�ڊ��� start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">�������O�Ɍf�ڏI������ꍇ������܂��B</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'�f�ڒ� or ��f��
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'�ЊO�Č��Ȃ�DspPublicLimitDay���A�Г��Č��Ȃ�PublicLimitDay��\��
	'�ЊO�Č� OrderType = 0
	'�Г��Č� OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
			'���C�Z���X�؂�̂Ƃ���"�f�ڏI��"�ƕ\��
			sPublishLimitStr = "�f�ڏI��"
			sCautionStr = ""
		Else
			sPublishLimitStr = "�펞��W��"
		End If
	End If
	'------------------------------------------------------------------------------
	'���l�[�f�ڊ��� end
	'******************************************************************************

	'<�Г��Č��p�ʐ^>
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			End If
		End If
		Call RSClose(oRS)
	End If
	'</�Г��Č��p�ʐ^>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)


	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"">"

		Response.Write "<div class=""main_pics""><div>"
		Response.Write "<img src=""" & sImg1 & """ alt="""" id=""big_pics"">"
		Response.Write "</div></div>"

		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"">" & rRS.Collect("CatchCopy") & "</p><br>"
		Response.Write "<div>"

		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">�y��W�̓����z</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write "</div>"

		%>
        	<div id="lissapo">
			<div><span>���ڂ�����Č�</span><br>
			���̃y�[�W�����Ƃ֒��ڂ�����ł��鋁�l 
			</div>
			<p>�f�ڊ����F<%= sPublishLimitStr %><br>
			���ԉ{���񐔁F<%= vAccessCount %>��<br>
			�X�V���F<%= sUpdateDay %></p>
			<span>�����O�Ɍf�ڏI������ꍇ������܂��B</span>
			</div>
           <br clear="all">
           <% If G_USERTYPE = "" Then %> 
            <div id="top_reg_button">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="�������o�^���ĉ���" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="���O�C�����ĉ���" border="0">
            </a>
			</div>
            <% End If %> 
		<%	

		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
	Else
		Response.Write "<div id=""catchcopy2"">"
		Response.Write "<div id=""in_catch"">"		
		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"" style=""padding-top:20px;"">" & rRS.Collect("CatchCopy") & "</p><br><br>"


		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">�y��W�̓����z</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write"</div>"
		
			%>

        	<div id="lissapo">
			<div><span>���ڂ�����Č�</span><br>
			���̃y�[�W�����Ƃ֒��ڂ�����ł��鋁�l 
			</div>
			<p>�f�ڊ����F<%= sPublishLimitStr %><br>
			���ԉ{���񐔁F<%= vAccessCount %>��<br>
			�X�V���F<%= sUpdateDay %></p>
			<span>�����O�Ɍf�ڏI������ꍇ������܂��B</span>
			</div>
           
 

		<%
		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
			
		  If G_USERTYPE = "" Then  %>
			
<div class="center">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="�������o�^���ĉ���" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="���O�C�����ĉ���" border="0">
            </a>
			</div>

		<% End If 

	End If
End Function

'******************************************************************************
'�T�@�v�F���l�[�̃L���b�`�R�s�[�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F���X�T�|�[�g�Č��p

'******************************************************************************
Function DspOrderCatchCopy3(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vAccessCount)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode

	Dim sOptionNo			'�傫���ʐ^�̔ԍ�
	Dim sCompanyPictureFlag	'��Ǝʐ^�t���O ["1"]�L ["0"]��
	Dim sImg1
	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic
	Dim dbCompanyName '���X���̂̋��l�Ɏg����Ж�

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")

		sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

        if dbCompanyCode = "C0001533" Then
            dbCompanyName = rRS.Collect("CompanyName")
            %>
            <div id="c_name"><%= dbCompanyName %></div>
            <%
        else
            dbCompanyName = ""
        End if
	'******************************************************************************
	'�傫���摜 start
	'------------------------------------------------------------------------------
	dbImageLimit = rRS.Collect("ImageLimit")
	sOptionNo = ""
	sImg1 = ""
	If dbImageLimit > 0 Then
		If dbImageLimit > 1 Then
			sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
					sOptionNo = oRS.Collect("OptionNo1")
					sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & sOptionNo
				End If
			End If
		End If

		If sImg1 = "" And dbOrderType = "0" Then
			sSQL = "sp_GetDataPicture '" & dbCompanyCode & "', '1'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'�傫���摜 end
	'******************************************************************************

	'�X�V��
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'���l�[�f�ڊ��� start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">�������O�Ɍf�ڏI������ꍇ������܂��B</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'�f�ڒ� or ��f��
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'�ЊO�Č��Ȃ�DspPublicLimitDay���A�Г��Č��Ȃ�PublicLimitDay��\��
	'�ЊO�Č� OrderType = 0
	'�Г��Č� OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
			'���C�Z���X�؂�̂Ƃ���"�f�ڏI��"�ƕ\��
			sPublishLimitStr = "�f�ڏI��"
			sCautionStr = ""
		Else
			sPublishLimitStr = "�펞��W��"
		End If
	End If
	'------------------------------------------------------------------------------
	'���l�[�f�ڊ��� end
	'******************************************************************************

	'<�Г��Č��p�ʐ^>
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			End If
		End If
		Call RSClose(oRS)
	End If
	'</�Г��Č��p�ʐ^>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)


	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"">"

		Response.Write "<div class=""main_pics""><div>"
		Response.Write "<img src=""" & sImg1 & """ alt="""" id=""big_pics"">"
		Response.Write "</div></div>"

		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"">" & rRS.Collect("CatchCopy") & "</p><br>"
		Response.Write "<div>"

		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">�y��W�̓����z</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write "</div>"

		%>
			<div id="lissapo">
			<div><span>�]�E�T�|�[�g�Č�</span><br>
			�l�މ�Ђ̓]�E�x�����󂯂ĉ���ł��鋁�l
			</div>
			<p>�f�ڊ����F<%= sPublishLimitStr %><br>
			���ԉ{���񐔁F<%= vAccessCount %>��<br>
			�X�V���F<%= sUpdateDay %></p>
			<span>�����O�Ɍf�ڏI������ꍇ������܂��B</span>
			</div>
           <br clear="all">
           <% If G_USERTYPE = "" Then %> 
            <div id="top_reg_button">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="�������o�^���ĉ���" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="���O�C�����ĉ���" border="0">
            </a>
			</div>
            <% End If %>
		

		<br clear="all">
		</div>
       
    <%
	Else
		Response.Write "<div id=""catchcopy2"">"
		Response.Write "<div id=""in_catch"">"		
		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"" style=""padding-top:20px;"">" & rRS.Collect("CatchCopy") & "</p><br><br>"


		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">�y��W�̓����z</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write"</div>"
		
			%>

        	<div id="lissapo">
			<div><span>�]�E�T�|�[�g�Č�</span><br>
			�l�މ�Ђ̓]�E�x�����󂯂ĉ���ł��鋁�l
			</div>
			<p>�f�ڊ����F<%= sPublishLimitStr %><br>
			���ԉ{���񐔁F<%= vAccessCount %>��<br>
			�X�V���F<%= sUpdateDay %></p>
			<span>�����O�Ɍf�ڏI������ꍇ������܂��B</span>
			</div>

		<%
		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
			
		  If G_USERTYPE = "" Then  %>
			
<div class="center">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="�������o�^���ĉ���" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="���O�C�����ĉ���" border="0">
            </a>
			</div>
			
		<% End If 

	End If
End Function

'******************************************************************************
'�T�@�v�F���l�[�̃L���b�`�R�s�[�������o��(�ߋ����l)
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F���X�T�|�[�g�Č��p

'******************************************************************************
Function DspOrderCatchCopy_OldPlan(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vAccessCount,ByVal YearlyIncomeMin,ByVal MonthlyIncomeMin,ByVal DailyIncomeMin,ByVal HourlyIncomeMin)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode

	Dim sOptionNo			'�傫���ʐ^�̔ԍ�
	Dim sCompanyPictureFlag	'��Ǝʐ^�t���O ["1"]�L ["0"]��
	Dim sImg1
	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic

	Dim JobTypeBigCode
	Dim JobTypeCode
	Dim WorkingTypeCode1
	Dim WorkingTypeCode2
	Dim WorkingTypeCode3
	Dim PrefectureCode
	
	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")

		sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'�傫���摜 start
	'------------------------------------------------------------------------------
	dbImageLimit = rRS.Collect("ImageLimit")
	sOptionNo = ""
	sImg1 = ""
	If dbImageLimit > 0 Then
		If dbImageLimit > 1 Then
			sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
					sOptionNo = oRS.Collect("OptionNo1")
					sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & sOptionNo
				End If
			End If
		End If

		If sImg1 = "" And dbOrderType = "0" Then
			sSQL = "sp_GetDataPicture '" & dbCompanyCode & "', '1'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'�傫���摜 end
	'******************************************************************************

	'�X�V��
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'���l�[�f�ڊ��� start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">�������O�Ɍf�ڏI������ꍇ������܂��B</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'�f�ڒ� or ��f��
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'�ЊO�Č��Ȃ�DspPublicLimitDay���A�Г��Č��Ȃ�PublicLimitDay��\��
	'�ЊO�Č� OrderType = 0
	'�Г��Č� OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
			'���C�Z���X�؂�̂Ƃ���"�f�ڏI��"�ƕ\��
			sPublishLimitStr = "�f�ڏI��"
			sCautionStr = ""
		Else
			sPublishLimitStr = "�펞��W��"
		End If
	End If
	'------------------------------------------------------------------------------
	'���l�[�f�ڊ��� end
	'******************************************************************************

	'<�Г��Č��p�ʐ^>
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			End If
		End If
		Call RSClose(oRS)
	End If
	'</�Г��Č��p�ʐ^>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)


	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"">"

		Response.Write "<div class=""main_pics""><div>"
		Response.Write "<img src=""" & sImg1 & """ alt="""" id=""big_pics"">"
		Response.Write "</div></div>"

		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"">" & rRS.Collect("CatchCopy") & "</p><br>"
		Response.Write "<div>"

		If sImgSpeciality <> "" Then

			
			%>			

<div class="right">
	<img src="/img/order/oubo_end.png" class="spSmart">
</div>
 <br clear="both">  

		<div class="center" style="margin-top:25px;">
                    <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?sdf=1&amp;sjtbig1=<%= JobTypeBigCode %>&amp;sjt1=<%= JobTypeCode %>&amp;swt1=<%= WorkingTypeCode1 %>&amp;swt2=<%= WorkingTypeCode2 %>&amp;swt3=<%= WorkingTypeCode3 %>&amp;spc=<%= PrefectureCode %>&amp;syimin=<%= YearlyIncomeMin %>&amp;smimin=<%= MonthlyIncomeMin %>&amp;sdimin=<%= DailyIncomeMin %>&amp;shimin=<%= HourlyIncomeMin %>">
           		<img src="<%= HTTP_NAVI_CURRENTURL %>img/order/top_reg_button03.png" alt="����o�^�����ď����ɋ߂����l�։���" border="0">
            </a>
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_list.asp?sdf=1&amp;sjtbig1=<%= JobTypeBigCode %>&amp;sjt1=<%= JobTypeCode %>&amp;swt1=<%= WorkingTypeCode1 %>&amp;swt2=<%= WorkingTypeCode2 %>&amp;swt3=<%= WorkingTypeCode3 %>&amp;spc=<%= PrefectureCode %>&amp;syimin=<%= YearlyIncomeMin %>&amp;smimin=<%= MonthlyIncomeMin %>&amp;sdimin=<%= DailyIncomeMin %>&amp;shimin=<%= HourlyIncomeMin %>">
           		<img src="<%= HTTP_NAVI_CURRENTURL %>img/order/top_login_button03.png" alt="���O�C�����ď����ɋ߂����l�։���" border="0">
            </a>
      
            
		</div>
  
<% 
			
		End If

		Response.Write "</div>"


		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
	Else
		Response.Write "<div id=""catchcopy3"" class=""left"">"
	
		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"" style=""padding-top:20px;"">" & rRS.Collect("CatchCopy") & "</p><br><br>"



		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
			
%>			

<div class="right">
	<img src="/img/order/oubo_end.png" class="spSmart">
</div>
 <br clear="both">   
<%    		
		If G_USERTYPE = "" Then
			'�ٗp�`�ԁA�Ζ��n�A�E�팟��
			sSQL = "select CJT.jobtypecode, BJT.Bigclasscode from C_JobType AS CJT INNER JOIN B_JobType AS BJT ON CJT.JobTypeCode = BJT.AllConnectCode where CJT.id = '1' and CJT.OrderCode = '" & dbOrderCode & "';"
			flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("Bigclasscode")) <> "" Then
					JobTypeBigCode = oRS.Collect("Bigclasscode")
				End If
				If ChkStr(oRS.Collect("jobtypecode")) <> "" Then
					JobTypeCode = oRS.Collect("jobtypecode")
				End If
			End If
			Call RSClose(oRS)

			sSQL = "select prefecturecode from c_workingplace where ordercode = '" & dbOrderCode & "';"
			flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
			PrefectureCode  = ""
			Do While GetRSState(oRS) = True
				If ChkStr(oRS.Collect("prefecturecode")) <> "" Then
					PrefectureCode = PrefectureCode & oRS.Collect("prefecturecode") & ","
				End If
				oRS.MoveNext
			Loop
			Call RSClose(oRS)
			PrefectureCode = Left(PrefectureCode, Len(PrefectureCode) -1)

			sSQL = "select CWT1.workingtypecode as workingtypecode1,CWT2.workingtypecode as workingtypecode2,CWT3.workingtypecode as workingtypecode3 from c_workingtype AS CWT1 "
			sSQL = sSQL & " left join c_workingtype AS CWT2 on CWT1.ordercode = '" & dbOrderCode & "' and CWT2.id = 2"
			sSQL = sSQL & " left join c_workingtype AS CWT3 on CWT2.ordercode = '" & dbOrderCode & "' and CWT3.id = 3"
			sSQL = sSQL & " where CWT3.ordercode = '" & dbOrderCode & "' and CWT1.id = 1;"
			flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("workingtypecode1")) <> "" Then
					WorkingTypeCode1 = oRS.Collect("workingtypecode1")
				End If
				If ChkStr(oRS.Collect("workingtypecode2")) <> "" Then
					WorkingTypeCode2 = oRS.Collect("workingtypecode2")
				End If
				If ChkStr(oRS.Collect("workingtypecode3")) <> "" Then
					WorkingTypeCode3 = oRS.Collect("workingtypecode3")
				End If
			End If
			Call RSClose(oRS)
%>

			
		<div class="center">
		<a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?sdf=1&amp;sjtbig1=<%= JobTypeBigCode %>&amp;sjt1=<%= JobTypeCode %>&amp;swt1=<%= WorkingTypeCode1 %>&amp;swt2=<%= WorkingTypeCode2 %>&amp;swt3=<%= WorkingTypeCode3 %>&amp;spc=<%= PrefectureCode %>&amp;syimin=<%= YearlyIncomeMin %>&amp;smimin=<%= MonthlyIncomeMin %>&amp;sdimin=<%= DailyIncomeMin %>&amp;shimin=<%= HourlyIncomeMin %>">
		<img src="<%= HTTP_NAVI_CURRENTURL %>img/order/top_reg_button03.png" alt="����o�^�����ď����ɋ߂����l�։���" border="0">
		</a>
		<a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_list.asp?sdf=1&amp;sjtbig1=<%= JobTypeBigCode %>&amp;sjt1=<%= JobTypeCode %>&amp;swt1=<%= WorkingTypeCode1 %>&amp;swt2=<%= WorkingTypeCode2 %>&amp;swt3=<%= WorkingTypeCode3 %>&amp;spc=<%= PrefectureCode %>&amp;syimin=<%= YearlyIncomeMin %>&amp;smimin=<%= MonthlyIncomeMin %>&amp;sdimin=<%= DailyIncomeMin %>&amp;shimin=<%= HourlyIncomeMin %>">
		<img src="<%= HTTP_NAVI_CURRENTURL %>img/order/top_login_button03.png" alt="���O�C�����ď����ɋ߂����l�։���" border="0">
		</a>
		</div>
<% End If 

	End If
	

End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̃t���[�o�q���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F
'���@���F2007/02/11 LIS K.Kokubo �쐬
'******************************************************************************
Function DspOrderFreePR(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sPRTitle1			'�o�q�^�C�g��1
	Dim sPRTitle2			'�o�q�^�C�g��2
	Dim sPRTitle3			'�o�q�^�C�g��3
	Dim sPRContents1		'�o�q��1
	Dim sPRContents2		'�o�q��2
	Dim sPRContents3		'�o�q��3
	Dim flgPR				'�o�q�L���t���O [True]�L [False]��

	Dim dbOrderCode

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")

	'******************************************************************************
	'PR start
	'------------------------------------------------------------------------------
	flgPR = False
	sPRTitle1 = ChkStr(rRS.Collect("PRTitle1"))
	sPRTitle2 = ChkStr(rRS.Collect("PRTitle2"))
	sPRTitle3 = ChkStr(rRS.Collect("PRTitle3"))
	sPRContents1 = Replace(ChkStr(rRS.Collect("PRContents1")), vbCrLf, "<br>")
	sPRContents1 = Replace(sPRContents1, vbCr, "<br>")
	sPRContents1 = Replace(sPRContents1, vbLf, "<br>")
	sPRContents2 = Replace(ChkStr(rRS.Collect("PRContents2")), vbCrLf, "<br>")
	sPRContents2 = Replace(sPRContents2, vbCr, "<br>")
	sPRContents2 = Replace(sPRContents2, vbLf, "<br>")
	sPRContents3 = Replace(ChkStr(rRS.Collect("PRContents3")), vbCrLf, "<br>")
	sPRContents3 = Replace(sPRContents3, vbCr, "<br>")
	sPRContents3 = Replace(sPRContents3, vbLf, "<br>")

	If sPRTitle1 & sPRTitle2 & sPRTitle3 & sPRContents1 & sPRContents2 & sPRContents3 <> "" Then flgPR = True
	'------------------------------------------------------------------------------
	'PR end
	'******************************************************************************

	If flgPR = True Then
	%>
    	<img src="/img/order/tab_detail_pr.png" class="tab_img">
		<table class="detail_table">
        <tbody>
        <tr>
        <th class="fast_th"></th>
        <td>
        <%
		Response.Write "<div>"
		If sPRTitle1 <> "" Or sPRContents1 <> "" Then
			Response.Write "<h4>" & sPRTitle1 & "</h4>"
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "<p class=""m0"">" & sPRContents1 & "</p>"
		End If

		If sPRTitle2 <> "" Or sPRContents2 <> "" Then
			Response.Write "<h4>" & sPRTitle2 & "</h4>"
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "<p class=""m0"">" & sPRContents2 & "</p>"
		End If

		If sPRTitle3 <> "" Or sPRContents3 <> "" Then
			Response.Write "<h4>" & sPRTitle3 & "</h4>"
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "<p class=""m0"">" & sPRContents3 & "</p>"
		End If
		Response.Write "</div>"
		%>
        </td>
        </tr>
        </tbody>
        </table>
        
        <div class="to_top"><a class="stext_middle" href="#pagetop">���y�[�WTOP��</a></div>  
        
     <%   
	End If
End Function

'******************************************************************************
'�T�@�v�F���l��Ɖ摜�ꗗ�\���g�s�l�k�\��
'���@���FrDB			�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvCategoryCode	�F�J�e�S���R�[�h
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F
'���@���F2006/12/27 LIS K.Kokubo �쐬
'�@�@�@�F2008/01/28 LIS K.Kokubo ���C�Z���X�ύX�ɂ��Ή�
'�@�@�@�F2010/05/06 LIS K.Kokubo �Г��Č��p�ʐ^
'******************************************************************************
Function DspOrderPictureNow(ByRef rDB, ByRef rRS, ByVal vCategoryCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode
	Dim dbCompanyCode
	Dim dbOrderType
	Dim dbImageLimit

	Dim sURL
	Dim sImg1,sImg2,sImg3,sCap1,sCap2,sCap3

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbCompanyCode = rRS.Collect("CompanyCode")
	dbOrderType = rRS.Collect("OrderType")
	dbImageLimit = rRS.Collect("ImageLimit")

	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo2")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo2")
				sCap1 = ChkStr(oRS.Collect("Caption2"))
			End If
			If ChkStr(oRS.Collect("PicNo3")) <> "" Then
				sImg2 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo3")
				sCap2 = ChkStr(oRS.Collect("Caption3"))
			End If
			If ChkStr(oRS.Collect("PicNo4")) <> "" Then
				sImg3 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo4")
				sCap3 = ChkStr(oRS.Collect("Caption4"))
			End If
		End If
		Call RSClose(oRS)
	ElseIf dbImageLimit > 1 Then
		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', '" & vCategoryCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("OptionNo2")) <> "" Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo2")
				sCap1 = ChkStr(oRS.Collect("Caption2"))
			End If
			If ChkStr(oRS.Collect("OptionNo3")) <> "" Then
				sImg2 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo3")
				sCap2 = ChkStr(oRS.Collect("Caption3"))
			End If
			If ChkStr(oRS.Collect("OptionNo4")) <> "" Then
				sImg3 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo4")
				sCap3 = ChkStr(oRS.Collect("Caption4"))
			End If
		End If
	End If

	If sImg1 & sImg2 & sImg3 <> "" Then
		Response.Write "<div id=""sub_pics"">"
		Response.Write "<div class=""auto"">"

		If sImg1 <> "" Then
			Response.Write "<div class=""sub_waku"">"
			Response.Write "<div class=""sub_pics sub_pics1""><div><img src=""" & sImg1 & """ alt=""" & sCap1 & """></div></div>"
			Response.Write "<p class=""m0"" align=""left"" style=""width:213px; font-size:10px;"">" & sCap1 & "</p>"
			Response.Write "</div>"
		End If

		If sImg2 <> "" Then
			Response.Write "<div class=""sub_waku"">"
			Response.Write "<div class=""sub_pics sub_pics2""><div><img src=""" & sImg2 & """ alt=""" & sCap2 & """></div></div>"
			Response.Write "<p class=""m0"" align=""left"" style=""width:213px; font-size:10px;"">" & sCap2 & "</p>"
			Response.Write "</div>"
		End If

		If sImg3 <> "" Then
			Response.Write "<div class=""sub_waku"">"
			Response.Write "<div class=""sub_pics sub_pics3""><div><img src=""" & sImg3 & """ alt=""" & sCap3 & """></div></div>"
			Response.Write "<p class=""m0"" align=""left"" style=""width:213px; font-size:10px;"">" & sCap3 & "</p>"
			Response.Write "</div>"
		End If


		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
		Response.Write "</div>"
	End If
End Function

'******************************************************************************
'�T�@�v�F���l�[�̗̍p�̔w�i���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F
'���@���F2007/05/13 LIS K.Kokubo �쐬
'******************************************************************************
Function DspOrderBackGround(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderBackGround	'�̗p�̔w�i

	DspOrderBackGround = False

	If GetRSState(rRS) = False Then Exit Function

	'�̗p�̔w�i�擾
	dbOrderBackGround = Replace(ChkStr(rRS.Collect("OrderBackGround")), vbCrLf, "<br>")

	'�̗p�̔w�i�o��
	If dbOrderBackGround <> "" Then
	%>
    <img src="/img/order/tab_detail_bb.png" class="tab_img">
	<table class="detail_table">
    <tbody>
    <tr>
    <th class="fast_th"></th>
    <td><p class="m0"><%= dbOrderBackGround %></p></td>
    </tr>
    </tbody>
    </table>
    
	<div class="to_top"><a class="stext_middle" href="#pagetop">���y�[�WTOP��</a></div>  
		
        <%
        DspOrderBackGround = True
	End If


End Function

'******************************************************************************
'�T�@�v�F���l�[�̋Ɩ����e���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F
'���@���F2007/02/11 LIS K.Kokubo �쐬
'******************************************************************************
Function DspBusiness(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode			'���R�[�h
	Dim sCompanyCode		'��ƃR�[�h
	Dim sPlanType			'���l�[���C�Z���X�v�������
	Dim sBizName1			'�d����������1
	Dim sBizName2			'�d����������2
	Dim sBizName3			'�d����������3
	Dim sBizName4			'�d����������4
	Dim sBizPercentage1		'�d������1
	Dim sBizPercentage2		'�d������2
	Dim sBizPercentage3		'�d������3
	Dim sBizPercentage4		'�d������4
	Dim sBiz				'�d������HTML
	Dim sBusinessDetail		'�S���Ɩ�
	Dim sClearSolid
	Dim flgBusiness
	Dim flgLine				'�������t���O

	If GetRSState(rRS) = False Then Exit Function

	flgBusiness = False
	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	sCompanyCode = rRS.Collect("CompanyCode")
	sPlanType = rRS.Collect("PlanTypeName")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'******************************************************************************
	'�d���̊��� start
	'------------------------------------------------------------------------------
	sBiz = ""
	sBizName1 = ""
	sBizName2 = ""
	sBizName3 = ""
	sBizName4 = ""
	sBizPercentage1 = ""
	sBizPercentage2 = ""
	sBizPercentage3 = ""
	sBizPercentage4 = ""

	sBizName1 = ChkStr(rRS.Collect("BizName1"))
	sBizName2 = ChkStr(rRS.Collect("BizName2"))
	sBizName3 = ChkStr(rRS.Collect("BizName3"))
	sBizName4 = ChkStr(rRS.Collect("BizName4"))
	sBizPercentage1 = ChkStr(rRS.Collect("BizPercentage1"))
	sBizPercentage2 = ChkStr(rRS.Collect("BizPercentage2"))
	sBizPercentage3 = ChkStr(rRS.Collect("BizPercentage3"))
	sBizPercentage4 = ChkStr(rRS.Collect("BizPercentage4"))

	If sBizName1 & sBizName2 & sBizName3 & sBizName4 <> "" Then
		If sBizName1 <> "" And sBizPercentage1 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName1 & "</td><td class=""biz2"">" & sBizPercentage1 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage1) * 3 & """ height=""20""></td></tr>"
		If sBizName2 <> "" And sBizPercentage2 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName2 & "</td><td class=""biz2"">" & sBizPercentage2 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage2) * 3 & """ height=""20""></td></tr>"
		If sBizName3 <> "" And sBizPercentage3 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName3 & "</td><td class=""biz2"">" & sBizPercentage3 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage3) * 3 & """ height=""20""></td></tr>"
		If sBizName4 <> "" And sBizPercentage4 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName4 & "</td><td class=""biz2"">" & sBizPercentage4 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage4) * 3 & """ height=""20""></td></tr>"
		sBiz = "<table>" & sBiz & "</table>"
		flgBusiness = True
	End If
	'------------------------------------------------------------------------------
	'�d���̊��� end
	'******************************************************************************

	'******************************************************************************
	'�S���Ɩ� start
	'------------------------------------------------------------------------------
	sBusinessDetail = Replace(ChkStr(rRS.Collect("BusinessDetail")), vbCrLf, "<br>")
	sBusinessDetail = Replace(sBusinessDetail, vbCr, "<br>")
	sBusinessDetail = Replace(sBusinessDetail, vbLf, "<br>")
	If sBusinessDetail <> "" Then flgBusiness = True
	'------------------------------------------------------------------------------
	'�S���Ɩ� end
	'******************************************************************************

	flgLine = False
	If flgBusiness = True Then
			If sBusinessDetail <> "" Then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
	%>
    <img src="/img/order/tab_detail_job.png" class="tab_img">
   	<table class="detail_table">
    <tbody>
    <tr>
    <th rowspan="2" class="fast_th"></th>
    <td><h4>�S���Ɩ�</h4>
    <p class=""m0""><%= sBusinessDetail %></p></td>
    </tr>
    <tr>
    <td><%
	 If (sPlanType = "platinum" Or sPlanType = "gold" Or sPlanType = "old") And sBiz <> "" Then


			Response.Write "<h4>�d���̊���</h4>"
			Response.Write "<div class=""value1"">"
			Response.Write "<table border=""0"">"
			Response.Write "<tbody>"
			Response.Write "<tr>"
			Response.Write "<td>"
			Response.Write "<script type=""text/javascript"" language=""javascript"">"
			Response.Write "viewWorkAvg(" & sBizPercentage1 & ", " & sBizPercentage2 & ", " & sBizPercentage3 & ", " & sBizPercentage4 & ")"
			Response.Write "</script>"
			Response.Write "</td>"
			Response.Write "<td style=""padding-left:5px; vertical-align:middle;"">"
			Response.Write "<table border=""0"">"
			Response.Write "<tbody>"
			If sBizName1 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#ff9999; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage1 & "%</td><td>" & sBizName1 & "</td></tr>"
			If sBizName2 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#9999ff; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage2 & "%</td><td>" & sBizName2 & "</td></tr>"
			If sBizName3 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#99ff99; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage3 & "%</td><td>" & sBizName3 & "</td></tr>"
			If sBizName4 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#ffff99; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage4 & "%</td><td>" & sBizName4 & "</td></tr>"
			Response.Write "</tbody>"
			Response.Write "</table>"
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "</tbody>"
			Response.Write "</table>"
			Response.Write "</div>"
		End If
		 %>
         </td>
    </tr>
    </tbody>
    </table>
    <div class="to_top"><a class="stext_middle" href="#pagetop">���y�[�WTOP��</a></div>  
    
    <%
  		End If


	End If
End Function

'******************************************************************************
'�T�@�v�F���l�[�̋Ζ��������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/include/func_order.asp
'���@�l�F
'���@���F2007/02/11 �쐬
'�@�@�@�F2008/10/22 LIS K.Kokubo �Ζ��n�������Ή�
'�@�@�@�F2009/04/16 LIS K.Kokubo ���[���ۋ����C�Z���X�̏ꍇ�͋Ζ��n�̕\������ʂ̋��l�L���ł��s��S�܂ł����\�������Ȃ�
'�@�@�@�F2009/04/22 LIS K.Kokubo �Љ��̋Ζ��`��(TTP�p)�Ή�
'�@�@�@�F2009/11/02 LIS K.Kokubo �r�n�g�n,�e�b�̋Ζ��n�\���Ή�
'******************************************************************************
Function DspCondition(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	'<�ϐ��錾>
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim oRS3
	Dim flgQE
	Dim sError

	Dim dbOrderCode			'���R�[�h
	Dim dbOrderType			'���l�[���
	Dim dbCompanyKbn		'��Ƌ敪
	Dim dbJobTypeDetail		'�E��ڍ�
	Dim dbYearlyIncomeMin	'�N������
	Dim dbYearlyIncomeMax	'�N�����
	Dim dbMonthlyIncomeMin	'��������
	Dim dbMonthlyIncomeMax	'�������
	Dim dbDailyIncomeMin	'��������
	Dim dbDailyIncomeMax	'�������
	Dim dbHourlyIncomeMin	'��������
	Dim dbHourlyIncomeMax	'�������
	Dim dbPercentagePay		'������
	Dim dbSalaryRemark		'���^���l
	Dim dbTrafficFeeType	'
	Dim dbTrafficFeeMonth	'��ʔ�^�P����
	Dim dbAfterWorkingTypeCode'�Љ��̋Ζ��`��
	Dim dbWorkStartDay		'�A�ƊJ�n��
	Dim dbWorkEndDay		'�A�ƏI����
	Dim dbWorkTimeRemark	'�A�Ǝ��Ԕ��l
	Dim dbWeeklyHolidayType	'�T�x
	Dim dbHolidayRemark		'�x�����l
	Dim dbHumanNumber		'��W�l��
	Dim dbWorkingPlaceSeq	'�Ζ��n�ԍ�
	Dim dbWorkingPlacePrefectureName'�Ζ��n�s���{����
	Dim dbWorkingPlaceCity	'�Ζ��n�s��S
	Dim dbWorkingPlaceAddressAll'�Ζ��n�Z���S��
	Dim dbWorkingPlaceSection'�Ζ��n����
	Dim dbWorkingPlaceTelephoneNumber'�Ζ��nTEL
	Dim dbMapFlag			'�n�}�L���t���O
	Dim dbTransfer			'�]��
	Dim dbPlanTypeName		'�i�r���C�Z���X�̎��
	Dim dbTTPOrderFlag		'�Љ�\��h���Č��t���O

	Dim sHTML
	Dim sWorkingType		'�Ζ��`��
	Dim sJobType			'�E��
	Dim sSalary				'���^
	Dim sYearlyIncome		'�N��
	Dim sMonthlyIncome		'����
	Dim sDailyIncome		'����
	Dim sHourlyIncome		'����
	Dim sTrafficFee			'��ʔ�
	Dim sAfterWorkingType	'�Љ��̋Ζ��`��
	Dim sWorkRange			'�A�Ɗ���
	Dim sWorkUpdate			'�A�Ɗ��Ԃ̍X�V�L��
	Dim sWorkingTime		'�A�Ǝ���
	Dim sMAP				'�n�}���
	Dim sWorkingPlace		'�A�Əꏊ
	Dim sNearbyStation		'�Ŋ�w
	Dim sNearbyRailway		'����
	Dim sNearbyStationBlock	'�Ŋ�w,�����u���b�N
	Dim iMaxRow
	Dim sDisplay
	Dim flgDspWorkingType
	Dim flgDspJobType
	Dim flgDspSalary
	Dim flgDspTime
	Dim flgDspHoliday
	Dim flgDspHumanNumber
	Dim flgDspWorkingPlace
	Dim flgLine
	Dim flgSOHOFC
	'</�ϐ��錾>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyKbn = rRS.Collect("CompanyKbn")
	dbPlanTypeName = rRS.Collect("PlanTypeName")
	dbTTPOrderFlag = rRS.Collect("TTPOrderFlag")
	
		'<�Ζ��`��>
	flgDspWorkingType = False
	dbAfterWorkingTypeCode = ChkStr(rRS.Collect("AfterWorkingTypeCode"))
	dbWorkStartDay = ChkStr(rRS.Collect("WorkStartDay"))
	dbWorkEndDay = ChkStr(rRS.Collect("WorkEndDay"))

	'�Ζ��`��
	sWorkingType = GetWorkingType(rDB, rRS)
	flgSOHOFC = False
	If IsRE(sWorkingType,"((SOHO)|(FC))",True) = True Then flgSOHOFC = True

	'�Љ��̋Ζ��`��
	sAfterWorkingType = ""
	If dbAfterWorkingTypeCode <> "" Then
		sAfterWorkingType = "���Љ��̋Ζ��`��&nbsp;���&nbsp;" & GetDetail("WorkingType", dbAfterWorkingTypeCode)
	End If

	'�A�Ɗ���
	sWorkRange = ""
	If dbWorkStartDay & dbWorkEndDay <> "" Then
		If dbWorkStartDay <> "" Then sWorkRange = sWorkRange & GetDateStr(dbWorkStartDay, "/")
		If sWorkRange <> "" Then sWorkRange = sWorkRange & "�`"
		If dbWorkEndDay <> "" Then sWorkRange = sWorkRange & GetDateStr(dbWorkEndDay, "/")
	End If

	If dbOrderType = "1" Then
		If rRS.Collect("WorkUpdateFlag") = "1" Then
			sWorkUpdate = "�L"
		Else
			sWorkUpdate = "��"
		End If
		sWorkRange = sWorkRange & "(�X�V" & sWorkUpdate & ")"
	End If

	If sWorkingType & sAfterWorkingType & sWorkRange <> "" Then flgDspWorkingType = True
	'</�Ζ��`��>

	'<�E��>
	flgDspJobType = False
	sJobType = GetJobType(rDB, rRS)
	dbJobTypeDetail = rRS.Collect("JobTypeDetail")
	If sJobType & dbJobTypeDetail <> "" Then flgDspJobType = True
	'</�E��>

	'<���^>
	flgDspSalary = False
	dbYearlyIncomeMin = ChkStr(rRS.Collect("YearlyIncomeMin"))
	dbYearlyIncomeMax = ChkStr(rRS.Collect("YearlyIncomeMax"))
	If dbYearlyIncomeMin = "0" Then dbYearlyIncomeMin = ""
	If dbYearlyIncomeMax = "0" Then dbYearlyIncomeMax = ""
	If dbYearlyIncomeMin <> "" Then dbYearlyIncomeMin = GetJapaneseYen(dbYearlyIncomeMin)
	If dbYearlyIncomeMax <> "" Then dbYearlyIncomeMax = GetJapaneseYen(dbYearlyIncomeMax)
	If dbYearlyIncomeMin & dbYearlyIncomeMax <> "" Then
		If dbYearlyIncomeMin <> "" Then sYearlyIncome = sYearlyIncome & dbYearlyIncomeMin
		sYearlyIncome = sYearlyIncome & "&nbsp;�`&nbsp;"
		If dbYearlyIncomeMax <> "" Then sYearlyIncome = sYearlyIncome & dbYearlyIncomeMax
	End If

	dbMonthlyIncomeMin = ChkStr(rRS.Collect("MonthlyIncomeMin"))
	dbMonthlyIncomeMax = ChkStr(rRS.Collect("MonthlyIncomeMax"))
	If dbMonthlyIncomeMin = "0" Then dbMonthlyIncomeMin = ""
	If dbMonthlyIncomeMax = "0" Then dbMonthlyIncomeMax = ""
	If dbMonthlyIncomeMin <> "" Then dbMonthlyIncomeMin = GetJapaneseYen(dbMonthlyIncomeMin)
	If dbMonthlyIncomeMax <> "" Then dbMonthlyIncomeMax = GetJapaneseYen(dbMonthlyIncomeMax)
	If dbMonthlyIncomeMin & dbMonthlyIncomeMax <> "" Then
		If dbMonthlyIncomeMin <> "" Then sMonthlyIncome = sMonthlyIncome & dbMonthlyIncomeMin
		sMonthlyIncome = sMonthlyIncome & "&nbsp;�`&nbsp;"
		If dbMonthlyIncomeMax <> "" Then sMonthlyIncome = sMonthlyIncome & dbMonthlyIncomeMax
	End If

	dbDailyIncomeMin = ChkStr(rRS.Collect("DailyIncomeMin"))
	dbDailyIncomeMax = ChkStr(rRS.Collect("DailyIncomeMax"))
	If dbDailyIncomeMin = "0" Then dbDailyIncomeMin = ""
	If dbDailyIncomeMax = "0" Then dbDailyIncomeMax = ""
	If dbDailyIncomeMin <> "" Then dbDailyIncomeMin = GetJapaneseYen(dbDailyIncomeMin)
	If dbDailyIncomeMax <> "" Then dbDailyIncomeMax = GetJapaneseYen(dbDailyIncomeMax)
	If dbDailyIncomeMin & dbDailyIncomeMax <> "" Then
		If dbDailyIncomeMin <> "" Then sDailyIncome = sDailyIncome & dbDailyIncomeMin
		sDailyIncome = sDailyIncome & "&nbsp;�`&nbsp;"
		If dbDailyIncomeMax <> "" Then sDailyIncome = sDailyIncome & dbDailyIncomeMax
	End If

	dbHourlyIncomeMin = ChkStr(rRS.Collect("HourlyIncomeMin"))
	dbHourlyIncomeMax = ChkStr(rRS.Collect("HourlyIncomeMax"))
	If dbHourlyIncomeMin = "0" Then dbHourlyIncomeMin = ""
	If dbHourlyIncomeMax = "0" Then dbHourlyIncomeMax = ""
	If dbHourlyIncomeMin <> "" Then dbHourlyIncomeMin = GetJapaneseYen(dbHourlyIncomeMin)
	If dbHourlyIncomeMax <> "" Then dbHourlyIncomeMax = GetJapaneseYen(dbHourlyIncomeMax)
	If dbHourlyIncomeMin & dbHourlyIncomeMax <> "" Then
		If dbHourlyIncomeMin <> "" Then sHourlyIncome = sHourlyIncome & dbHourlyIncomeMin
		sHourlyIncome = sHourlyIncome & "&nbsp;�`&nbsp;"
		If dbHourlyIncomeMax <> "" Then sHourlyIncome = sHourlyIncome & dbHourlyIncomeMax
	End If

	dbPercentagePay = ChkStr(rRS.Collect("PercentagePayFlag"))
	dbSalaryRemark = Replace(ChkStr(rRS.Collect("IncomeRemark")), vbCrLf, "<br>")
	dbSalaryRemark = Replace(dbSalaryRemark, vbCr, "<br>")
	dbSalaryRemark = Replace(dbSalaryRemark, vbLf, "<br>")
	sTrafficFee = ""
	dbTrafficFeeType = ChkStr(rRS.Collect("TrafficFeeType"))
	dbTrafficFeeMonth = ChkStr(rRS.Collect("MonthTrafficFee"))

	'������
	If dbPercentagePay <> "" Then
		If dbPercentagePay = "1" Then
			dbPercentagePay = "����"
		ElseIf dbPercentagePay = "0" Then
			dbPercentagePay = "�Ȃ�"
		End If
	End If

	'��ʔ�
	If ChkStr(rRS.Collect("NaviTrafficPayFlag")) = "1" Then 
		sTrafficFee = "��ʔ�x������" & dbTrafficFeeType
		If IsNumber(dbTrafficFeeMonth, 0, False) = True Then
			sTrafficFee = sTrafficFee & "(" & FormatCanma(dbTrafficFeeMonth) & "�~�^��)"
		End If
	End If

	If sYearlyIncome & sMonthlyIncome & sDailyIncome & sHourlyIncome & dbPercentagePay & sTrafficFee & dbSalaryRemark <> "" Then flgDspSalary = True
	'</���^>

	'<����>
	flgDspTime = False
	sWorkingTime = GetWorkingTime(rDB, rRS)
	dbWorkTimeRemark = ChkStr(rRS.Collect("WorkTimeRemark"))
	dbWorkTimeRemark = Replace(ChkStr(rRS.Collect("WorkTimeRemark")), vbCrLf, "<br>")
	dbWorkTimeRemark = Replace(dbWorkTimeRemark, vbCr, "<br>")
	dbWorkTimeRemark = Replace(dbWorkTimeRemark, vbLf, "<br>")

	If sWorkingTime & dbWorkTimeRemark <> "" Then flgDspTime = True
	'</����>

	'<�x��>
	flgDspHoliday = False
	dbWeeklyHolidayType = ChkStr(rRS.Collect("WeeklyHolidayTypeName"))
	dbHolidayRemark = ChkStr(rRS.Collect("HolidayRemark"))
	dbHolidayRemark = Replace(ChkStr(rRS.Collect("HolidayRemark")), vbCrLf, "<br>")
	dbHolidayRemark = Replace(dbHolidayRemark, vbCr, "<br>")
	dbHolidayRemark = Replace(dbHolidayRemark, vbLf, "<br>")

	If dbWeeklyHolidayType & dbHolidayRemark <> "" Then flgDspHoliday = True
	'</�x��>

	'<��W�l��>
	flgDspHumanNumber = False
	dbHumanNumber = ChkStr(rRS.Collect("HumanNumber"))

	If dbHumanNumber <> "" Then
		dbHumanNumber = dbHumanNumber & "�l"
	End If

	If dbHumanNumber <> "" Then flgDspHumanNumber = True
	'</��W�l��>

	'<�Ζ��n>
	flgDspWorkingPlace = False

	iMaxRow = 0
	sWorkingPlace = ""
	sNearbyStationBlock = ""
	sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing
		iMaxRow = oRS.RecordCount
		'<�Ŋ�w>
		sSQL = "EXEC up_LstC_NearbyStation '" & dbOrderCode & "', '';"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then Set oRS2.ActiveConnection = Nothing
		'</�Ŋ�w>
		'<�Ŋ񉈐�>
		sSQL = "EXEC up_LstC_NearbyRailwayLine '" & rRS.Collect("OrderCode") & "','','';"
		flgQE = QUERYEXE(rDB, oRS3, sSQL, sError)
		If GetRSState(oRS3) = True Then Set oRS3.ActiveConnection = Nothing
		'</�Ŋ񉈐�>
	End If
	Do While GetRSState(oRS) = True
		dbWorkingPlaceSeq = ChkStr(oRS.Collect("WorkingPlaceSeq"))
		dbWorkingPlacePrefectureName = ChkStr(oRS.Collect("WorkingPlacePrefectureName"))
		dbWorkingPlaceCity = ChkStr(oRS.Collect("WorkingPlaceCity"))
		dbWorkingPlaceAddressAll = ChkStr(oRS.Collect("WorkingPlaceAddressAll"))
		dbWorkingPlaceSection = ChkStr(oRS.Collect("WorkingPlaceSection"))
		dbWorkingPlaceTelephoneNumber = ChkStr(oRS.Collect("WorkingPlaceTelephoneNumber"))
		dbMapFlag = ChkStr(oRS.Collect("MapFlag"))

		If sWorkingPlace <> "" And flgSOHOFC = True Then sWorkingPlace = sWorkingPlace & "�A"

		'<�Ζ��n>
		sWorkingPlace = sWorkingPlace & "<div"
		If flgSOHOFC = True Then sWorkingPlace = sWorkingPlace & " style=""display:inline;"""
		sWorkingPlace = sWorkingPlace & ">"
		If iMaxRow > 1 And flgSOHOFC = False Then sWorkingPlace = sWorkingPlace & "�y�Ζ��n" & dbWorkingPlaceSeq & "�z"

		If dbOrderType <> "0" Then
			sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
		ElseIf dbPlanTypeName = "mail" Then
			sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
		Else
			sWorkingPlace = sWorkingPlace & dbWorkingPlaceAddressAll
			If dbWorkingPlaceSection & dbWorkingPlaceTelephoneNumber <> "" Then
				sWorkingPlace = sWorkingPlace & "("
				If dbWorkingPlaceSection <> "" Then sWorkingPlace = sWorkingPlace & dbWorkingPlaceSection
				If dbWorkingPlaceSection <> "" And dbWorkingPlaceTelephoneNumber <> "" Then sWorkingPlace = sWorkingPlace 
				If dbWorkingPlaceTelephoneNumber <> "" Then sWorkingPlace = sWorkingPlace '& "TEL:" & dbWorkingPlaceTelephoneNumber
				sWorkingPlace = sWorkingPlace & ")"
			End If
			If dbMapFlag = "1" Then sWorkingPlace = sWorkingPlace & "&nbsp;[<span style=""color:#0045f9;cursor:pointer;"" onclick=""open('" & HTTPS_CURRENTURL & "map/showmap.asp?ordercode=" & dbOrderCode & "&wpseq=" & dbWorkingPlaceSeq & "', 'map', 'width=700,height=650');"">�n�}</span>]"
		End If

		'<�Ŋ�w>
		sNearbyStation = ""
		oRS2.Filter = "WorkingPlaceSeq = " & dbWorkingPlaceSeq
		If GetRSState(oRS2) = True Then
			sNearbyStation = GetNearbyStation(rDB, oRS2)
		End If
		oRS2.Filter = 0
		'</�Ŋ�w>
		'<�Ŋ񉈐�>
		sNearbyRailway = ""
		oRS3.Filter = "WorkingPlaceSeq = " & dbWorkingPlaceSeq
		If GetRSState(oRS3) = True Then
            'sNearbyRailway = GetNearbyRailway(rDB, oRS3)
			sNearbyRailway = GetNearbyRailway2(rDB, oRS3)
		End If
		oRS3.Filter = 0
		'</�Ŋ񉈐�>
		If sNearbyStation <> "" Then
			sWorkingPlace = sWorkingPlace & "<p class=""m0"""
			If flgSOHOFC = True Then
				sWorkingPlace = sWorkingPlace & " style=""display:inline;"""
			Else
				sWorkingPlace = sWorkingPlace & " style=""padding-left:15px;"""
			End If
			sWorkingPlace = sWorkingPlace & ">"
			sWorkingPlace = sWorkingPlace & "[�Ŋ�w]"
			sWorkingPlace = sWorkingPlace & sNearbyStation
			If flgSOHOFC = False Then sWorkingPlace = sWorkingPlace & "<br>"
			sWorkingPlace = sWorkingPlace & "[����]"
			sWorkingPlace = sWorkingPlace & sNearbyRailway
			sWorkingPlace = sWorkingPlace & "</p>"
		End If
		'</�Ζ��n>

		sWorkingPlace = sWorkingPlace & "</div>"
		oRS.MoveNext
	Loop

	'�]��
	If (dbOrderType = "0" Or dbOrderType = "2") And dbCompanyKbn <> "4" Then
		'ؽ�̔h�����l�[ �܂��� �h����Ђ̋��l�[�̏ꍇ�͕\�����Ȃ�

		dbTransfer = ChkStr(rRS.Collect("Transfer"))
		If dbTransfer <> "" Then
			If dbTransfer = "�L" Then
				dbTransfer = "�]�΂���"
			ElseIf dbTransfer = "��" Then
				dbTransfer = "�]�΂Ȃ�"
			End If
		End If
	End If
	If sWorkingPlace & sNearbyStationBlock & dbTransfer <> "" Then flgDspWorkingPlace = True
	'</�Ζ��n>

flgLine = False


%>	
	<img src="/img/order/tab_detail_li.png" class="tab_img">
	<table class="detail_table">
    <tbody>


<% If sWorkingType <> "" Then
	If flgLine = True Then sHTML = sHTML & ""
		flgLine = True
 %>
     <tr>
    <th class="dborder_bottom">
     �Ζ��`��
    </th>
    <td class="dborder_bottom">
	<p class="m0 get_job_type"><%= sWorkingType %></p>
	<% End If %>

	<% If dbTTPOrderFlag = "1" And sAfterWorkingType <> "" Then %>
	<p class="m0"><%= sAfterWorkingType %></p>
	<% End If %>
    
    <% If sWorkRange <> "" Then %>
    <p class="m0">���L���̏ꍇ�F<%= sWorkRange %></p>
	<% End If %>
   
	<% If sWorkingType <> "" Then %>
    </td>
    </tr>

	<% End If %>

	<%
		If flgDspJobType = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True
	%>
    <tr>
    <th class="dborder_bottom">�E��</th>
    <td class="dborder_bottom">
    
    <p class="m0"><strong><%= dbJobTypeDetail %></strong></p>
	<p class="m0"><%= sJobType %></p>
    
    
    </td>
    </tr>
    <% End If %>
    
	<%
		If flgDspSalary = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True
	%>
    <tr>
    <th class="dborder_bottom">���^</th>
    <td class="dborder_bottom">
    
   <% If sYearlyIncome <> "" Then %>
	<h5>�N��</h5>
	<p class="m0"><%= sYearlyIncome %></p>
    <% End If %>
    
    <% If sMonthlyIncome <> "" Then %>
	<h5>����</h5>
	<p class="m0"><%= sMonthlyIncome %></p>
    <% End If %>
    
    <% If sDailyIncome <> "" Then %>
	<h5>����</h5>
	<p class="m0"><%= sDailyIncome %></p>
    <% End If %>
    
    <% If sHourlyIncome <> "" Then %>
	<h5>����</h5>
	<p class="m0"><%= sHourlyIncome %></p>
    <% End If %>
    
    <% If dbPercentagePay <> "" Then %>
	<h5>������</h5>
	<p class="m0"><%= dbPercentagePay %></p>
    <% End If %>
    
    <% If sTrafficFee <> "" Then %>
	<h5>��ʔ�</h5>
	<p class="m0"><%= sTrafficFee %></p>
    <% End If %>
    
    <% If dbSalaryRemark <> "" Then %>
	<h5>���^���l</h5>
	<p class="m0"><%= dbSalaryRemark %></p>
    <% End If %>
    
    <% If sYearlyIncome & sMonthlyIncome & sDailyIncome & sHourlyIncome <> "" AND dbOrderCode <> "J0074418" Then %>
	<p class="m0">���Œ�z�͏����Ɋ֌W�Ȃ�������z�ł��B<br>(�N���̍Œ�z�͏����Ɋ֌W�Ȃ������錎���̍��v�ł��B)</p>
    <% End If %>
    <% '���X�̃L�����R����W�������O���Ƃ�
    If dbOrderCode = "J0074418" Then
        Call RegMailFromAccess(GetForm("mailfromaccess", 2))
    End If 
    %>
    </td>
    </tr>
    <% End If %>
    
    <%
		If flgDspTime = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True
	%>
    <tr>
    <th class="dborder_bottom">����</th>
    <td class="dborder_bottom">
    
    <% If sWorkingTime <> "" Then %>
	<h5>�A�Ǝ���</h5>
	<p class="m0"><%= sWorkingTime %></p>
    <% End If %>
    
    <% If dbWorkTimeRemark <> "" Then %>
	<h5>�A�Ǝ��Ԕ��l</h5>
	<p class="m0"><%= dbWorkTimeRemark %></p>
    <% End If %>
       
    </td>
    </tr>
    <% End If %>
    
    <%
		If flgDspHoliday = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True
	%>
    <tr>
    <th class="dborder_bottom">�x��</th>
    <td class="dborder_bottom">    
    
    <% If dbWeeklyHolidayType <> "" Then %>
	<h5>�x��</h5>
	<p class="m0"><%= dbWeeklyHolidayType %></p>
    <% End If %>
    
    <% If dbHolidayRemark <> "" Then %>
	<h5>�x�����l</h5>
	<p class="m0"><%= dbHolidayRemark %></p>
    <% End If %>
    
    </td>
    </tr>
    <% End If %>
    
    <%
		If flgDspHumanNumber = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True
	%>
    <tr>
    <th class="dborder_bottom">��W�l��</th>
    <td class="dborder_bottom">
	<p class="m0"><%= dbHumanNumber %></p>

    </td>
    </tr>
    <% End If %>  
    
    <%
		If flgDspWorkingPlace = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True
			%>
    <tr>
    <th class="dborder_bottom">�Ζ��n</th>
    <td class="dborder_bottom">

	<% If sWorkingPlace <> "" Then %>
	<h5>�Z��</h5>
	<p class="m0"><%= sWorkingPlace %></p>
	<% If sNearbyStationBlock <> "" Then
				sHTML = sHTML & sNearbyStationBlock
			End If 
			End If
			%>
   <% If dbTransfer <> "" Then %>
	<h5>�]��</h5>
	<p class="m0"><%= dbTransfer %></p>
    <% End If %>

    </td>
    </tr>
    <% End If %>     
    
     
    </tbody>
    </table>

<div class="to_top"><a class="stext_middle" href="#pagetop">���y�[�WTOP��</a></div>  

<%

	

End Function

'******************************************************************************
'�T�@�v�F���l�[�̕K�v�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�����ƃi�r/order/order_detail.asp
'���@�l�F
'���@���F2007/02/11 LIS K.Kokubo �쐬
'�@�@�@�F2008/11/12 LIS K.Kokubo �x�X�g�E�x�^�[�p�^�[���o��
'�@�@�@�F2010/10/18 LIS T.Ezaki  �x�^�[�p�^�[���̔�\���Ή�
'�@�@�@�F2012/03/12 LIS K.Kokubo ���ƔN�o��
'******************************************************************************
Function DspNeedCondition(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	
	Dim dbOrderCode			'���R�[�h
	Dim sCompanyCode		'��ƃR�[�h
	Dim sOrderType			'���l�[���
	Dim sCompanyKbn			'��Ƌ敪
	Dim dbTempOrderFlag		'�h���Č��t���O
	Dim dbBestMatchStr		'�x�X�g�p�^�[��
	Dim dbBetterMatchStr	'�x�^�[�p�^�[��
	Dim sAge				'�N���
	Dim sAgeMin				'�N���
	Dim sAgeMax				'�N����
	Dim sAgeReasonFlag		'�N��R�t���O
	Dim sAgeReason			'�N��R
	Dim sAgeReasonDetail	'�N������R�ڍ�
	Dim sFEHistory			'�w��
	Dim dbGraduateYearMin	'���ƔN����
	Dim dbGraduateYearMax	'���ƔN���
	Dim sSkillOS			'�n�r
	Dim sSkillApp			'�A�v���P�[�V����
	Dim sSkillDL			'�J������
	Dim sSkillDB			'�c�a
	Dim sSkillOther			'���̑��X�L��
	Dim sLicense			'���i
	Dim sLicenseOther		'���̑����i
	Dim sOtherNote			'���̑����L����

	Dim sClearSolid			'border�����p
	Dim flgLicense			'���C�Z���X�̍��ڂ̗L�� [True]�L [False]��
	Dim flgSkill			'�X�L���̍��ڂ̗L�� [True]�L [False]��
	Dim flgLine				'�������t���O
	Dim flgLine2			'�������t���O�Q

	DspNeedCondition = False

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	dbOrderCode = rRS.Collect("OrderCode")
	sCompanyCode = rRS.Collect("CompanyCode")
	sOrderType = rRS.Collect("OrderType")
	sCompanyKbn = rRS.Collect("CompanyKbn")
	dbTempOrderFlag = rRS.Collect("TempOrderFlag")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'<�x�X�g�E�x�^�[�p�^�[��>
	'�Љ�E�Љ�\��h���̂�
    '2014/05/12 �K�v�����̗��Ƀ}�b�`���O�|�C���g��ǉ����邽�߁A���L��if���ɔh���isOrderType = "1"�j��ǋL�F�ؑ�
	If sOrderType = "2" Or sOrderType = "3" Or sOrderType = "1" Then
		dbBestMatchStr = ChkStr(rRS.Collect("BestMatchStr"))
		dbBetterMatchStr = ChkStr(rRS.Collect("BetterMatchStr"))
	End If
	'</�x�X�g�E�x�^�[�p�^�[��>

	'******************************************************************************
	'�N�� start
	'------------------------------------------------------------------------------
	sAge = ""
	sAgeMin = ChkStr(rRS.Collect("AgeMin"))
	sAgeMax = ChkStr(rRS.Collect("AgeMax"))
	sAgeReasonFlag = ChkStr(rRS.Collect("AgeReasonFlag"))
	sAgeReason = ChkStr(rRS.Collect("AgeReason"))
	sAgeReasonDetail = Replace(ChkStr(rRS.Collect("AgeReasonDetail")), vbCrLf, "<br>")

	If dbTempOrderFlag = "1" Then
		sAge = "�h���Č��̂��߁A�N��f�ڂ��Ă��܂���B<br>"
		sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[�H]�����ɂ���</a>"
	ElseIf sAgeReasonFlag = "0" Or sAgeReasonFlag = "" Or (sAgeMin & sAgeMax = "") Then
		sAge = "�N��s��<br>"
		'sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[�H]�����ɂ���</a>"
	Else
		If sAgeMin <> "" Then sAgeMin = sAgeMin & "��"
		If sAgeMax <> "" Then sAgeMax = sAgeMax & "��"
		sAge = sAgeMin & "�`" & sAgeMax
		If sAgeReason <> "" Then sAge = sAge & "&nbsp;(" & sAgeReason & ")<br>"
		If sAgeReasonDetail <> "" Then sAge = sAge & sAgeReasonDetail & "<br>"
		sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[�H]�����ɂ���</a><br>"
	End If

	If sAge <> "" Then DspNeedCondition = True
	'------------------------------------------------------------------------------
	'�N�� end
	'******************************************************************************

	'******************************************************************************
	'�w�� start
	'------------------------------------------------------------------------------
	sFEHistory = ChkStr(rRS.Collect("HopeSchoolHistory"))
	dbGraduateYearMin = rRS.Collect("GraduateYearMin")
	dbGraduateYearMax = rRS.Collect("GraduateYearMax")

	If sFEHistory <> "" Then sFEHistory = sFEHistory & "���ȏ�"

	If dbGraduateYearMin + dbGraduateYearMax > 0 Then
		sFEHistory = sFEHistory & "<br>[���ƔN] "
		If dbGraduateYearMin > 0 Then
			sFEHistory = sFEHistory & dbGraduateYearMin & "�N��"
		End If
		sFEHistory = sFEHistory & " �` "
		If dbGraduateYearMax > 0 Then
			sFEHistory = sFEHistory & dbGraduateYearMax & "�N��"
		End If
	End If

	If sFEHistory <> "" Then DspNeedCondition = True
	'------------------------------------------------------------------------------
	'�w�� end
	'******************************************************************************

	'******************************************************************************
	'���i start
	'------------------------------------------------------------------------------
	sLicense = GetLicense(rDB, rRS)
	sLicenseOther = GetOrderNote(rDB, rRS, "OtherLicense")
	sLicenseOther = Replace(sLicenseOther, vbCrLf, "<br>")
	flgLicense = False
	If sLicense & sLicenseOther <> "" Then
		flgLicense = True
		DspNeedCondition = True
	End If
	'------------------------------------------------------------------------------
	'���i end
	'******************************************************************************

    '2014/04/25 ���}���i�ǉ� �r�c
	'******************************************************************************
	'���}���i start
	'------------------------------------------------------------------------------
	Dim sLicense_want

    sLicense_want = GetLicense_Want(rDB, rRS)

	If sLicense_want <> "" Then
		flgLicense = True
		DspNeedCondition = True
	End If
	'------------------------------------------------------------------------------
	'���}���i end
	'******************************************************************************


    'Dim sLicense_MustWant

    'If rRS.Collect("LicenseMustFlag") = "0" Then
    '    sLicense_MustWant = "����L�����ꂩ�̎��i��ۗL"
    'ElseIf rRS.Collect("LicenseMustFlag") = "1" Then
    '    sLicense_MustWant = "����L�S�Ă̎��i��ۗL"
    'Else
    '     sLicense_MustWant = "�� �K�{���i�������ݒ�"
    'End If



	'******************************************************************************
	'�X�L�� start
	'------------------------------------------------------------------------------
	sSkillOS = GetSkill(rDB, rRS, "OS")
	sSkillApp = GetSkill(rDB, rRS, "Application")
	sSkillDL = GetSkill(rDB, rRS, "DevelopmentLanguage")
	sSkillDB = GetSkill(rDB, rRS, "Database")
	sSkillOther = GetOrderNote(rDB, rRS, "OtherSkill")
	sSkillOther = Replace(sSkillOther, vbCrLf, "<br>")
	flgSkill = False
	If sSkillOS & sSkillApp & sSkillDL & sSkillDB & sSkillOther <> "" Then
		flgSkill = True
		DspNeedCondition = True
	End If
	'------------------------------------------------------------------------------
	'�X�L�� end
	'******************************************************************************

	'******************************************************************************
	'���̑����L���� start
	'------------------------------------------------------------------------------
	sOtherNote = ""
	If sOrderType = "0" Then
		sOtherNote = Replace(GetOrderNote(rDB, rRS, "OtherNote"), vbCrLf, "<br>")
		DspNeedCondition = True
	End If
	'------------------------------------------------------------------------------
	'���̑����L���� end
	'******************************************************************************

	flgLine = False

%>	
	<img src="/img/order/tab_detail_ne.png" class="tab_img">
	<table class="detail_table">
	<tbody>
    <%
	If dbBestMatchStr & dbBetterMatchStr <> "" Then
		If flgLine = True Then Response.Write ""
		flgLine = True
	
	%>
    <tr>
    <th class="dborder_bottom">�}�b�`���O�|�C���g
    <p class="smartNone">
    	[<span style="color:#0045F9;cursor:pointer;" onclick="window.open('<%= HTTPS_CURRENTURL %>infomation/matchingpoint.asp','matchingpoint','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=400,height=300');">�H</span>]
    </p>
    </th>
    <td class="dborder_bottom">
   <% If dbBestMatchStr <> "" Then %>
    <h4>�x�X�g</h4>
    <% Response.Write "<p class=""m0"">" & Replace(dbBestMatchStr, vbCrLf, "<br>") %></p>
    <% End If %>
    </td>
    </tr>
    <% End If %>
   
    <%
	If flgLine = True Then Response.Write ""
	flgLine = True
	
	%>
    <tr>
    <th class="dborder_bottom">�N��</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sAge %></p>    
    </td>
    </tr>
    
    
    <%
		If sFEHistory <> "" Then
		If flgLine = True Then Response.Write ""
		flgLine = True
	
	%>
    <tr>
    <th class="dborder_bottom">��]�w��</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sFEHistory %></p>    
    </td>
    </tr>
    <% End If

	sClearSolid = " style=""border-top-width:0px;"""
	If flgLicense = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write ""
		flgLine = True

		If sLicense <> "" Then
%>

            <tr>
            <th class="dborder_bottom">���i</th>
            <td class="dborder_bottom">
            <h4>�K�{���i</h4>
            <p class="m0"><%= sLicense %></p>
            
            </td>
            </tr>
 <% 	End If 


        '2014/04/25 ���}���i�ǉ� �r�c
		If sLicense_Want <> "" Then
%>

<tr>
    <th class="dborder_bottom"></th>
    <td class="dborder_bottom">
    <h4>���}���i</h4>
    <p class="m0"><%= sLicense_Want %></p>
    </td>
    </tr>
 <% 	End If 
 

		If sLicenseOther <> "" Then
	%>
    
    <tr>
    <th class="dborder_bottom"></th>
    <td class="dborder_bottom">
    <h4>���̑����i</h4>
     <p class="m0"><%= sLicenseOther %></p>
    </td>
    </tr>   
	<%   End If
	
	End If 
		
	sClearSolid = " style=""border-top-width:0px;"""
	If flgSkill = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write ""
		flgLine = True
		
		%>
        	
	<tr>
    <th class="dborder_bottom">�X�L��</th>
    <td class="dborder_bottom">
		<% If sSkillOS <> "" Then 
			
			Response.Write "<h5 class=""skill_h5"">OS</h5>" & vbCrLf
			Response.Write "<div>" & sSkillOS & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sSkillApp <> "" Then
			Response.Write "<h5 class=""skill_h5"">�A�v���P�[�V����</h5>" & vbCrLf
			Response.Write "<div>" & sSkillApp & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sSkillDL <> "" Then
			Response.Write "<h5 class=""skill_h5"">�J������</h5>" & vbCrLf
			Response.Write "<div>" & sSkillDL & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sSkillDB <> "" Then
			Response.Write "<h5 class=""skill_h5"">�f�[�^�x�[�X</h5>" & vbCrLf
			Response.Write "<div>" & sSkillDB & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sSkillOther <> "" Then
			Response.Write "<h5 class=""skill_h5"">���̑��X�L��</h5>" & vbCrLf
			Response.Write "<div><p class=""m0"">" & sSkillOther & "</p></div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If


		%>
    </td>
    </tr>   
	<% End If %>
    <%
		If sOtherNote <> "" Then
		If flgLine = True Then Response.Write ""
		flgLine = True

	%>
	<tr>
    <th class="dborder_bottom">���L����</th>
    <td class="dborder_bottom">
     <p class="m0"><%= sOtherNote %></p>
    </td>
    </tr>   
	<%
		sClearSolid = ""
	End If


%>
    </tbody>
    </table>
<div class="to_top"><a class="stext_middle" href="#pagetop">���y�[�WTOP��</a></div>  

<%
End Function

'******************************************************************************
'�T�@�v�F���l�[�̉�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F
'���@���F2007/02/11 LIS K.Kokubo �쐬
'******************************************************************************
Function DspHowToEntry(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim dbOrderCode			'���R�[�h
	Dim sCompanyCode		'��ƃR�[�h
	Dim sEntryInfo			'������@
	Dim sProcess1			'STEP1
	Dim sProcess2			'STEP2
	Dim sProcess3			'STEP3
	Dim sProcess4			'STEP4
	Dim sCSectionName		'���X�S������
	Dim sCPersonName		'���X�S���Җ�
	Dim sCTel				'���X�A����
	Dim sLis				'���X�S����
	Dim dbWValueURL			'�v�o�����[�̎��Ѝ̗p�y�[�W�t�q�k
	Dim flgEntryInfo		'�����񂪗L�邩������ [True]���� [False]�Ȃ�
	Dim flgProcess			'�I�l�菇���L�邩������ [True]���� [False]�Ȃ�
	Dim sClearSolid
	Dim flgLine				'�������t���O

	DspHowToEntry = False

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sOrderType = ChkStr(rRS.Collect("OrderType"))
	dbOrderCode = ChkStr(rRS.Collect("OrderCode"))
	sCompanyCode = rRS.Collect("CompanyCode")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'******************************************************************************
	'������@ start
	'------------------------------------------------------------------------------
	flgEntryInfo = False

	sEntryInfo = Replace(ChkStr(rRS.Collect("EntryInfo")), vbCrLf, "<br>")
	sEntryInfo = Replace(sEntryInfo, vbCr, "<br>")
	sEntryInfo = Replace(sEntryInfo, vbLf, "<br>")

	If sEntryInfo <> "" Then
		flgEntryInfo = True
		DspHowToEntry = True
	End If
	'------------------------------------------------------------------------------
	'������@ end
	'******************************************************************************

	'******************************************************************************
	'�I�l�菇 start
	'------------------------------------------------------------------------------
	flgProcess = False

	sProcess1 = ChkStr(rRS.Collect("Process1"))
	sProcess2 = ChkStr(rRS.Collect("Process2"))
	sProcess3 = ChkStr(rRS.Collect("Process3"))
	sProcess4 = ChkStr(rRS.Collect("Process4"))

	If sProcess1 & sProcess2 & sProcess3 & sProcess4 <> "" Then
		flgProcess = True
		DspHowToEntry = True
	End If
	'------------------------------------------------------------------------------
	'�I�l�菇 end
	'******************************************************************************

	'******************************************************************************
	'�A���� start
	'------------------------------------------------------------------------------
	sCSectionName = ChkStr(rRS.Collect("LisDepartment"))
	sCPersonName = ChkStr(rRS.Collect("EmployeeName"))
	sCTel = ChkStr(rRS.Collect("LisTelephoneNumber"))
	sLis = sCPersonName & "�m���X�������" & sCSectionName & "�n�@" & sCTel & "<br>(���̈Č��̓��X������Ђ����܂Ƃ߂Ă��܂��B)"
	DspHowToEntry = True
	'------------------------------------------------------------------------------
	'�A���� end
	'******************************************************************************

	'******************************************************************************
	'�v�o�����[�̎��Ѝ̗p�y�[�W�t�q�k start
	'------------------------------------------------------------------------------
	dbWValueURL = ChkStr(rRS.Collect("WValueURL"))
	If dbWValueURL <> "" Then
		DspHowToEntry = True
	End If
	'------------------------------------------------------------------------------
	'�v�o�����[�̎��Ѝ̗p�y�[�W�t�q�k end
	'******************************************************************************

	flgLine = False
	
	%>
    <img src="/img/order/tab_detail_ji.png" class="tab_img">
	<table class="detail_table">
	<tbody>
    <tr>
    <th class="dborder_bottom">���R�[�h</th>
    <td class="dborder_bottom">
    <p class="m0"><%= dbOrderCode %></p>
    </td>
	</tr>    
    
    <% If flgEntryInfo = True Then %>
    <!--<tr>
    <th class="dborder_bottom">������@</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sEntryInfo %></p>
    </td>
	</tr>   -->   
    <% End If %>
    
    <% If flgProcess = True Then %>
    <tr>
    <th class="dborder_bottom">�I�l�菇</th>
    <td class="dborder_bottom">
    
    	<table>
    		<tr>
            <% If sProcess1 <> "" Then %>
            	<td class="stepTd">�X�e�b�v1</td>
                <td><%= sProcess1 %></td>
            <% Else %>
            	<td class="stepTd">�X�e�b�v1</td>
                <td>�����ƃi�r�ɓo�^�܂��̓��O�C����A����{�^����育���咸�����ޑI�l�������܂��B</td>
            <% End If %>
        	</tr>
            <% If sProcess2 <> "" Then %>
            <tr>
            	<td class="stepTd">��</td>
                <td></td>
        	</tr>
            <tr>
            	<td class="stepTd">�X�e�b�v2</td>
                <td><%= sProcess2 %></td>
        	</tr>
            <% End If %>
            <% If sProcess3 <> "" Then %>
            <tr>
            	<td class="stepTd">��</td>
                <td></td>
        	</tr>
            <tr>
            	<td class="stepTd">�X�e�b�v3</td>
                <td><%= sProcess3 %></td>
        	</tr>
            <% End If %>
            <% If sProcess4 <> "" Then %>
            <tr>
            	<td class="stepTd">��</td>
                <td></td>
        	</tr>
            <tr>
            	<td class="stepTd">�X�e�b�v4</td>
                <td><%= sProcess4 %></td>
        	</tr>
			<% End If %>
        </table>
   
    </td>
	</tr>      
    <% End If %>
    
    <% If dbWValueURL <> "" Then %>
    <tr>
    <th class="dborder_bottom">���Ѝ̗p�y�[�W</th>
    <td class="dborder_bottom">
    <p class="m0"><a href="<%= dbWValueURL %>" target="_blank"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_wvalue.gif" border="0" alt="���Ѝ̗p�y�[�W"></a></p>
    </td>
	</tr>      
    <% End If %>
    
    
    </tbody>
    </table>
<div class="to_top"><a class="stext_middle" href="#pagetop">���y�[�WTOP��</a></div> 
<%

End Function


'******************************************************************************
'�T�@�v�F���l�[�̉�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F
'���@���F2013/09/09 LIS T.seki �쐬
'******************************************************************************
Function DspHowToEntry2(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim dbOrderCode			'���R�[�h
	Dim sCompanyCode		'��ƃR�[�h
	Dim sEntryInfo			'������@
	Dim sProcess1			'STEP1
	Dim sProcess2			'STEP2
	Dim sProcess3			'STEP3
	Dim sProcess4			'STEP4
	Dim sCSectionName		'���X�S������
	Dim sCPersonName		'���X�S���Җ�
	Dim sCTel				'���X�A����
	Dim sLis				'���X�S����
	Dim dbWValueURL			'�v�o�����[�̎��Ѝ̗p�y�[�W�t�q�k
	Dim flgEntryInfo		'�����񂪗L�邩������ [True]���� [False]�Ȃ�
	Dim flgProcess			'�I�l�菇���L�邩������ [True]���� [False]�Ȃ�
	Dim sClearSolid
	Dim flgLine				'�������t���O

	DspHowToEntry2 = False

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sOrderType = ChkStr(rRS.Collect("OrderType"))
	dbOrderCode = ChkStr(rRS.Collect("OrderCode"))
	sCompanyCode = rRS.Collect("CompanyCode")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'******************************************************************************
	'������@ start
	'------------------------------------------------------------------------------
	flgEntryInfo = False

	sEntryInfo = Replace(ChkStr(rRS.Collect("EntryInfo")), vbCrLf, "<br>")
	sEntryInfo = Replace(sEntryInfo, vbCr, "<br>")
	sEntryInfo = Replace(sEntryInfo, vbLf, "<br>")

	If sEntryInfo <> "" Then
		flgEntryInfo = True
		DspHowToEntry2 = True
	End If
	'------------------------------------------------------------------------------
	'������@ end
	'******************************************************************************

	'******************************************************************************
	'�I�l�菇 start
	'------------------------------------------------------------------------------
	flgProcess = False

	sProcess1 = ChkStr(rRS.Collect("Process1"))
	sProcess2 = ChkStr(rRS.Collect("Process2"))
	sProcess3 = ChkStr(rRS.Collect("Process3"))
	sProcess4 = ChkStr(rRS.Collect("Process4"))

	If sProcess1 & sProcess2 & sProcess3 & sProcess4 <> "" Then
		flgProcess = True
		DspHowToEntry2 = True
	End If
	'------------------------------------------------------------------------------
	'�I�l�菇 end
	'******************************************************************************

	'******************************************************************************
	'�A���� start
	'------------------------------------------------------------------------------
	sCSectionName = ChkStr(rRS.Collect("LisDepartment"))
	sCPersonName = ChkStr(rRS.Collect("EmployeeName"))
	sCTel = ChkStr(rRS.Collect("LisTelephoneNumber"))
	sLis = sCPersonName & "�m���X�������" & sCSectionName & "�n�@" & sCTel & "<br>(���̈Č��̓��X������Ђ����܂Ƃ߂Ă��܂��B)"
	DspHowToEntry2 = True
	'------------------------------------------------------------------------------
	'�A���� end
	'******************************************************************************

	'******************************************************************************
	'�v�o�����[�̎��Ѝ̗p�y�[�W�t�q�k start
	'------------------------------------------------------------------------------
	dbWValueURL = ChkStr(rRS.Collect("WValueURL"))
	If dbWValueURL <> "" Then
		DspHowToEntry2 = True
	End If
	'------------------------------------------------------------------------------
	'�v�o�����[�̎��Ѝ̗p�y�[�W�t�q�k end
	'******************************************************************************

	flgLine = False
	
	%>
   
	<table class="jCodeOnly">
    	<thead>
        	<td colspan="2">�����l�Ɋւ��邨�₢���킹�₨�j�����\���̍ۂɂ́A�K�����L�́u���R�[�h�v�����m�点���������B�u�����ƃi�r�������v�Ƃ���������Ē����܂��ƃX���[�X�ł��B</td>
        </thead>
        <tbody>
            <tr>
                <th class="dborder_bottom">���R�[�h</th>
                <td class="dborder_bottom">
                	<p class="m0"><%= dbOrderCode %></p>
                </td>
            </tr>    
        </tbody>
    </table>
<%

End Function

'******************************************************************************
'�T�@�v�F���l�[�̒S���ҘA������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�p���F
'���@�l�F
'���@���F2007/02/11 LIS K.Kokubo �쐬
'�@�@�@�F2009/04/02 LIS K.Kokubo ���[���ۋ��v�����̏ꍇ�͘A������\����
'******************************************************************************
Function DspContact(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim dbOrderCode			'���R�[�h
	Dim sCompanyCode		'��ƃR�[�h
	Dim sCompanyName		'��Ɩ���
	Dim sCompanyNameF		'��Ɩ��̃J�i
	Dim sCompanyKbn			'��Ƌ敪
	Dim sCompanySpeciality	'��Ɠ���
	Dim sCSectionName		'�d���̘A����S������
	Dim sCPersonPost		'�d���̘A����S���Җ�E
	Dim sCPersonName		'�d���̘A����S���Җ�
	Dim sCPersonNameF		'�d���̘A����S���҃J�i
	Dim sCTel				'�d���̘A����d�b�ԍ�
	Dim sCMail				'�d���̘A���惁�[���A�h���X
	Dim sPerson
	Dim sContact
	Dim sOrderType
	Dim dbPlanTypeName
	Dim flgLine				'�������t���O

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sCompanyCode = rRS.Collect("CompanyCode")
	sOrderType = rRS.Collect("OrderType")
	If sOrderType <> "0" Then Exit Function
	dbPlanTypeName = rRS.Collect("PlanTypeName")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'******************************************************************************
	'��Ж� start
	'------------------------------------------------------------------------------
	sCompanyName = rRS.Collect("CompanyName")
	sCompanyNameF = rRS.Collect("CompanyName_F")
	sCompanyKbn = rRS.Collect("CompanyKbn")
	sCompanySpeciality = rRS.Collect("CompanySpeciality")

	'Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
	'------------------------------------------------------------------------------
	'��Ж� end
	'******************************************************************************

	'******************************************************************************
	'�d���̘A���� start
	'------------------------------------------------------------------------------
	If sOrderType = "0" Then
		sCSectionName = ChkStr(rRS.Collect("ContactSectionName"))
		sCPersonPost = ChkStr(rRS.Collect("ContactPersonPost"))
		sCPersonName = ChkStr(rRS.Collect("ContactPersonName"))
		sCPersonNameF = ChkStr(rRS.Collect("ContactPersonName_F"))
		sCTel = ChkStr(rRS.Collect("ContactTelNumber"))
		sCMail = ChkStr(rRS.Collect("ContactMailAddress"))

		If sCompanyKbn = "2" Then
			'�l�މ�Ђ̋��l�[�̏ꍇ�́u���O�v�{�u�l�މ�Ж��v
			sPerson = sCPersonName & "&nbsp;(�l�މ�ЁF" & sCompanyName & ")"
		Else
			'��ʊ�Ƃ̋��l�[�̏ꍇ�́u���O�v�{�u�J�i�v
			sPerson = sCPersonName
			If sCPersonNameF <> "" Then sPerson = sPerson & "(" & sCPersonNameF & ")"
		End If
	End If

	sContact = ""
	If sCTel <> "" Then sContact = sContact & sCTel & "	<SPAN style='font-size:10px;'>�@���d�b���ł̂��₢���킹�̍ہA�u�����ƃi�r�������v�ƌ����ƃX���[�Y�ł��B</SPAN>"
	If sContact <> "" Then sContact = sContact & "<br>"
	If sCMail <> "" Then sContact = sContact & sCMail
	'------------------------------------------------------------------------------
	'�d���̘A����
	'******************************************************************************

	flgLine = False
	
	%>
    <img src="/img/order/tab_detail_tn.png" class="tab_img">
    <table class="detail_table">
	<tbody>
    <tr>
    <th class="dborder_bottom">�S���ҏ��</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sPerson %></p>
    </td>
    
	<% If sCSectionName <> "" Then %>
    <tr>
    <th class="dborder_bottom">�S������</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sCSectionName %></p>
    </td>
	</tr>     
   	<% End If %>
    
   	<% If dbPlanTypeName <> "mail" Then %>
    <tr>
    <th class="dborder_bottom">�A����</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sContact %></p>
    </td>
	</tr>     
   	<% End If %>
   
    </tbody>
    </table>
    <div class="to_top"><a class="stext_middle" href="#pagetop">���y�[�WTOP��</a></div> 
    
    <%

End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍׂ̐�y�C���^�r���[���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail.asp
'���@���F2008/01/30 LIS K.Kokubo
'******************************************************************************
Function DspElderInterview(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode
	Dim dbSeq
	Dim dbProfile
	Dim dbQuestion
	Dim dbAnswer
	Dim dbPublicFlag
	Dim dbPictureFlag

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")

	sSQL = "EXEC up_LstC_ElderInterview '" & dbOrderCode & "', '1'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = True Then
%>
<h3 id="interview_h3">��Ƃ���̃��b�Z�[�W</h3>
<div class="freeprblock">
<%
		Do While GetRSState(oRS) = True
			dbSeq = oRS.Collect("Seq")
			dbProfile = Replace(oRS.Collect("Profile"), vbCrLf, "<br>")
			dbQuestion = Replace(oRS.Collect("Question"), vbCrLf, "<br>")
			dbAnswer = Replace(oRS.Collect("Answer"), vbCrLf, "<br>")
			dbPublicFlag = oRS.Collect("PublicFlag")
			dbPictureFlag = oRS.Collect("PictureFlag")
%>
		
		<div class="interview">
        	
<%
			If dbPictureFlag = "1" Then
				'��y�ʐ^�L��
%>
			<h4 class="interview_h4"><%= dbProfile %></h4>
			<div>
				<img src="<%= HTTP_NAVI_CURRENTURL %>company/elderinterview/picture.asp?ordercode=<%= dbOrderCode %>&amp;seq=<%= dbSeq %>" alt="">
			</div>
			<h5 class="interview_p"><%= dbQuestion %></h5>
			<p><%= dbAnswer %></p>
			
			<br clear="both">
		</div>
<%
			Else
				'��y�ʐ^����
%>			
			<h4 class="interview_h4_no"><%= dbProfile %></h4>
            <h5 class="interview_h"><%= dbQuestion %></h5>
            <p><%= dbAnswer %></p>
            <br clear="both">
		</div>
<%
			End If
			oRS.MoveNext
		Loop
%>
</div>
<br>
<%
	End If
End Function

'******************************************************************************
'�T�@�v�F���X�̈Č��S���ҁA�R���T���������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/company_order.asp
'******************************************************************************
Function DspConsultantComment(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sCompanyCode			'��ƃR�[�h
	Dim sOrderType				'�󒍎��
	Dim sEmployeeCode			'�R���T���^���g�Ј��ԍ�
	Dim sEmployeeName			'�R���T���^���g��
    Dim sEmployeeFrigana		'�R���T���^���g���t���K�i
	Dim sBranchName				'�R���T���^���g�̋��_��
	Dim sTel					'�R���T���^���g�̋��_�̓d�b�ԍ�
	Dim sConsultantLink			'�R���T���Љ�y�[�W�ւ̃����N
	Dim sImg					'�R���T���^���g�̎ʐ^
	Dim sComment				'�R���T���^���g�R�����g
	Dim sConsultantPublicFlag	'�R���T���^���g�̏Љ�y�[�W�f�ڃt���O
	Dim sPictureFlag			'�R���T���^���g�ʐ^�t���O
	Dim sTitle					'�^�C�g���@�����������"���̋��l����S�����Ă���R���T���^���g�̏���"�@�Ȃ����"�S���ҘA����"
	Dim sClearSolid
	Dim flgLine

	If GetRSState(rRS) = False Then Exit Function

	flgLine = False

	sCompanyCode = rRS.Collect("CompanyCode")
	sOrderType = rRS.Collect("OrderType")

	If sOrderType <> "0" Then
		'******************************************************************************
		'�R���T���^���g start
		'------------------------------------------------------------------------------
		'���X�󒍕[�̏ꍇ�́u���X�S���Җ��v�{�u���X�S���҃J�i�v
		sEmployeeCode = ChkStr(rRS.Collect("EmployeeCode"))
		sEmployeeName = ChkStr(rRS.Collect("EmployeeName"))
        sEmployeeFrigana = ChkStr(rRS.Collect("EmployeeFrigana"))
		sBranchName = ChkStr(rRS.Collect("LisDepartment"))
		sTel = ChkStr(rRS.Collect("LisTelephoneNumber"))

        '2017/05/16�@�썪����̂ݕ������ύX���W�b�N by �v��
        '2015/09/14�@�썪����̂ݕ������ύX���W�b�N by �r�c
        '2014/09/11�@��������̂ݕ������ύX���W�b�N by tanizawa
        '-------------------------
        Dim sSQL : sSQL =""
        Dim flgQE
	Dim sError
        Dim oRS

        IF sEmployeeCode = "L0000488" Then '��������
        	sSQL = "sp_GetDataJobType '" & rRS.Collect("OrderCode") & "'"
	        flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
            If GetRSState(rRS) = True Then
	            Do While GetRSState(oRS) = True
		            If Left(oRS.Collect("JobTypeCode"),2) = "13" Then
                        sBranchName = "���f�B�J���`�[��"
                        Exit Do
                    End IF
		            oRS.MoveNext
	            Loop
            End IF
            Call RSClose(oRS)
        End IF

        IF sEmployeeCode = "L0000381" Then '�썪����

        	sSQL = "sp_GetDataJobType '" & rRS.Collect("OrderCode") & "'"
	        flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
            If GetRSState(rRS) = True Then
	            Do While GetRSState(oRS) = True

		            If Left(oRS.Collect("JobTypeCode"),2) = "07" Then
                        sBranchName = "��v�������E�ŗ��m�@�l�E�č��@�l�`�[��"
                        Exit Do
                    Else 
                        sBranchName = "�l�ޏЉ�`�[��"

                    End IF
		            oRS.MoveNext
	            Loop
            End IF
            Call RSClose(oRS)
        End IF

        '-------------------------

		sImg = "<img src=""/consultant/consultantimage.asp?ec=" & sEmployeeCode & """ alt=""���̋��l����S�����Ă���R���T���^���g"" border=""1"" width=""180"" height=""180"" style=""border-color:#666666;"">"
		sComment = Replace(ChkStr(rRS.Collect("ConsultantComment")), vbCrLf, "<br>")
		sComment = Replace(sComment, vbCr, "<br>")
		sComment = Replace(sComment, vbLf, "<br>")
		sConsultantPublicFlag = ChkStr(rRS.Collect("ConsultantPublicFlag"))
		sPictureFlag = ChkStr(rRS.Collect("ConsultantPictureFlag"))

		'2016/06/22�@�ؑ��F�R���T���Љ�y�[�W�ւ̃����N�폜
		sConsultantLink = Split(sEmployeeName,"�@")(0)
		'If sConsultantPublicFlag = "1" Then
		'	sConsultantLink = "<a href=""" & HTTP_NAVI_CURRENTURL & "consultant/consultantdetail.asp?ec=" & sEmployeeCode & """>" & sConsultantLink & "</a>"
		'End If
        if sEmployeeFrigana = "" Then
            sConsultantLink = Split(sEmployeeName,"�@")(0)
        Else
            sConsultantLink = "<ruby><rb>" & Split(sEmployeeName,"�@")(0) & "</rb><rp>�i</rp><rt>" & Split(sEmployeeFrigana,"�@")(0) & "</rt><rp>�j</rp></ruby>"    
        End If
        

		sConsultantLink = sConsultantLink & "&nbsp;(�l�މ�ЁF���X�������)"
		'------------------------------------------------------------------------------
		'�R���T���^���g end
		'******************************************************************************

		sTitle = "�S���ҘA����"
		If sComment <> "" Then sTitle = "���̋��l����S�����Ă���R���T���^���g�̏���"
	%>
    <img src="/img/order/tab_detail_cn.png" class="tab_img">
	<table class="detail_table">
	<tbody>
    <tr>
    <th class="dborder_bottom">�R���T���^���g</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sConsultantLink %></p>
    </td>
    <tr>
    <th class="dborder_bottom">�S������</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sBranchName %></p>
    </td>
	</tr>     
    <tr>
    <th class="dborder_bottom">�A����</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sTel %><span>�����₢���킹�̍ہA��L�u���R�[�h�v�Ɓu�����ƃi�r�������v�Ƃ���������ĉ�����ƃX���[�Y�ł��B</span></p>
    </td>
	</tr>     
   	<% End If %>
    
    <% If sComment <> "" Then %>
    <tr>
    <th class="dborder_bottom">����</th>
    <td class="dborder_bottom">
    <p class="m0"><%= sComment %></p>
    </td>
	</tr>     
   	<% End If %>
   
    </tbody>
    </table>
    <div class="to_top"><a class="stext_middle" href="#pagetop">���y�[�WTOP��</a></div> 
    <%


End Function

'******************************************************************************
'�T�@�v�F�ŐV���[�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/company_order.asp
'******************************************************************************
Function DspNewMail(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sDateTime
	Dim sSubject
	Dim sDetail
	Dim flgLine

	DspNewMail = False

	If GetRSState(rRS) = False Then Exit Function

	flgLine = False

	If vUserType = "staff" THen
		sSQL = "up_DtlMailHistory_Order '" & vUserID & "', '" & rRS.Collect("OrderCode") & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			sDateTime = GetDateStr(oRS.Collect("SendDay"), "/") & "�@" & GetTimeStr(oRS.Collect("SendDay"), ":")
			sSubject = ChkStr(oRS.Collect("Subject"))
			sDetail = Replace(ChkStr(oRS.Collect("Body")), vbCrLf, "<br>")
			sDetail = Replace(sDetail, vbCr, "<br>")
			sDetail = Replace(sDetail, vbLf, "<br>")
			Response.Write "<h3 class=""sp"">�ŐV�̑��M�ς݃��[��</h3>"
			If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
			Response.Write "<div class=""category1""><h4>���M����</h4></div>"
			Response.Write "<div class=""value1""><p class=""m0"">" & sDateTime & "</p></div>"
			Response.Write "<div style=""clear:both;""></div>"
			If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
			Response.Write "<div class=""category1""><h4>�T�u�W�F�N�g</h4></div>"
			Response.Write "<div class=""value1""><p class=""m0"">" & sSubject & "</p></div>"
			Response.Write "<div style=""clear:both;""></div>"
			If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
			Response.Write "<div class=""category1""><h4>���e</h4></div>"
			Response.Write "<div class=""value1""><p class=""m0"">" & sDetail & "</p></div>"
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "<br>"
		End If
	End If

	Call RSClose(oRS)

	DspNewMail = True
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̋Ζ��`�ԕ���
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�쐬�ҁFLis Kokubo
'�쐬���F2006/05/08
'���@�l�F
'�g�p���Fstaff/company_detail.asp
'******************************************************************************
Function GetWorkingType(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sWorkingType

	If GetRSState(rRS) = False Then Exit Function

	sOrderCode = rRS.Collect("OrderCode")
	sWorkingType = ""
	sSQL = "sp_GetDataWorkingType '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	Do While GetRSState(oRS) = True
		sWorkingType = sWorkingType & oRS.Fields("WorkingTypeName").Value

		'���X�Љ�or�Љ���'�]����If (rRS.Fields("OrderType") ="" and rRS.Fields("Companykbn") = "2") or (rRS.Fields("OrderType") ="2") Then
		If (rRS.Collect("OrderType") ="0" And rRS.Collect("Companykbn") = "2") Or (rRS.Collect("OrderType") ="2") Then
			Select Case oRS.Collect("WorkingTypeCode")
				Case "001": sWorkingType = sWorkingType & "<span class=""smartNone"">�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/koyoukeitai_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")' class=""haken_tr"">�h���Ƃ�</a>�z</span>" 
				Case "002","003": sWorkingType = sWorkingType & "<span class=""smartNone"">�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")' class=""shokai_tr"">�l�ޏЉ�Ƃ�</a>�z</span>" 
				Case "004": sWorkingType = sWorkingType & "<span class=""smartNone"">�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/syoukaiyotei_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>�Љ�\��h���Ƃ�</a>�z</span>" 
			End Select
		End If

		oRS.MoveNext
		If GetRSState(oRS) = True Then sWorkingType = sWorkingType & "<br>"
	Loop
	Call RSClose(oRS)

	GetWorkingType = sWorkingType
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̐E�핔��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�쐬�ҁFLis Kokubo
'�쐬���F2006/05/08
'���@�l�F
'�g�p���Fstaff/company_detail.asp
'******************************************************************************
Function GetJobType(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sJobType

	If GetRSState(rRS) = False Then Exit Function

	sOrderCode = rRS.Collect("OrderCode")
	sJobType = ""

	sSQL = "sp_GetDataJobType '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

    Dim count 
    count = 1
	Do While GetRSState(oRS) = True
		sJobType = sJobType & "(" & count & ") " & oRS.Collect("JobTypeName") & ""
        count = count + 1
		oRS.MoveNext
		If GetRSState(oRS) = True Then sJobType = sJobType & "<br>"
	Loop
	Call RSClose(oRS)

	GetJobType = sJobType
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̋Ζ��`�ԕ���
'���@���FrDB	�F�ڑ�����DBConnection
'�@�@�@�FrRS	�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'���@�l�F
'�X�@�V�F2006/05/08 LIS K.Kokub �쐬
'�@�@�@�F2009/11/17 LIS K.Kokubo FC,SOHO�Č��̏ꍇ�͋Ζ����Ԃ�Ԃ��Ȃ�
'******************************************************************************
Function GetWorkingTime(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sWST
	Dim sWET

	Dim sWorkingTime

	If GetRSState(rRS) = False Then Exit Function

	sWorkingTime = ""
    Dim count
    count = 1
	If rRS.Collect("FCSOHOOrderFlag") = "0" Then
		sSQL = "sp_GetDataWorkingTime '" & rRS.Collect("OrderCode") & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			sWST = ChkStr(oRS.Collect("DspWorkStartTime"))
			sWET = ChkStr(oRS.Collect("DspWorkEndTime"))
			If sWST & sWET <> "" Then
				sWorkingTime = sWorkingTime & "(" & count & ") " & sWST & "�`" & sWET
                count = count + 1
			End If
			oRS.MoveNext
			If GetRSState(oRS) = True And sWST & sWET <> "" Then sWorkingTime = sWorkingTime & "<br>"
		Loop
		Call RSClose(oRS)
	End If

	GetWorkingTime = sWorkingTime
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̍Ŋ�w����
'���@���FrDB	�F�ڑ�����DBConnection
'�@�@�@�FrRS	�Fup_LstC_NearbyStation�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvWPSeq	�F�Ζ��n�ԍ�
'�g�@�p�F�i�r/include/func_order.asp
'���@�l�F
'���@���F2006/05/08 LIS K.Kokubo �쐬
'�@�@�@�F2008/10/22 LIS K.Kokubo ���l�[�Ζ��n�������Ή�
'******************************************************************************
Function GetNearbyStation(ByRef rDB, ByRef rRS)
	Dim dbWorkingPlaceSeq
	Dim dbStationName
	Dim dbToStationTime
	Dim dbToStationRemark

	Dim idx
	Dim sStation
	Dim sToStation
	Dim iStation

	If GetRSState(rRS) = False Then Exit Function

	iStation = 0
	sStation = ""
	Do While GetRSState(rRS) = True
		dbWorkingPlaceSeq = rRS.Collect("WorkingPlaceSeq")
		dbStationName = ChkStr(rRS.Collect("StationName"))
		dbToStationTime = ChkStr(rRS.Collect("ToStationTime"))
		dbToStationRemark = ChkStr(rRS.Collect("ToStationRemark"))
		iStation = iStation + 1

		sToStation = ""
		If dbToStationTime <> "" Then sToStation = dbToStationTime & "��"
		If dbToStationRemark <> "" Then sToStation = dbToStationRemark & sToStation
		If sToStation <> "" Then sToStation = "(" & sToStation & ")"

		If sStation <> "" Then sStation = sStation & "/"
		sStation = sStation & dbStationName & "�w" & sToStation

		rRS.MoveNext
	Loop

	GetNearbyStation = sStation
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̍Ŋ񉈐�����
'���@���FrDB	�F�ڑ�����DBConnection
'�@�@�@�FrRS	�Fup_LstC_NearbyRailwayLine�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�g�@�p�F�i�r/include/func_order.asp
'���@�l�F
'���@���F2006/05/08 LIS K.Kokubo �쐬
'�@�@�@�F2008/10/22 LIS K.Kokubo ���l�[�Ζ��n�������Ή�
'******************************************************************************
Function GetNearbyRailway(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbWorkingPlaceSeq
	Dim dbRailwayLineName2

	Dim idx
	Dim iRowCnt
	Dim sRailway
	Dim iRailway

	If GetRSState(rRS) = False Then Exit Function

	iRowCnt = rRS.RecordCount
	iRailway = 0
	sRailway = ""
	Do While GetRSState(rRS) = True And iRailway < 3
		dbWorkingPlaceSeq = rRS.Collect("WorkingPlaceSeq")
		dbRailwayLineName2 = rRS.Collect("RailwayLineName2")
		iRailway = iRailway + 1

		If sRailway <> "" Then sRailway = sRailway & ","
		sRailway = sRailway & dbRailwayLineName2

		rRS.MoveNext
	Loop
	If iRowCnt > 3 Then sRailway = sRailway & "&nbsp;��"

	GetNearbyRailway = sRailway
End Function


'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̍Ŋ񉈐�����
'���@���FrDB	�F�ڑ�����DBConnection
'�@�@�@�FrRS	�Fup_LstC_NearbyRailwayLine�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�g�@�p�F�i�r/include/func_order.asp
'���@�l�F
'���@���F2015/09/14 �r�c ���֐����R�s�[���A�V�����͕\�����Ȃ��悤�Ɏd�l�ǉ�
'******************************************************************************
Function GetNearbyRailway2(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbWorkingPlaceSeq
	Dim dbRailwayLineName2

	Dim idx
	Dim iRowCnt
	Dim sRailway
	Dim iRailway

	If GetRSState(rRS) = False Then Exit Function

	'iRowCnt = rRS.RecordCount

    iRowCnt = 0
	iRailway = 0
	sRailway = ""
	Do While GetRSState(rRS) = True And iRailway < 3
		dbWorkingPlaceSeq = rRS.Collect("WorkingPlaceSeq")
		dbRailwayLineName2 = rRS.Collect("RailwayLineName2")
		

		If sRailway <> "" Then sRailway = sRailway & ","
		
        If InStr(dbRailwayLineName2 ,"�V����") Then
            '�V�����͕\�����Ȃ�
        Else
            sRailway = sRailway & dbRailwayLineName2

            iRailway = iRailway + 1
            iRowCnt = iRowCnt + 1
        End If

		rRS.MoveNext
	Loop
	If iRowCnt > 3 Then sRailway = sRailway & "&nbsp;��"

	GetNearbyRailway2 = sRailway
End Function


'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̃X�L������
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�쐬�ҁFLis Kokubo
'�쐬���F2006/05/08
'���@�l�F
'�g�p���F
'******************************************************************************
Function GetSkill(ByRef rDB, ByRef rRS, ByVal vCategoryCode)
	Const SKILLCOL = 2

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim idx
	Dim sSkill
	Dim iSkill

	If GetRSState(rRS) = False Then Exit Function

	iSkill = 0
	sSkill = ""
	sSQL = "sp_GetDataSkill '" & rRS.Collect("OrderCode") & "', '" & vCategoryCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iSkill = iSkill + 1

		sSkill = sSkill & "<p style=""min-width:25%; max-width:50%; float:left; height:40px;""><span style=""color:#339933;"">��</span> " & oRS.Collect("SkillName")
		If ChkStr(oRS.Collect("Period")) <> "" Then
			sSkill = sSkill & "<br>�@" & oRS.Collect("Period") & "�N�ȏ�͏���"
		End If
		sSkill = sSkill & "</p>"
		If iSkill Mod SKILLCOL = 0 Then sSkill = sSkill & ""

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'���r���[�ŏI������ꍇ�̒���
	If sSkill <> "" And iSkill Mod SKILLCOL <> 0 Then sSkill = sSkill & "<br clear=""all"">"

	GetSkill = sSkill
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̎��i����
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�쐬�ҁFLis Kokubo
'�쐬���F2006/05/08
'���@�l�F
'******************************************************************************
Function GetLicense(ByRef rDB, ByRef rRS)
	Const LICENSECOL = 2

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim idx
	Dim iLicense
	Dim sLicense

	If GetRSState(rRS) = False Then Exit Function

	iLicense = 0
	sLicense = ""

	sSQL = "sp_GetDataLicense '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iLicense = iLicense + 1

		sLicense = sLicense & "<p style=""width:50%; float:left;"">(" & iLicense & ") " & oRS.Collect("LicenseName") & "</p>"
		If iLicense Mod LICENSECOL = 0 Then sLicense = sLicense & "<br clear=""all"">"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

    '2014/07/17 �K�{���i�t���O�ǉ� ����
	'******************************************************************************
	'�K�{���i�t���O start
	'------------------------------------------------------------------------------
    If sLicense <> "" Then
        if iLicense > 1 Then
            If rRS.Collect("LicenseMustFlag") = "0" Then
                sLicense = sLicense & "<p style=""width:50%; float:left;"">����L�����ꂩ�̎��i��ۗL���Ă��邱��</p>"
            ElseIF rRS.Collect("LicenseMustFlag") = "1" Then
                sLicense = sLicense & "<p style=""width:50%; float:left;"">����L�S�Ă̎��i��ۗL���Ă��邱��</p>"
            End If
        End IF
    End If
    '------------------------------------------------------------------------------
	'�K�{���i�t���O end
	'******************************************************************************

	'���r���[�ŏI������ꍇ�̒���
	If sLicense <> "" And iLicense Mod LICENSECOL <> 0 Then sLicense = sLicense & "<br clear=""all"">"

	GetLicense = sLicense
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̊��}���i����
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�쐬�ҁFLis ikeda
'�쐬���F2014/04/25
'���@�l�F
'******************************************************************************
Function GetLicense_Want(ByRef rDB, ByRef rRS)
	Const LICENSECOL = 2

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim idx
	Dim iLicense
	Dim sLicense

	If GetRSState(rRS) = False Then Exit Function

	iLicense = 0
	sLicense = ""

	sSQL = "sp_GetDataLicense_Want '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iLicense = iLicense + 1

		sLicense = sLicense & "<p style=""width:50%; float:left;"">(" & iLicense & ") " & oRS.Collect("LicenseName") & "</p>"
		If iLicense Mod LICENSECOL = 0 Then sLicense = sLicense & "<br clear=""all"">"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'���r���[�ŏI������ꍇ�̒���
	If sLicense <> "" And iLicense Mod LICENSECOL <> 0 Then sLicense = sLicense & "<br clear=""all"">"

	GetLicense_Want = sLicense
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̂��̑����擾
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvCode			�FC_Note�e�[�u���� Code �t�B�[���h�l
'�쐬�ҁFLis Kokubo
'�쐬���F2006/05/08
'���@�l�F
'******************************************************************************
Function GetOrderNote(ByRef rDB, ByRef rRS, ByVal vCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sNote

	If GetRSState(rRS) = False Then Exit Function

	sSQL = "sp_GetDataNote '" & rRS.Collect("OrderCode") & "', '"  & vCode &"'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sNote = oRS.Collect("Note")
	End If
	Call RSClose(oRS)

	GetOrderNote = sNote
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍׂ̃^�C�g���ƃf�B�X�N���v�V�������擾
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/12
'�߂�l�FrTitle			�F�^�C�g���i��̓I�E�햼�j
'�@�@�@�FrDescription	�F�������i�S���Ɩ��j
'�g�p���F�����ƃi�r/order/order_detail.asp
'���@�l�F
'******************************************************************************
Function GetOrderTitle(ByRef rDB, ByVal vOrderCode, ByRef rTitle, ByRef rKeywords, ByRef rDescription)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sWorkingType
	
	Dim rRS

	sSQL = "EXEC up_DtlOrderTitle '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		'rTitle = ChkStr(oRS.Collect("JobTypeDetail")) & "&nbsp;" & ChkStr(oRS.Collect("PrefectureName"))
		rTitle = ChkStr(oRS.Collect("JobTypeDetail")) & "&nbsp;" & ChkStr(oRS.Collect("CatchCopy"))
		rKeywords = "���l���,�]�E," & ChkStr(oRS.Collect("JobTypeDetail")) & "," & ChkStr(oRS.Collect("PrefectureName"))
		If ChkStr(oRS.Collect("JobTypeName")) <> "" Then rKeywords = rKeywords & "," & ChkStr(oRS.Collect("JobTypeName"))
		If ChkStr(oRS.Collect("WorkingTypeName")) <> "" Then rKeywords = rKeywords & "," & ChkStr(oRS.Collect("WorkingTypeName"))
		rDescription = "�]�E�E���l���F" & ChkStr(oRS.Collect("BusinessDetail"))
		If rDescription = "" Then rDescription = "�]�E�E���l���F" & ChkStr(oRS.Collect("JobTypeDetail"))
	End If
	Call RSClose(oRS)

	If rTitle <> "" Then rTitle = rTitle & "&nbsp;"
	rTitle = rTitle & sWorkingType

	GetOrderTitle = flgQE
End Function

'******************************************************************************
'�T�@�v�F�X�L���̊e���ڕ\��
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/14
'�߂�l�F
'�@�@�@�F
'�g�p���F�����ƃi�r/order/order_detail.asp
'���@�l�F
'******************************************************************************
Function GetSkillList(ByVal vTitleImg, ByVal vTitleAlt, ByVal vSkill)
	GetSkillList = ""
	If Len(vSkill) = 0 Then Exit Function
	GetSkillList = "<tr><td valign=""top""><img src=""" & vTitleImg & """ alt=""" & vTitleAlt & """ width=""50"" height=""12""></td><td style=""padding-left:5px;"">" & vSkill & "</td></tr>"
End Function

'******************************************************************************
'�T�@�v�F���R�����h���d�����ꗗ�o��
'���@���FrDB		�FDB�ڑ��I�u�W�F�N�g
'�@�@�@�FvUserType	�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID	�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvOrderCode	�F�{�������l�[�̏��R�[�h
'�@�@�@�FvRCMD		�F���R�����h��� ["1"]����Ȃ��d���������Ă܂� ["2"]�߂������̂��d�����
'�@�@�@�FvMyOrder	�F���Ћ��l�[���ۂ� ["1"]���Ћ��l�[
'�߂�l�F
'�쐬���F2007/05/31
'�쐬�ҁFLis Kokubo
'���@�l�F
'���@���F
'******************************************************************************
Function DspRecommendOrderList(ByRef rDB, ByVal vUserType, ByVal vUserID, ByVal vOrderCode, ByVal vRCMD, ByVal vMyOrder)
	Const MAXCOLS = 3

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sTitle
	Dim idx			'���[�v�J�E���g�A�b�v�ϐ�
	Dim iCols		'��
	Dim aPadding(2)	'�e��̃p�f�B���O
	Dim aJobTypeDetail()
	Dim aCompanyName()
	Dim aImg()
	Dim aWorkingTypeIcon()
	Dim aWorkingPlace()
	Dim aStation()
	Dim aYearlyIncome()
	Dim aMonthlyIncome()
	Dim aDailyIncome()
	Dim aHourlyIncome()

	If vMyOrder = "1" Then Exit Function

	Select Case vRCMD
		Case "1"
			sSQL = "up_SearchRelationAccessOrder '" & vOrderCode & "'"
			sTitle = "���̋��l���������l�͂���ȋ��l�������Ă��܂�"
		Case "2"
			sSQL = "up_SearchHighRelationOrder '" & vOrderCode & "'"
			sTitle = "���̋��l���̏����ɋ߂����l���"
		Case Else
			Exit Function
	End Select

	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = False Then Exit Function
%>
<h2 class="ssubtitle"><%= sTitle %></h2>
<div class="subcontent" style="margin-bottom:15px;">
<%
	Call DspOrderListDetail3(rDB, oRS, 3, 1, vRCMD)
%>
</div>
<%
End Function

'******************************************************************************
'�T�@�v�F���R�����h�̋��l�[�ꗗ�́A���l�[���̊e���ځi�E��A��Ɩ��Ȃǁj���擾
'���@���FrDB		�FDB�ڑ��I�u�W�F�N�g
'�@�@�@�FrRS		�F���l�[�ꗗ�̃��R�[�h�Z�b�g
'�@�@�@�FvRCMD		�F���R�����h��� ["1"]����Ȃ��d���������Ă܂� ["2"]�߂������̂��d�����
'�o�@�́FrJobTypeDetail		�F��̓I�E�햼
'�@�@�@�FrCompanyName		�F��Ɩ�
'�@�@�@�FrImg				�F��ƃC���[�W
'�@�@�@�FrWorkingTypeIcon	�F�Ζ��`�ԃA�C�R��
'�@�@�@�FrWorkingPlace		�F�Ζ��n
'�@�@�@�FrStation			�F�Ŋ�w '2008/10/22 LIS K.Kokubo �s�g�p
'�@�@�@�FrYearlyIncome		�F�N��
'�@�@�@�FrMonthlyIncome		�F����
'�@�@�@�FrDailyIncome		�F����
'�@�@�@�FrHourlyIncome		�F����
'�߂�l�F
'���@�l�F
'���@���F2007/05/31 LIS K.Kokubo �쐬
'�@�@�@�F2008/10/22 LIS K.Kokubo �Ζ��n�������ɂ��C��
'******************************************************************************
Function GetRecommendValues(ByRef rDB, ByRef rRS, ByVal vRCMD, ByRef rJobTypeDetail, ByRef rCompanyName, ByRef rImg, ByRef rWorkingTypeIcon, ByRef rWorkingPlace, ByRef rStation, ByRef rYearlyIncome, ByRef rMonthlyIncome, ByRef rDailyIncome, ByRef rHourlyIncome)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode			'���R�[�h
	Dim dbCompanyCode		'��ƃR�[�h
	Dim dbOrderType			'�󒍋敪
	Dim dbCompanyKbn		'��Ћ敪
	Dim dbCompanyName		'��Ɩ�
	Dim dbCompanyNameF		'��Ɩ��J�i
	Dim dbCompanySpeciality	'��Ɩ��i�����j
	Dim dbJobTypeDetail		'��̓I�E�햼(alt��title�ŏo�͂���)
	Dim dbYearlyIncomeMin	'�N������
	Dim dbYearlyIncomeMax	'�N�����
	Dim dbMonthlyIncomeMin	'��������
	Dim dbMonthlyIncomeMax	'�������
	Dim dbDailyIncomeMin	'��������
	Dim dbDailyIncomeMax	'�������
	Dim dbHourlyIncomeMin	'��������
	Dim dbHourlyIncomeMax	'�������
	Dim dbWorkingPlacePrefectureCode
	Dim dbWorkingPlacePrefectureName
	Dim dbWorkingPlaceCity
	Dim dbImageLimit

	Dim sViewJobTypeDetail	'���E�҂Ɍ������̓I�E�햼(����������̓J�b�g�����)
	Dim sYearlyIncome		'�N��
	Dim sMonthlyIncome		'����
	Dim sDailyIncome		'����
	Dim sHourlyIncome		'����
	Dim sWorkingTypeIcon	'�Ζ��`�ԃA�C�R������
	Dim sWorkingPlace		'�Ζ��n
	Dim sImg				'�摜URL

	Dim idx
	Dim sURL				'���l�[�ڍׂ�URL
	Dim sAlign				'�g�� [vCols = 1]left [vCols = vMaxCols]right [����ȊO]center

	If GetRSState(rRS) = False Then Exit Function

	sURL = HTTPS_CURRENTURL & "order/order_detail.asp"

	sSQL = "up_DtlOrder '" & rRS.Collect("OrderCode") & "', ''"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'���R�[�h
	dbOrderCode = ChkStr(oRS.Collect("OrderCode"))
	'��ƃR�[�h
	dbCompanyCode = ChkStr(oRS.Collect("CompanyCode"))
	'�󒍋敪
	dbOrderType = ChkStr(oRS.Collect("OrderType"))
	'��Ƌ敪
	dbCompanyKbn = ChkStr(oRS.Collect("CompanyKbn"))
	'��Ɩ�, ��Ɩ��J�i
	dbCompanyName = ChkStr(oRS.Collect("CompanyName"))
	dbCompanyNameF = ChkStr(oRS.Collect("CompanyName_F"))
	dbCompanySpeciality = ChkStr(oRS.Collect("CompanySpeciality"))
	Call SetOrderCompanyName(dbCompanyName, dbCompanyNameF, dbOrderType, dbCompanyKbn, dbCompanySpeciality)
	'��̓I�E�햼
	dbJobTypeDetail = ChkStr(oRS.Collect("JobTypeDetail"))
	sViewJobTypeDetail = dbJobTypeDetail
	If Len(sViewJobTypeDetail) > 14 Then sViewJobTypeDetail = Left(sViewJobTypeDetail, 14) & ".."
	'�ʐ^
	dbImageLimit = oRS.Collect("ImageLimit")

	'******************************************************************************
	'���^ start
	'------------------------------------------------------------------------------
	'�N��
	dbYearlyIncomeMin = ChkStr(oRS.Collect("YearlyIncomeMin"))
	dbYearlyIncomeMax = ChkStr(oRS.Collect("YearlyIncomeMax"))
	If dbYearlyIncomeMin = "0" Then dbYearlyIncomeMin = ""
	If dbYearlyIncomeMax = "0" Then dbYearlyIncomeMax = ""
	If dbYearlyIncomeMin <> "" Then dbYearlyIncomeMin = GetJapaneseYen(dbYearlyIncomeMin)
	If dbYearlyIncomeMax <> "" Then dbYearlyIncomeMax = GetJapaneseYen(dbYearlyIncomeMax)
	If dbYearlyIncomeMin & dbYearlyIncomeMax <> "" Then
		If dbYearlyIncomeMin <> "" Then sYearlyIncome = sYearlyIncome & dbYearlyIncomeMin
		sYearlyIncome = sYearlyIncome & "&nbsp;�`&nbsp;"
		If dbYearlyIncomeMax <> "" Then sYearlyIncome = sYearlyIncome & dbYearlyIncomeMax
	End If
	'����
	dbMonthlyIncomeMin = ChkStr(oRS.Collect("MonthlyIncomeMin"))
	dbMonthlyIncomeMax = ChkStr(oRS.Collect("MonthlyIncomeMax"))
	If dbMonthlyIncomeMin = "0" Then dbMonthlyIncomeMin = ""
	If dbMonthlyIncomeMax = "0" Then dbMonthlyIncomeMax = ""
	If dbMonthlyIncomeMin <> "" Then dbMonthlyIncomeMin = GetJapaneseYen(dbMonthlyIncomeMin)
	If dbMonthlyIncomeMax <> "" Then dbMonthlyIncomeMax = GetJapaneseYen(dbMonthlyIncomeMax)
	If dbMonthlyIncomeMin & dbMonthlyIncomeMax <> "" Then
		If dbMonthlyIncomeMin <> "" Then sMonthlyIncome = sMonthlyIncome & dbMonthlyIncomeMin
		sMonthlyIncome = sMonthlyIncome & "&nbsp;�`&nbsp;"
		If dbMonthlyIncomeMax <> "" Then sMonthlyIncome = sMonthlyIncome & dbMonthlyIncomeMax
	End If
	'����
	dbDailyIncomeMin = ChkStr(oRS.Collect("DailyIncomeMin"))
	dbDailyIncomeMax = ChkStr(oRS.Collect("DailyIncomeMax"))
	If dbDailyIncomeMin = "0" Then dbDailyIncomeMin = ""
	If dbDailyIncomeMax = "0" Then dbDailyIncomeMax = ""
	If dbDailyIncomeMin <> "" Then dbDailyIncomeMin = GetJapaneseYen(dbDailyIncomeMin)
	If dbDailyIncomeMax <> "" Then dbDailyIncomeMax = GetJapaneseYen(dbDailyIncomeMax)
	If dbDailyIncomeMin & dbDailyIncomeMax <> "" Then
		If dbDailyIncomeMin <> "" Then sDailyIncome = sDailyIncome & dbDailyIncomeMin
		sDailyIncome = sDailyIncome & "&nbsp;�`&nbsp;"
		If dbDailyIncomeMax <> "" Then sDailyIncome = sDailyIncome & dbDailyIncomeMax
	End If
	'����
	dbHourlyIncomeMin = ChkStr(oRS.Collect("HourlyIncomeMin"))
	dbHourlyIncomeMax = ChkStr(oRS.Collect("HourlyIncomeMax"))
	If dbHourlyIncomeMin = "0" Then dbHourlyIncomeMin = ""
	If dbHourlyIncomeMax = "0" Then dbHourlyIncomeMax = ""
	If dbHourlyIncomeMin <> "" Then dbHourlyIncomeMin = GetJapaneseYen(dbHourlyIncomeMin)
	If dbHourlyIncomeMax <> "" Then dbHourlyIncomeMax = GetJapaneseYen(dbHourlyIncomeMax)
	If dbHourlyIncomeMin & dbHourlyIncomeMax <> "" Then
		If dbHourlyIncomeMin <> "" Then sHourlyIncome = sHourlyIncome & dbHourlyIncomeMin
		sHourlyIncome = sHourlyIncome & "&nbsp;�`&nbsp;"
		If dbHourlyIncomeMax <> "" Then sHourlyIncome = sHourlyIncome & dbHourlyIncomeMax
	End If
	'------------------------------------------------------------------------------
	'���^ end
	'******************************************************************************

	'******************************************************************************
	'�Ζ��`�ԃA�C�R�� start
	'------------------------------------------------------------------------------
	sWorkingTypeIcon = ""
	sSQL = "sp_GetListWorkingType '" & dbOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		Select Case ChkStr(oRS.Collect("WorkingTypeCode"))
			Case "001": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/haken.gif"" alt=""�h��"" style=""margin-right:1px;"">"
			Case "002": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/seishain.gif"" alt=""���Ј�"" style=""margin-right:1px;"">"
			Case "003": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/keiyaku.gif"" alt=""�_��Ј�"" style=""margin-right:1px;"">"
			Case "004": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/syoha.gif"" alt=""�Љ�\��h��"" style=""margin-right:1px;"">"
			Case "005": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/arbeit.gif"" alt=""�A���o�C�g�E�p�[�g"" style=""margin-right:1px;"">"
			Case "006": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/soho.gif"" alt=""SOHO"" style=""margin-right:1px;"">"
			Case "007": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/fc.gif"" alt=""FC"" style=""margin-right:1px;"">"
		End Select
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
	'------------------------------------------------------------------------------
	'�Ζ��`�ԃA�C�R�� end
	'******************************************************************************

	'******************************************************************************
	'�摜 start
	'------------------------------------------------------------------------------
	sImg = ""
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If sImg = "" And ChkStr(oRS.Collect("PicNo1")) <> "" Then sImg = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			If sImg = "" And ChkStr(oRS.Collect("PicNo2")) <> "" Then sImg = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo2")
			If sImg = "" And ChkStr(oRS.Collect("PicNo3")) <> "" Then sImg = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo3")
			If sImg = "" And ChkStr(oRS.Collect("PicNo4")) <> "" Then sImg = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo4")
		End If
		Call RSClose(oRS)
	ElseIf dbImageLimit > 1 Then
		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If sImg = "" And ChkStr(oRS.Collect("OptionNo1")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo1")
			If sImg = "" And ChkStr(oRS.Collect("OptionNo2")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo2")
			If sImg = "" And ChkStr(oRS.Collect("OptionNo3")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo3")
			If sImg = "" And ChkStr(oRS.Collect("OptionNo4")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo4")
		End If
	End If

	If sImg = "" Then sImg = "/img/no%20image.png"
	'sImg = "<img src=""" & sImg & """ alt=""" & dbCompanyName & """ width=""156"" height=""117"">"
	sImg = "<img src=""" & sImg & """ alt=""" & dbCompanyName & """ width=""88"" height=""66"" border=""0"" align=""left"" style=""margin:0px; padding:0px;"">"
	'------------------------------------------------------------------------------
	'�摜 end
	'******************************************************************************

	'******************************************************************************
	'�Ζ��n start
	'------------------------------------------------------------------------------
	idx = 0
	sWorkingPlace = ""
	sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True And idx < 3
		dbWorkingPlacePrefectureCode = ChkStr(oRS.Collect("WorkingPlacePrefectureCode"))
		dbWorkingPlacePrefectureName = ChkStr(oRS.Collect("WorkingPlacePrefectureName"))
		dbWorkingPlaceCity = ChkStr(oRS.Collect("WorkingPlaceCity"))

		'<�Ζ��n>
		If sWorkingPlace <> "" Then sWorkingPlace = sWorkingPlace & "/"
		sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
		'</�Ζ��n>

		oRS.MoveNext
		idx = idx + 1
	Loop
	Call RSClose(oRS)
	'------------------------------------------------------------------------------
	'�Ŋ�w end
	'******************************************************************************

	rJobTypeDetail = "<a href=""" & sURL & "?ordercode=" & dbOrderCode & "&amp;rcmd=" & vRCMD & """>" & sViewJobTypeDetail & "</a>"
	rCompanyName = dbCompanyName
	rImg = "<a href=""" & sURL & "?ordercode=" & dbOrderCode & "&amp;rcmd=" & vRCMD & """>" & sImg & "</a>"
	rWorkingTypeIcon = sWorkingTypeIcon
	rWorkingPlace = sWorkingPlace
	rYearlyIncome = sYearlyIncome
	rMonthlyIncome = sMonthlyIncome
	rDailyIncome = sDailyIncome
	rHourlyIncome = sHourlyIncome
End Function

'******************************************************************************
'�T�@�v�F���Ћ��l�[�̌f�ڏ�Ԃ�ύX����
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FvOrderCodes	�F�X�V�Ώۂ̏��R�[�h�Q�i�J���}��؂�j
'�@�@�@�FvPublicFlags	�F�X�V�Ώۂ̌��J�t���O�Q�i�J���}��؂�j
'�쐬�ҁFLis Kokubo
'�쐬���F2007/04/02
'���@�l�F
'�g�p���F�����ƃi�r/order/order_list_entity.asp
'******************************************************************************
Function UpdMyOrderPublicFlag(ByRef rDB, ByVal vOrderCodes, ByVal vPublicFlags)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim aOrderCode
	Dim aPublicFlag
	Dim idx

	flgQE = True
	aOrderCode = Split(Replace(vOrderCodes, " ", ""), ",")
	aPublicFlag = Split(Replace(vPublicFlags, " ", ""), ",")

	sSQL = ""
	For idx = LBound(aOrderCode) To UBOund(aOrderCode)
		If aPublicFlag(idx) <> "" Then
			sSQL = sSQL & "EXEC sp_Reg_PublicFlag" & _
				" '" & G_USERID & "'" & _
				",'" & aOrderCode(idx) & "'" & _
				",'" & aPublicFlag(idx) & "'" & vbCrLf
		End If
	Next
	If sSQL <> "" Then flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	UpdMyOrderPublicFlag = flgQE
End Function

'******************************************************************************
'�T�@�v�F���Ћ��l�[���폜����
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FvOrderCodes	�F�X�V�Ώۂ̏��R�[�h�Q�i�J���}��؂�j
'�쐬�ҁFLis Kokubo
'�쐬���F2007/04/02
'���@�l�F
'�g�p���F�����ƃi�r/order/order_list_entity.asp
'******************************************************************************
Function DelMyOrder(ByRef rDB, vOrderCodes)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim aOrderCode
	Dim idx

	aOrderCode = Split(Replace(vOrderCodes, " ", ""), ",")
	For idx = LBound(aOrderCode) To UBound(aOrderCode)
		If aOrderCode(idx) <> "" Then
			sSQL = sSQL & "EXEC sp_Reg_RegistCommit" & _
				" '" & Replace(aOrderCode(idx), " ", "") & "'" & vbCrLf & _
				",'0'"
		End If
	Next
	If sSQL <> "" Then flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
End Function

'******************************************************************************
'�T�@�v�F���l�[�̓���
'���@���FrDB
'�@�@�@�FrRS
'�߂�l�F
'���@�l�F
'���@���F2007/02/14 LIS K.Kokubo �쐬
'�@�@�@�F2008/05/08 LIS K.Kokubo �����ǉ�(�V�[�N���b�g���l)
'�@�@�@�F2008/08/19 LIS M.Hayashi �����ǉ�
'�@�@�@�F2008/10/20 LIS K.Kokubo �Ζ��n�������ɂ��C��
'�@�@�@�F2009/03/18 LIS K.Kokubo �����ǉ�(�i�r�������Ή�)
'******************************************************************************
Function GetImgOrderSpeciality(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode
	Dim dbWorkingPlacePrefectureCode
	Dim dbWorkingPlacePrefectureName

	Dim sHTML
	Dim sWorkingCode

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")

	sHTML = ""
	'�A�N�Z�X����100�𒴂��Ă���΁uHOT�v�\���i���X�����j
	If rRS.Collect("AccessCount") > 100 Then sHTML = sHTML & "<img src=""/img/c_HOT_green.gif"" alt=""�l�C"" width=""50"" height=""15"">&nbsp;"
	'UPDATE�ƍ�������10�����������Łu�V���v�\��(���X����)
	If rRS.Collect("Updateday") > NOW()-10 Then sHTML = sHTML & "<img src=""/img/c_NEW_green.gif"" alt=""�V��"" width=""50"" height=""15"">&nbsp;"
	'���o���҂n�j�̏ꍇ�A�킩�΃}�[�N�\��(���X����)
	If rRS.Collect("InexperiencedPersonFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp01=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/no_experience.gif"" alt=""���o���Ҋ��}"" width=""50"" height=""15""></a>&nbsp;"
	'�t�^�[���E�h�^�[��
	If rRS.Collect("UITurnFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp04=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/ui_turn.gif"" alt=""�t�^�[���E�h�^�[��"" width=""50"" height=""15""></a>&nbsp;"
	'��w���������d��
	If rRS.Collect("UtilizeLanguageFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp02=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/linguistic_job.gif"" alt=""��w���������d��"" width=""50"" height=""15""></a>&nbsp;"
	'�N�ԋx��120���ȏ�
	If rRS.Collect("ManyHolidayFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp05=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/year_holidaycnt.gif"" alt=""�N�ԋx��120���ȏ�"" width=""50"" height=""15""></a>&nbsp;"
	'2006/01/10 M.Hayashi ADD �t���b�N�X�^�C�����x����
	If rRS.Collect("FlexTimeFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp06=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_flextime.gif"" alt=""�t���b�N�X�^�C�����x����"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NearStationFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp07=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_nearstation.gif"" alt=""�w��(�k��5���ȓ�)"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NoSmokingFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp08=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_nosmoking.gif"" alt=""�։��E����"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NewlyBuiltFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp09=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_newlybuilt.gif"" alt=""�V�z�r���E�I�t�B�X(5�N�ȓ�)"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("LandmarkFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp10=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_landmark.gif"" alt=""���w(15�K�ȏ�)�r��"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("RenovationFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp11=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_renovation.gif"" alt=""���m�x�[�V�����r���E�I�t�B�X(5�N�ȓ�)"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("DesignersFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp12=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_designers.gif"" alt=""�f�U�C�i�[�Y�r���E�I�t�B�X"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("CompanyCafeteriaFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp13=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_companycafeteria.gif"" alt=""�Ј��H��"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("ShortOvertimeFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp14=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_shortovertime.gif"" alt=""�c��10h/���ȓ�"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("MaternityFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp15=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_maternity.gif"" alt=""�Y�x�E��x���т���"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("DressFreeFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp16=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_dressfree.gif"" alt=""�������R"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("MammyFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp17=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_mammy.gif"" alt=""�q��ă}�}���}"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("FixedTimeFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp18=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_fixedtime.gif"" alt=""18���܂łɑގ�"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("ShortTimeFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp19=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_shorttime.gif"" alt=""1��6���Ԉȓ��J��"" width=""50"" height=""15""></a>&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("HandicappedFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp20=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_handicapped.gif"" alt=""��Q�Ҋ��}"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("RentAllFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp21=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_rentallflag.gif"" alt=""�Z���p�S�z�⏕����"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("RentPartFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp22=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_rentpartflag.gif"" alt=""�Z���p�ꕔ�⏕����"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MealsFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp23=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_mealsflag.gif"" alt=""�H���E�d���t���Č�"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MealsAssistanceFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp24=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_mealsassistanceflag.gif"" alt=""�H���⏕���x����"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("TrainingCostFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp25=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_trainingcostflag.gif"" alt=""���C������x����"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("EntrepreneurCostFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp26=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_entrepreneurcostflag.gif"" alt=""�N�Ƌ@�ޕ⏕���x����"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MoneyFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp27=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_moneyflag.gif"" alt=""�����q�E�ᗘ�q�⏕���x����"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("LandShopFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp28=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_landshopflag.gif"" alt=""�y�n�E�X�ܓ��񋟐��x����"" width=""50"" height=""15""></a>&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("FindJobFestiveFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp29=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_findjobfestiveflag.gif"" alt=""�A�E���j�������x����"" width=""50"" height=""15""></a>&nbsp;"
	'2009/12/01 LIS K.Kokubo ADD 
	If rRS.Collect("AppointmentFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp30=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_appointmentflag.gif"" alt=""���Ј��o�p���x����"" width=""50"" height=""15""></a>&nbsp;"
	'2009/12/01 LIS K.Kokubo ADD 
	If rRS.Collect("SocietyInsuranceFlag") = "1" Then sHTML = sHTML & "<a href=""/order/order_list.asp?sdf=1&sjtbig1=&sjt1=&sjtbig2=&sjt2=&sct=&swt1=&swt2=&swt3=&ssp31=1&sstc=&sgy=&syimin=&smimin=&sdimin=&shimin=&sppf=&swsh=&swsm=&sweh=&swem=&swht=&sat=&slg1=&slc1=&sl1=&skw=&skwflg=2&soc=""><img src=""/img/order_detail_icon/oc_societyinsuranceflag.gif"" alt=""�Еۊ���"" width=""50"" height=""15""></a>&nbsp;"
	'2008/05/08 LIS K.Kokubo ADD �V�[�N���b�g���l
	If rRS.Collect("SecretFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order/secret.gif"" alt=""�X�J�E�g���󂯂��l�������{���ł��鋁�l���"" width=""50"" height=""15"">&nbsp;"

	'����Yahoo!�̌������炨�d�����ڍ׃y�[�W�֗���l�փA�C�R���\��
	If InStr(Request.ServerVariables("HTTP_REFERER"),"search.yahoo.co.jp/") <> 0 Then
		sSQL = "sp_GetDataWorkingType '" & dbOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			sWorkingcode = oRS.Collect("WorkingTypecode")

			sHTML = sHTML & "<img src=""/img/order_detail_icon/icon_w" & sWorkingcode & ".gif"" alt=""�h���Ј�"" width=""50"" height=""15"">&nbsp;"

			oRS.MoveNext
		Loop
		Call RSClose(oRS)

		'<�Ζ��n>
		sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			dbWorkingPlacePrefectureCode = ChkStr(oRS.Collect("WorkingPlacePrefectureCode"))
			dbWorkingPlacePrefectureName = ChkStr(oRS.Collect("WorkingPlacePrefectureName"))
			If InStr(sHTML, "/icon_p" & dbWorkingPlacePrefectureCode & ".gif") = 0 Then
				'�����s���{���A�C�R���͏o���Ȃ��I
				sHTML = sHTML & "<img src=""/img/order_detail_icon/icon_p" & dbWorkingPlacePrefectureCode & ".gif"" alt=""" & dbWorkingPlacePrefectureName & """ width=""50"" height=""15"">&nbsp;"
			End If
		End If
		Call RSClose(oRS)
		'</�Ζ��n>
	End If

	GetImgOrderSpeciality = sHTML
End Function

'******************************************************************************
'�T�@�v�F�����ƃi�r�̋��l�[�ڍ׃y�[�W�̏㕔�ɒu���A���O�C���U���{�^���B
'���@���FvOrderCode	�F���O�C����̔�ѐ���R�[�h
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/20
'�߂�l�F�~
'�g�p���F�����ƃi�r/order/order_detail.asp
'���@�l�F
'******************************************************************************
Sub DspTopRegButton(ByVal vOrderCode)
%>
<div id="top_reg_button">

	<a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_reg_button1.gif" alt="�������o�^���ĉ���" border="0"></a>
	<a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_reg_button3.gif" alt="���O�C�����ĉ���" border="0"></a>
	
</div>
<%
End Sub

'******************************************************************************
'�T�@�v�F���������̋��l�[�ڍ׃y�[�W�̏㕔�ɒu���A���O�C���U���{�^���B
'���@���FvOrderCode	�F���O�C����̔�ѐ���R�[�h
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/20
'�߂�l�F�~
'�g�p���F�����ƃi�r/resume/order/order_detail.asp
'���@�l�F
'******************************************************************************
Sub DspTopRegButtonResume(ByVal vOrderCode)
%>
<div align="right" style="width:600px; margin-bottom:5px;">
	<div style="float:right; width:150px;"><a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_reg_button3.gif" alt="���O�C�����ĉ���" border="0"></a></div>
	<div style="float:right; width:150px;"><a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_reg_button1.gif" alt="�������o�^���ĉ���" border="0"></a></div>
	<div style="clear:both;"></div>
</div>
<%
End Sub

'******************************************************************************
'�T�@�v�F�����ƃi�r�̋��l�[�ڍ׃y�[�W�̉����ɒu���A���O�C���U���{�^���B
'���@���FvOrderCode	�F���O�C����̔�ѐ���R�[�h
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/20
'�߂�l�F�~
'�g�p���F�����ƃi�r/order/order_detail.asp
'���@�l�F
'******************************************************************************
Sub DspBottomRegButton(ByVal vOrderCode)
%>
<div class="center">
	<p>
��������o�^����Ή���⎿�₪�\�ɂȂ�܂��I����<BR>
����̂��߂̗������������쐬����܂��B</p>

	<div class="center left"><a href="<%= HTTPS_NAVI_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="�������o�^���ĉ���" border="0"></a></div>
	<div class="center right"><a href="<%= HTTPS_NAVI_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="���O�C�����ĉ���" border="0"></a></div>
	<br style="clear:both;">

</div>

<!--�VSNS�{�^��-->
<div id="sns_button" class="smartNone">
<!-- #INCLUDE FILE="../include/social_bookmark.asp" -->
<div class="right">
<!--G+-->
<!-- +1 �{�^�� ��\���������ʒu�Ɏ��̃^�O��\��t���Ă��������B -->
<div class="g-plusone" data-size="tall" data-annotation="none"></div>

<!-- �Ō�� +1 �{�^�� �^�O�̌�Ɏ��̃^�O��\��t���Ă��������B -->
<script type="text/javascript">
  window.___gcfg = {lang: 'ja'};

  (function() {
    var po = document.createElement('script'); po.type = 'text/javascript'; po.async = true;
    po.src = 'https://apis.google.com/js/plusone.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(po, s);
  })();
</script>
</div>

<!--facebook-->

<div class="fb-like" data-href="http://<%= Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") %>" data-send="false" data-layout="button_count" data-width="110" data-show-faces="false"></div>

<!--<script type="text/javascript">
    var url = encodeURIComponent(location.href);
    document.write('<iframe src="http://www.facebook.com/plugins/like.php?href=' + url + '&width=100&layout=button_count&show_faces=false&action=like&colorscheme=light&height=20" scrolling="no" frameborder="0" style="border:none; overflow:hidden;width:100px;height:20px;" allowTransparency="true"></iframe>');
</script>-->

<!--/facebook-->
<!--twitter-->
<a href="https://twitter.com/share" class="twitter-share-button" data-via="shigoto_navi" data-lang="ja" data-count="none">�c�C�[�g</a>
<script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>

</div>
<!--/�VSNS�{�^��-->
<br clear="both">

<%
End Sub

'******************************************************************************
'�T�@�v�F�����ƃi�r�̋��l�[�ڍ׃y�[�W�̉����ɒu���A���O�C���U���{�^���B
'���@���FvOrderCode	�F���O�C����̔�ѐ���R�[�h
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/20
'�߂�l�F�~
'�g�p���F�����ƃi�r/order/order_detail.asp
'���@�l�F
'******************************************************************************
Sub DspBottomRegButton_OldPlan(ByVal vOrderCode,ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vAccessCount,ByVal YearlyIncomeMin,ByVal MonthlyIncomeMin,ByVal DailyIncomeMin,ByVal HourlyIncomeMin)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode

	Dim sOptionNo			'�傫���ʐ^�̔ԍ�
	Dim sCompanyPictureFlag	'��Ǝʐ^�t���O ["1"]�L ["0"]��
	Dim sImg1
	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic

	Dim JobTypeBigCode
	Dim JobTypeCode
	Dim WorkingTypeCode1
	Dim WorkingTypeCode2
	Dim WorkingTypeCode3
	Dim PrefectureCode
	
	If GetRSState(rRS) = False Then Exit Sub

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")

	'�ٗp�`�ԁA�Ζ��n�A�E�팟��
			sSQL = "select CJT.jobtypecode, BJT.Bigclasscode from C_JobType AS CJT INNER JOIN B_JobType AS BJT ON CJT.JobTypeCode = BJT.AllConnectCode where CJT.id = '1' and CJT.OrderCode = '" & dbOrderCode & "';"
			flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("Bigclasscode")) <> "" Then
					JobTypeBigCode = oRS.Collect("Bigclasscode")
				End If
				If ChkStr(oRS.Collect("jobtypecode")) <> "" Then
					JobTypeCode = oRS.Collect("jobtypecode")
				End If
			End If
			Call RSClose(oRS)

			sSQL = "select prefecturecode from c_workingplace where ordercode = '" & dbOrderCode & "';"
			flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
			PrefectureCode  = ""
			Do While GetRSState(oRS) = True
				If ChkStr(oRS.Collect("prefecturecode")) <> "" Then
					PrefectureCode = PrefectureCode & oRS.Collect("prefecturecode") & ","
				End If
				oRS.MoveNext
			Loop
			Call RSClose(oRS)
			PrefectureCode = Left(PrefectureCode, Len(PrefectureCode) -1)

			sSQL = "select CWT1.workingtypecode as workingtypecode1,CWT2.workingtypecode as workingtypecode2,CWT3.workingtypecode as workingtypecode3 from c_workingtype AS CWT1 "
			sSQL = sSQL & " left join c_workingtype AS CWT2 on CWT1.ordercode = '" & dbOrderCode & "' and CWT2.id = 2"
			sSQL = sSQL & " left join c_workingtype AS CWT3 on CWT2.ordercode = '" & dbOrderCode & "' and CWT3.id = 3"
			sSQL = sSQL & " where CWT3.ordercode = '" & dbOrderCode & "' and CWT1.id = 1;"
			flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("workingtypecode1")) <> "" Then
					WorkingTypeCode1 = oRS.Collect("workingtypecode1")
				End If
				If ChkStr(oRS.Collect("workingtypecode2")) <> "" Then
					WorkingTypeCode2 = oRS.Collect("workingtypecode2")
				End If
				If ChkStr(oRS.Collect("workingtypecode3")) <> "" Then
					WorkingTypeCode3 = oRS.Collect("workingtypecode3")
				End If
			End If
			Call RSClose(oRS)

%>
<div class="center">
	<p>
��������o�^����Ή���⎿�₪�\�ɂȂ�܂��I����<BR>
����̂��߂̗������������쐬����܂��B</p>

	<div class="center left">            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?sdf=1&amp;sjtbig1=<%= JobTypeBigCode %>&amp;sjt1=<%= JobTypeCode %>&amp;swt1=<%= WorkingTypeCode1 %>&amp;swt2=<%= WorkingTypeCode2 %>&amp;swt3=<%= WorkingTypeCode3 %>&amp;spc=<%= PrefectureCode %>&amp;syimin=<%= YearlyIncomeMin %>&amp;smimin=<%= MonthlyIncomeMin %>&amp;sdimin=<%= DailyIncomeMin %>&amp;shimin=<%= HourlyIncomeMin %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/top_reg_button03.png" alt="�������o�^���ĉ���" border="0"></a></div>
	<div class="center right"><a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_list.asp?sdf=1&amp;sjtbig1=<%= JobTypeBigCode %>&amp;sjt1=<%= JobTypeCode %>&amp;swt1=<%= WorkingTypeCode1 %>&amp;swt2=<%= WorkingTypeCode2 %>&amp;swt3=<%= WorkingTypeCode3 %>&amp;spc=<%= PrefectureCode %>&amp;syimin=<%= YearlyIncomeMin %>&amp;smimin=<%= MonthlyIncomeMin %>&amp;sdimin=<%= DailyIncomeMin %>&amp;shimin=<%= HourlyIncomeMin %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/top_login_button03.png" alt="���O�C�����ĉ���" border="0"></a></div>
	<br style="clear:both;">

</div>

<!--�VSNS�{�^��-->
<div id="sns_button" class="smartNone">
<!-- #INCLUDE FILE="../include/social_bookmark.asp" -->
<div class="right">
<!--G+-->
<!-- +1 �{�^�� ��\���������ʒu�Ɏ��̃^�O��\��t���Ă��������B -->
<div class="g-plusone" data-size="tall" data-annotation="none"></div>

<!-- �Ō�� +1 �{�^�� �^�O�̌�Ɏ��̃^�O��\��t���Ă��������B -->
<script type="text/javascript">
  window.___gcfg = {lang: 'ja'};

  (function() {
    var po = document.createElement('script'); po.type = 'text/javascript'; po.async = true;
    po.src = 'https://apis.google.com/js/plusone.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(po, s);
  })();
</script>
</div>

<!--facebook-->

<div class="fb-like" data-href="http://<%= Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") %>" data-send="false" data-layout="button_count" data-width="110" data-show-faces="false"></div>

<!--<script type="text/javascript">
    var url = encodeURIComponent(location.href);
    document.write('<iframe src="http://www.facebook.com/plugins/like.php?href=' + url + '&width=100&layout=button_count&show_faces=false&action=like&colorscheme=light&height=20" scrolling="no" frameborder="0" style="border:none; overflow:hidden;width:100px;height:20px;" allowTransparency="true"></iframe>');
</script>-->

<!--/facebook-->
<!--twitter-->
<a href="https://twitter.com/share" class="twitter-share-button" data-via="shigoto_navi" data-lang="ja" data-count="none">�c�C�[�g</a>
<script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>

</div>
<!--/�VSNS�{�^��-->
<br clear="both">

<%
End Sub

'******************************************************************************
'�T�@�v�F���������̋��l�[�ڍ׃y�[�W�̉����ɒu���A���O�C���U���{�^���B
'���@���FvOrderCode	�F���O�C����̔�ѐ���R�[�h
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/20
'�߂�l�F�~
'�g�p���F�����ƃi�r/resume/order/order_detail.asp
'���@�l�F
'******************************************************************************
Sub DspBottomRegButtonResume(ByVal vOrderCode)
%>
<div align="center">
	<hr size="1">
	<p style="color:#ff0000;">������o�^����Ή���⎿�₪�\�ɂȂ�܂��I��</p>
	<hr size="1">
	<div align="center" style="float:left; width:300px;color:#C51035;">���܂�ID���������łȂ�����<br><a href="<%= HTTPS_NAVI_CURRENTURL %>resume/staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_reg_button1.gif" alt="�������o�^���ĉ���" border="0"></a></div>
	<div align="center" style="float:right; width:300px;color:#C51035;">�����ł�ID���������̕���<br><a href="<%= HTTPS_NAVI_CURRENTURL %>resume/login/login.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/resume/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="<%= HTTP_NAVI_CURRENTURL %>img/order/btn_reg_button3.gif" alt="���O�C�����ĉ���" border="0"></a></div>
	<div style="clear:both;"></div>
	<br>
</div>

<%
End Sub

'******************************************************************************
'�T�@�v�F�V�����l��񃁁[������A�N�Z�X���������ꍇ�̃��O��������
'���@���FrDB		
'�@�@�@�FrRS		
'�@�@�@�FvMU		�F�����}�K���[�U�h�c
'�@�@�@�FvOrderCode	�F�{�������l�[
'�쐬�ҁFLis Kokubo
'�쐬���F2007/05/08
'�߂�l�F
'�@�@�@�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'���@�l�F
'******************************************************************************
Function MailMagazineAccess(ByRef rDB, ByVal vMU, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	If IsNumber(vMU, 0, False) = True Then
		sSQL = "up_Reg_LOG_MailMagazineAccess '" & vMU & "', '" & vOrderCode & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Call RSClose(oRS)
	End If
End Function

'******************************************************************************
'�T�@�v�F���l�����}�K����A�N�Z�X���������ꍇ�̃��O��������
'���@���FrDB		
'�@�@�@�FrRS		
'�@�@�@�FvMU		�F�����}�K���[�U�h�c
'�@�@�@�FvOrderCode	�F�{�������l�[
'�쐬�ҁFLis Kokubo
'�쐬���F2007/05/08
'�߂�l�F
'�@�@�@�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'���@�l�F
'******************************************************************************
Function MailMagazineDelivery(ByRef rDB, ByVal vMI, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	If IsNumber(vMI, 0, False) = True Then
		sSQL = "up_Reg_LOG_MailMagazineDelivery '" & vMI & "', '" & vOrderCode & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Call RSClose(oRS)
	End If
End Function

'******************************************************************************
'�T�@�v�F���Ճ��O�̏�������
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_SearchOrder or ���l�[�ڍ׌���SQL �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvOrderCode		�F�{�������l�[
'���@�l�F
'�g�p���Forder/order_detail.asp
'���@���F2007/05/08 LIS K.Kokubo �쐬
'�@�@�@�F2009/05/19 LIS K.Kokubo �Г�����̃A�N�Z�X��S0018066�̃A�N�Z�X�̓��O�Ɏc���Ȃ�
'�@�@�@�F2009/06/01 LIS.T.Ezaki  �p�����[�^�[�iuc�j�ɃX�^�b�t�R�[�h���L�ڂ���΃��O�ɋL�^����
'******************************************************************************
Function AccessHistoryOrder(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	'�Г�����̃A�N�Z�X�ƁA�����낤����(S0018066)����̃A�N�Z�X�̓��O�Ɏc���Ȃ�
	If IsRE(G_IPADDRESS, "^192.168.", True) = False And vUserID <> "S0018066" Then
		If vUserType = "staff" Then
			sSQL = "up_Reg_LOG_AccessHistoryOrder '" & vOrderCode & "', '" & vUserID & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		ElseIf IsRE(Request.Cookies("id_memory"), "^S\d\d\d\d\d\d\d$", True) = True Then
			sSQL = "up_Reg_LOG_AccessHistoryOrder '" & vOrderCode & "', '" & Request.Cookies("id_memory") & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		ElseIf IsRE(GetForm("uc",2), "^S\d\d\d\d\d\d\d$", True) = True Then
			sSQL = "up_Reg_LOG_AccessHistoryOrder '" & vOrderCode & "', '" & GetForm("uc",2) & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
			sSQL = "update P_Userinfo set lastaccessday = getdate() where staffcode = '" & GetForm("uc",2) & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		End If
	End If
End Function

'******************************************************************************
'�T�@�v�F���Ճ��O�̏�������(�V�������}�K)
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_SearchOrder or ���l�[�ڍ׌���SQL �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvOrderCode		�F�{�������l�[
'���@�l�F
'�g�p���Forder/order_detail.asp
'���@���F2014/09/11 LIS TANIZAWA �쐬
'******************************************************************************
Function AccessHistoryOrderNEW(ByRef rDB, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	'�Г�����̃A�N�Z�X�ƁA�����낤����(S0018066)����̃A�N�Z�X�̓��O�Ɏc���Ȃ�
	If IsRE(G_IPADDRESS, "^192.168.", True) = False And vUserID <> "S0018066" Then
		If vUserType = "staff" Then
			sSQL = "up_Reg_LOG_NewOrderMailAccessHistory '" & vOrderCode & "', '" & vUserID & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		ElseIf IsRE(Request.Cookies("id_memory"), "^S\d\d\d\d\d\d\d$", True) = True Then
			sSQL = "up_Reg_LOG_NewOrderMailAccessHistory '" & vOrderCode & "', '" & Request.Cookies("id_memory") & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		End If
	End If
End Function

'******************************************************************************
'�T�@�v�F�A�N�Z�X�񐔂̃J�E���g�A�b�v
'���@���FrDB		�F�ڑ�����DBConnection
'�@�@�@�FvOrderCode	�F�{�������l�[�̏��R�[�h
'�쐬�ҁFLis Kokubo
'�쐬���F2007/05/08
'���@�l�F
'�g�p���Forder/order_detail_entity.asp
'******************************************************************************
Function AccessCountUp(ByRef rDB, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	AccessCountUp = 0

	sSQL = "sp_Reg_AccessCountUp '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		AccessCountUp = oRS.Collect("AccessCount")
	End If
	Call RSClose(oRS)
End Function

'******************************************************************************
'�T�@�v�F���l�[�̓��ʂo�u�̃J�E���g�A�b�v
'���@���FrDB		�F�ڑ�����DBConnection
'�@�@�@�FvOrderCode	�F�{�������l�[�̏��R�[�h
'���@�l�F
'�g�@�p�Forder/order_detail.asp
'���@���F2008/05/23 LIS K.Kokubo �쐬
'******************************************************************************
Function PVCountUp(ByRef rDB, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	sSQL = "up_RegC_PV '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	Call RSClose(oRS)
End Function

'*******************************************************************************
'�T�@�v�F�S�p���p����������������̃o�C�g���𐳊m�ɕԂ�(Web����̈��p)
'���@���Fstring		:�Ώە�����
'�߂�l�FInterger	:�Ώە�����̃o�C�g��
'�쐬���F2007/05/23 Lis Sotome
'���@���F
'*******************************************************************************
Function LenByte(ByRef string)

    Dim c, i, k

    c = 0

    For i = 0 To Len(string) - 1
        k = Mid(string, i + 1, 1)
        If (Asc(k) And &HFF00) = 0 Then
            c = c + 1
        Else
            c = c + 2
        End If
    Next

    LenByte = c

End Function

'*******************************************************************************
'�T�@�v�F������̍��[����w�肳�ꂽ�o�C�g�����̕�����𒊏o����(�S�p���p�̍�������������Ή�)
'�@�@�@�F���w�肳�ꂽ�o�C�g���Ŏ��܂�Ȃ��S�p�����͍���܂�
'�@�@�@�Fex:sStr="aa��", vByte=3 �E�E�E�߂�l:"aa"
'���@���FsStr		:�Ώە�����
'      �FvByte		:���o���镶����̃o�C�g��
'�߂�l�FString		:���o��̕�����
'�쐬���F2007/05/23 Lis Sotome
'���@���F
'*******************************************************************************
Function LeftByte(ByRef sStr, ByRef vByte)

    Dim cnt, i, k
	Dim sBuf	'������p�o�b�t�@

    cnt = 0

    For i = 0 To Len(sStr) - 1
        k = Mid(sStr, i + 1, 1)
        If (Asc(k) And &HFF00) = 0 Then
            cnt = cnt + 1
        Else
            cnt = cnt + 2
        End If

		If cnt > vByte Then	'�ړI�̕������𒴂���(���p�A�S�p�Ƒ�����)�Ƃ�
			LeftByte = sBuf
			Exit Function	'�����I��
		Elseif cnt = vByte Then	'�ړI�̕�������(���p�A���p�܂��͑S�p�A�S�p�Ƒ�����)�Ƃ�
			sBuf = sBuf & k
			LeftByte = sBuf
			Exit Function	'�����I��
		Elseif cnt < vByte Then
			sBuf = sBuf & k
		End If
	Next

	LeftByte = sBuf

End Function
%>
