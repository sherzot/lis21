<%
'**********************************************************************************************************************
'�T�@�v�F���l�[�ꗗ�y�[�W /order/order_list_entity.asp
'�@�@�@�F���l�[�ڍ׃y�[�W /order/order_detail_entity.asp
'�@�@�@�F��Ə��y�[�W /order/company_order.asp
'�@�@�@�F��L�y�[�W�ŏo�͗p�̊֐��Q�����̃t�@�C���ɗp�ӂ���B
'�@�@�@�F
'�@�@�@�F�������@�O������@������
'�@�@�@�F�v���O�C���N���[�h
'�@�@�@�F/config/personel.asp
'�@�@�@�F/include/commonfunc.asp
'��@���F�������@���l�[�ꗗ�y�[�W�o�͗p�@������
'�@�@�@�FDspOrderListDetail			�F���l�[�ꗗ�y�[�W�̊e���l�[�P�ʂ�\��
'�@�@�@�FDspOrderListDetail2		�F���l�[�ꗗ�����уo�[�W����1
'�@�@�@�FDspOrderListDetail3		�F���l�[�ꗗ�����уo�[�W����2
'�@�@�@�FDspPageControl				�F���l�[�ꗗ�y�[�W�̃y�[�W�R���g���[��
'�@�@�@�F
'�@�@�@�F�������@��Ə��y�[�W�o�͗p�@������
'�@�@�@�FDspCompanyInfo				�F��Ə��̊�{�����o��
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

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")

	DspOrderListDetail = False

	If G_USEFLAG = "0" And vMyOrder = "1" And G_OLDAPPLICATIONCODE <> "" Then
		sSQL = "EXEC up_DtlOrder '" & rRS.Collect("OrderCode") & "', '" & G_OLDAPPLICATIONCODE & "';"
	Else
		sSQL = "EXEC up_DtlOrder '" & rRS.Collect("OrderCode") & "', '';"
	End If

	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	sOrderType = ChkStr(oRS.Collect("OrderType"))
	sPlanType = ChkStr(oRS.Collect("PlanTypeName"))
	iImageLimit = oRS.Collect("ImageLimit")

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
	If iImageLimit > 0 Then
		sImgMain = ""
		sImgSub = ""
		sCompanyPictureFlag = ChkStr(oRS.Collect("CompanyPictureFlag"))

		sSQL = "up_GetListOrderPictureNow '" & oRS.Collect("CompanyCode") & "', '" & oRS.Collect("OrderCode") & "', 'orderpicture'"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then
			If ChkStr(oRS2.Collect("OptionNo1")) <> "" Or (sOrderType = "0" And sCompanyPictureFlag = "1") Then
				sImgMain = "<img src=""/company/imgdsp.asp?companycode=" & oRS2.Collect("CompanyCode") & "&amp;optionno=" & oRS2.Collect("OptionNo1") & """ alt="""" border=""0"" width=""" & PICSIZEW & """ height=""" & PICSIZEH & """>"
				flgImg = True
			End If

			If sPlanType = "platinum" Or sPlanType = "old" Then
				If ChkStr(oRS2.Collect("OptionNo2")) <> "" Then
					sImgSub = sImgSub & "<div align=""center"" style=""float:left; width:80px;"">" & _
						"<img src=""/company/imgdsp.asp?companycode=" & oRS2.Collect("CompanyCode") & "&amp;optionno=" & oRS2.Collect("OptionNo2") & """ alt=""" & oRS2.Collect("Caption2") & """ border=""1"" width=""" & PICSIZESUBW & """ height=""" & PICSIZESUBH & """ style=""border:1px solid #666666;""><br>"
					sImgSub = sImgSub & "</div>"
					flgImg = True
				End If

				If ChkStr(oRS2.Collect("OptionNo3")) <> "" Then
					sImgSub = sImgSub & "<div align=""center"" style=""float:left; width:80px;"">" & _
						"<img src=""/company/imgdsp.asp?companycode=" & oRS2.Collect("CompanyCode") & "&amp;optionno=" & oRS2.Collect("OptionNo3") & """ alt=""" & oRS2.Collect("Caption3") & """ border=""1"" width=""" & PICSIZESUBW & """ height=""" & PICSIZESUBH & """ style=""border:1px solid #666666;""><br>"
					sImgSub = sImgSub & "</div>"
					flgImg = True
				End If

				If ChkStr(oRS2.Collect("OptionNo4")) <> "" Then
					sImgSub = sImgSub & "<div align=""center"" style=""float:left; width:80px;"">" & _
						"<img src=""/company/imgdsp.asp?companycode=" & oRS2.Collect("CompanyCode") & "&amp;optionno=" & oRS2.Collect("OptionNo4") & """ alt=""" & oRS2.Collect("Caption4") & """ border=""1"" width=""" & PICSIZESUBW & """ height=""" & PICSIZESUBH & """ style=""border:1px solid #666666;""><br>"
					sImgSub = sImgSub & "</div>"
					flgImg = True
				End If
				If sImgSub <> "" Then sImgSub = "<div style=""padding-top:1px;"">" & sImgSub & "<div style=""clear:both;""></div></div>"
			End If
		Else
			If sCompanyPictureFlag = "1" And sOrderType = "0" Then
				sImgMain = "<img src=""/company/imgdsp.asp?companycode=" & oRS2.Collect("CompanyCode") & "&amp;optionno=1"" alt="""" border=""0"" width=""" & PICSIZEW & """ height=""" & PICSIZEH & """>"
				flgImg = True
			End If
		End If

		Call RSClose(oRS2)
	End If
	'--------------------------------------------------------------------------
	'�摜 end
	'**************************************************************************

	'**************************************************************************
	'�S���Ɩ� start
	'--------------------------------------------------------------------------
	If flgImg = True Then
		'�摜���L��ꍇ�͕��͂�Z�߂ɃJ�b�g
		sBusinessDetail = Left(oRS.Collect("BusinessDetail"),100) & "&nbsp;"
		If Len(sBusinessDetail) > 100 Then sBusinessDetail = sBusinessDetail & "..."
	Else
		'�摜�������ꍇ�͕��͂𒷂߂ɃJ�b�g
		sBusinessDetail = Left(oRS.Collect("BusinessDetail"),155) & "&nbsp;"
		If Len(sBusinessDetail) > 155 Then sBusinessDetail = sBusinessDetail & "..."
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
	Do While GetRSState(oRS2) = True
		sWorkingType = sWorkingType & oRS2.Collect("WorkingTypeName")
		If (oRS.Collect("OrderType") ="0" And oRS.Collect("Companykbn") = "2") Or oRS.Collect("OrderType") ="2" Then
			Select Case oRS2.Collect("WorkingTypeCode")
				Case "001": sWorkingType = sWorkingType & "�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/koyoukeitai_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>�h���Ƃ�</a>�z" 
				Case "002","003": sWorkingType = sWorkingType & "�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>�l�ޏЉ�Ƃ�</a>�z" 
				Case "004": sWorkingType = sWorkingType & "�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/syoukaiyotei_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>�Љ�\��h���Ƃ�</a>�z" 
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
	Do While GetRSState(oRS2) = True And idx < 3
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
	If oRS2.RecordCount > 3 Then sWorkingPlace = sWorkingPlace & "&nbsp;...��"
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
	Response.Write "<table border=""0"" class=""old"">"
	Response.Write "<tbody>"
	Response.Write "<tr>"
	Response.Write "<td class=""old11"" style=""padding-left:0px; width:600px;"" valign=""middle"">"

	If vUserType = "" Or vUserType = "staff" Then
		'�񃍃O�C�����A�X�^�b�t���O�C����

		'�E���l�[�t�q�k�����[�����M
		'�E�E�H�b�`���X�g�֕ۑ�
		Response.Write "<div style=""float:left;width:420px;"">"
		Response.Write "<img src=""/img/list_companyicon.gif"" alt="""" align=""left"">" & sTitleCompanyName
		Response.Write "<h3 style=""margin-left:5px;"">��<a href=""" & HTTP_CURRENTURL & "order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """>" & sTitleJobName & "</a>" & sImgMail & "</h3>"
		Response.Write "</div>"
		Response.Write "<div align=""right"" style=""float:right;font-size:11px;width:113px;"">"
		Response.Write "<a href=""" & HTTPS_CURRENTURL & "order/sendmail_jobofferaddress.asp?OrderCode=" & oRS.Collect("OrderCode") & """ onclick=""window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;""><img src=""/img/order/ordermail.gif"" style=""margin-bottom:6px;"" border=""0"" alt=""���l�������[�����M"" align=""top""></a>"
		Response.Write "<a href=""" & HTTPS_CURRENTURL & "order/sendmail_jobofferaddress.asp?OrderCode=" & oRS.Collect("OrderCode") & """ onclick=""window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;""><img src=""/img/order/orderwachlist.gif"" border=""0"" alt=""�E�H�b�`���X�g�ɒǉ�"" align=""top""></a>"
		Response.Write "</div>"
		Response.Write "<div style=""clear:both;""></div>"
	ElseIf vUserType = "company" Then
		'��ƃ��O�C����
		Response.Write "<p class=""m0""><img src=""/img/list_companyicon.gif"" alt="""" align=""left"">" & sTitleCompanyName & "</p>"
		Response.Write "<h3 style=""margin-left:5px;"">��<a href=""../order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """>" & sTitleJobName & "</a>" & sImgMail & "</h3>"
	End If

	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""old12"">"
	'**TOP 08/08/19 Lis�� REP
	'Response.Write "<div style=""float:left;"">" & sImgOrderState & "</div>"
	'Response.Write "<div align=""right"" style=""font-size:10px;line-height:14px;"">�f�ڊ����F" & sPublishLimitStr & "</div>"
	'Response.Write "<div style=""clear:both;""></div>"
	Response.Write "<table style='width:600px;'><tr><td style='width:500px;padding-left:5px;'>" & sImgOrderState & "</td>"
	Response.Write "<td style='width:100px;vertical-align:top;font-size:10px;text-align:right;'>�f�ڊ����F"
	Response.Write sPublishLimitStr & "</td></tr></table>"
	'**BTM 08/08/19 Lis�� REP
	Response.Write "<table border=""0"" class=""old2"">"

	If sCatchCopy <> "" Then
		Response.Write "<caption>" & sCatchCopy & "</caption>"
	End If

	Response.Write "<tbody>"
	Response.Write "<tr>"
	Response.Write "<td rowspan=""2"" valign=""top"">"

	If flgImg = True Then
		'�摜���L��ꍇ�̃��C�A�E�g
		Response.Write "<div class=""old21"" style=""margin:0px 12px;"">"
		Response.Write "<b>�y�S���Ɩ��̐����z</b><br>" & sBusinessDetail
		Response.Write "</div>"
		Response.Write "<div class=""old21"" style=""width:240px; float:left; margin:0px 5px;"">"
		Response.Write "<a href=""" & HTTP_NAVI_CURRENTURL & "order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """ title=""" & sTitleCompanyName & """>" & sImgMain & "</a>"
		Response.Write sImgSub
		Response.Write "</div>"
	Else
		'�摜�������ꍇ�̃��C�A�E�g
		Response.Write "<div class=""old21"" style=""width:239px; float:left; margin:0px 5px;"">"
		Response.Write "<b>�y�S���Ɩ��̐����z</b><br>" & sBusinessDetail
		Response.Write "</div><br>"
	End If

	Response.Write "<table style=""width:330px; margin-left:3px;"">"
	Response.Write "<tr>"
	Response.Write "<td style=""font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;"">"
	Response.Write "�Ζ��`��"
	Response.Write "</td>"
	Response.Write "<td style=""background-color:#eeeeee; padding:5px 0px 5px 10px; border-bottom:solid 3px #ffffff;"">"
	Response.Write sWorkingType
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td style=""font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;"">"
	Response.Write "�Ζ��n"
	Response.Write "</td>"
	Response.Write "<td style=""background-color:#eeeeee; padding-left:10px; border-bottom:solid 3px #ffffff;"">"
	Response.Write sWorkingPlace & "" & sStationName
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td style=""font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;"">"
	Response.Write "���^"
	Response.Write "</td>"
	Response.Write "<td style=""background-color:#eeeeee; padding:5px 0px 5px 10px; border-bottom:solid 3px #ffffff;"">"

	If sYearlyIncome <> "" Then
		Response.Write "<p>�N��&nbsp;" & sYearlyIncome & "</p>"
	End If

	If sMonthlyIncome <> "" Then
		Response.Write "<p>����&nbsp;" & sMonthlyIncome & "</p>"
	End If

	If sDailyIncome <> "" Then
		Response.Write "<p>����&nbsp;" & sDailyIncome & "</p>"
	End If

	If sHourlyIncome <> "" Then
		Response.Write "<p>����&nbsp;" & sHourlyIncome & "</p>"
	End If

	Response.Write "</td>"
	Response.Write "</tr>"

	If sBizName1 <> "" Then

		Response.Write "<tr>"
		Response.Write "<td style=""font-weight:bold; background-color:#E1FBCD; width:70px; border-bottom:solid 3px #ffffff; text-align:center;"">"
		Response.Write "�d���̊���"
		Response.Write "</td>"
		Response.Write "<td style=""background-color:#eeeeee; border-bottom:solid 3px #ffffff; padding-left:0px; line-height:14px;"">"
		Response.Write "<table>"
		Response.Write "<tr>"
		Response.Write "<td style=""padding:5px 0px 5px 7px;"">"
		Response.Write "<script type=""text/javascript"" language=""javascript"">"
		Response.Write "viewWorkAvg(" & sBizPercentage1 & ", " & sBizPercentage2 & ", " & sBizPercentage3 & ", " & sBizPercentage4 & ")"
		Response.Write "</script>"
		Response.Write "</td>"
		Response.Write "<td>"

		If sBizName1 <> "" Then Response.Write "<p style=""font-size:10px; line-height:12px;""><span style=""color:#ff9999;"">��</span>" & sBizPercentage1 & "%�@" & sBizName1 & "</p>"
		If sBizName2 <> "" Then Response.Write "<p style=""font-size:10px; line-height:12px;""><span style=""color:#9999ff;"">��</span>" & sBizPercentage2 & "%�@" & sBizName2 & "</p>"
		If sBizName3 <> "" Then Response.Write "<p style=""font-size:10px; line-height:12px;""><span style=""color:#99ff99;"">��</span>" & sBizPercentage3 & "%�@" & sBizName3 & "</p>"
		If sBizName4 <> "" Then Response.Write "<p style=""font-size:10px; line-height:12px;""><span style=""color:#ffff99;"">��</span>" & sBizPercentage4 & "%�@" & sBizName4 & "</p>"

		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</table>"
		Response.Write "</td>"
		Response.Write "</tr>"
	End If

	Response.Write "</table>"
	Response.Write "<div align=""right"" style=""margin:3px 5px;"">"

	If dbWValueURL <> "" Then
		Response.Write "<a href=""" & dbWValueURL & """ target=""_blank""><img src=""/img/order/btn_wvalue.gif"" border=""0"" alt=""���l���:" & sTitleCompanyName & "�̎��Ѝ̗p�y�[�W""></a>"
	End If

	If dbTopInterviewFlag = "1" Then
		Response.Write "<a href=""" & HTTP_CURRENTURL & "order/order_interview.asp?ordercode=" & dbOrderCode & """><img src=""/img/order/interview_icon.gif"" border=""0"" alt=""���l���:�g�b�v�C���^�r���[""></a>"
	End If

	Response.Write "<a href=""" & HTTP_CURRENTURL & "order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """><img src=""/img/detail_button2.gif"" border=""0"" alt=""���l���:�ڍ�""></a>"
	Response.Write "</div>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "<div style=""clear:both;""></div>"
	Response.Write "</td>"
	Response.Write "</tr>"

	If oRS.Collect("CompanyCode") = vUserID And vMyOrder = "1" And G_USEFLAG = "1" Then
		Response.Write "<tr>"
		Response.Write "<td class=""old13"">"
		Response.Write "<table class=""old3"">"
		Response.Write "<tbody>"
		Response.Write "<tr>"
		Response.Write "<td class=""old31"">���R�[�h(" & oRS.Collect("OrderCode") & ")</td>"
		Response.Write "<td class=""old32"">���</td>"
		Response.Write "<td class=""old33"">"
		Response.Write sProgress
		Response.Write "<select name=""CONF_PublicFlags"" " & sPublicListDsp & ">"
		If oRS.Collect("PublicFlag") = "1" Then
			Response.Write "<option value=""1"" selected>�f��</option>"
			Response.Write "<option value=""0"">��f��</option>"
		Else
			Response.Write "<option value=""1"">�f��</option>"
			Response.Write "<option value=""0"" selected>��f��</option>"
		End If
		Response.Write "</select>"
		Response.Write "</td>"
		Response.Write "<td class=""old34"">�f�ړ�<br>�o�^��</td>"
		Response.Write "<td class=""old35"">" & sPublicDay & "<br>" & sRegistDay & "</td>"
		Response.Write "<td class=""old36""><input type=""checkbox"" name=""CONF_DeleteFlags"" value=""" & oRS.Collect("OrderCode") & """>�폜</td>"
		Response.Write "</tr>"
		Response.Write "</tbody>"
		Response.Write "</table>"
		Response.Write "</td>"
		Response.Write "</tr>"
	End If

	Response.Write "<tr>"
	Response.Write "<td class=""old14""></td>"
	Response.Write "</tr>"
	Response.Write "</table>"

	DspOrderListDetail = True
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

	sURL = HTTP_CURRENTURL & "order/order_detail.asp"

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

			Response.Write "<div style=""float:left; width:200px;""><div style=""line-height:16px; " & aPadding(iCols) & """>"

			Response.Write aImg(idx)
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
			Response.Write "</div></div>"
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
<table style="width:600px; margin:10px 0px;">
	<tbody>
	<tr>
		<td style="width:88px; padding:5px; border-width:1px 0px 1px 1px; text-align:center;">
<%
	If vPage > 1 Then Response.Write "<a href='javascript:ChgPage(" & vPage - 1 & ");'>�O�̃y�[�W</a>"
%>
		</td>
		<td style="width:489px; padding:5px; border-width:1px 0px 1px 0px; text-align:center;">
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
	If sCompanyPictureFlag = "1" And dbImageLimit > 0 Then
	Response.Write "<div style=""width:302px; float:right;""><img id=""imgcompany"" src=""" & HTTPS_NAVI_CURRENTURL & "company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=1"" alt=""�C���[�W�ʐ^"" width=""300"" height=""225"" style=""border:1px solid #999999;""></div>"
	Response.Write "<div style=""float:left; width:295px;"">"
	End If

	If sCompanyCode <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
		Response.Write "<div class=""category""><h4>��ƃR�[�h</h4></div>"
		Response.Write "<div class=""" & sClass & """><p class=""m0"">" & sCompanyCode & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>"
	End If

	If sEstablishYear <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
		Response.Write "<div class=""category""><h4>�ݗ��N�x</h4></div>"
		Response.Write "<div class=""" & sClass & """><p class=""m0"">" & sEstablishYear & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>"
	End If

	If sCapitalAmount <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
		Response.Write "<div class=""category""><h4>���{�z</h4></div>"
		Response.Write "<div class=""" & sClass & """><p class=""m0"">" & sCapitalAmount & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>"
	End If

	If sListClass <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
		Response.Write "<div class=""category""><h4>�������J</h4></div>"
		Response.Write "<div class=""" & sClass & """><p class=""m0"">" & sListClass & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>"
	End If

	If sEmployeeNum <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
		Response.Write "<div class=""category""><h4>�Ј���</h4></div>"
		Response.Write "<div class=""" & sClass & """><p class=""m0"">" & sEmployeeNum & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>"
	End If

	If sIndustryType <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
		Response.Write "<div class=""category""><h4>�Ǝ�</h4></div>"
		Response.Write "<div class=""" & sClass & """><p class=""m0"">" & sIndustryType & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>"
	End If

	If sAddress <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
		Response.Write "<div class=""category""><h4>�{�ЏZ��</h4></div>"
		Response.Write "<div class=""" & sClass & """><p class=""m0"">" & sAddress & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>"
	End If

	If sNearbyStation <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
		Response.Write "<div class=""category""><h4>�{�ЍŊ�w</h4></div>"
		Response.Write "<div class=""" & sClass & """><p class=""m0"">" & sNearbyStation & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>"
	End If

	If sHomePage <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
		Response.Write "<div class=""category""><h4>�z�[���y�[�W</h4></div>"
		Response.Write "<div class=""" & sClass & """><p class=""m0""><a href=""" & sHomePage & """ target=""_blank"">���̊�Ƃ̃z�[���y�[�W</a></p></div>"
		Response.Write "<div style=""clear:both;""></div>"
	End If

	If sCompanyPictureFlag = "1" And dbImageLimit > 0 Then
	Response.Write "</div>"
	Response.Write "<div style=""clear:both;""></div>"
	End If
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

	'******************************************************************************
	'���Ɠ��e start
	'------------------------------------------------------------------------------
	sBusiness = ""
	sBusiness = Replace(ChkStr(rRS.Collect("BusinessContents")), vbCrLf, "<br>")
	sBusiness = Replace(sBusiness, vbCr, "<br>")
	sBusiness = Replace(sBusiness, vbLf, "<br>")
	'------------------------------------------------------------------------------
	'���Ɠ��e end
	'******************************************************************************

	'******************************************************************************
	'��ЏЉ� start
	'------------------------------------------------------------------------------
	sPR = ""
	sPR = Replace(ChkStr(rRS.Collect("CompanyPR")), vbCrLf, "<br>")
	sPR = Replace(sPR, vbCr, "<br>")
	sPR = Replace(sPR, vbLf, "<br>")
	'------------------------------------------------------------------------------
	'��ЏЉ� end
	'******************************************************************************

	'******************************************************************************
	'�������� start
	'------------------------------------------------------------------------------
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

	sWelfareProgramRemark = ChkStr(rRS.Collect("WelfareProgramRemark"))
	'------------------------------------------------------------------------------
	'�������� end
	'******************************************************************************

	flgPR = False
	If sBusiness & sPR & sWelfare <> "" Then flgPR = True

	flgLine = False
	sClass = "value2"

	If flgPR = True Then
		Response.Write "<div class=""companyblock"">"
		Response.Write "<h3>" & sAddTitle & "�o�q���</h3>"
		If sBusiness <> "" Then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
			Response.Write "<div class=""category""><h4>���Ɠ��e</h4></div>"
			Response.Write "<div class=""" & sClass & """><p class=""m0"">" & sBusiness & "</p></div>"
			Response.Write "<div style=""clear:both;""></div>"
		End If

		If sPR <> "" Then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
			Response.Write "<div class=""category""><h4>��Ђo�q</h4></div>"
			Response.Write "<div class=""" & sClass & """><p class=""m0"">" & sPR & "</p></div>"
			Response.Write "<div style=""clear:both;""></div>"
		End If

		If sWelfare <> "" Then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
			Response.Write "<div class=""category""><h4>��������</h4></div>"
			Response.Write "<div class=""" & sClass & """>" & sWelfare & sWelfareProgramRemark & "</div>"
			Response.Write "<div style=""clear:both;""></div>"
		End If
		Response.Write "</div>"
		Response.Write "<br>"
	End If
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
			sIntrDisp = "�Љ��"
		Else
			sImgTitle = "/img/order/lisorderpr1.gif"
			sIntrDisp = "�h����"
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
<h3><%= sIntrDisp %>��Ə��</h3>
<%
		If sListClass <> "" Then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
%>
<div class="category1"><h4>�������J</h4></div>
<div class="value1"><p class="m0"><%= sListClass %></p></div>
<div style="clear:both;"></div>
<%
		End If

		If sIndustryType <> "" Then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
%>
<div class="category1"><h4>�Ǝ�</h4></div>
<div class="value1"><p class="m0"><%= sIndustryType %></p></div>
<div style="clear:both;"></div>
<%
		End If


		If sPR <> "" Then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
			

%>
<div class="category1"><h4>���Ɠ��e</h4></div>
<div class="value1"><p class="m0"><%= sPR %></p></div>
<div style="clear:both;"></div>
<%		End If
		'**TOP 08/08/21 Lis�� ADD
		If sCapitalAmount <> "" Then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
%>
<div class="category1"><h4>���{��</h4></div>
<div class="value1"><p class="m0"><%= sCapitalAmount %></p></div>
<div style="clear:both;"></div>
<%		End If
		If sEmployeeNum <> "" Then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
%>
<div class="category1"><h4>�Ј���</h4></div>
<div class="value1"><p class="m0"><%= sEmployeeNum %></p></div>
<div style="clear:both;"></div>
<%		End If
		sflgAct = ""
		If InStr(sImportantNotice,"����J") <= 0 and _
		((sAccountingPeriod1 <> "" and sSalesAmount1 <> "" and InStr(sAccountingPeriod1 & sSalesAmount1,"����J") <= 0) or _
		 (sAccountingPeriod2 <> "" and sSalesAmount2 <> "" and InStr(sAccountingPeriod2 & sSalesAmount2,"����J") <= 0) or _
		 (sAccountingPeriod3 <> "" and sSalesAmount3 <> "" and InStr(sAccountingPeriod3 & sSalesAmount3,"����J") <= 0)) then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
%>
<div class="category1"><h4>�������</h4></div>
<div class="value1"><p class="m0">
<%			'���㍂�P�E�o�험�v�P�E���Z���P
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
%>
</p></div>
<div style="clear:both;"></div>
<%		End If
%><p class="m0" style="font-size:10px;margin:0px 20px;color:red;">
���l��<%= left(sIntrDisp,2) %>�ł��ē����邨�d���̂��߁A�ڂ�����Џ��͉��̃{�^���₨�d�b�ȂǂŒ��ڂ��⍇�����������B</p>
<%		response.write "<p>�@</p>"
		'**BTM 08/08/21 Lis�� ADD
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

	If vMyOrder = "1" Then
		'******************************************************************************
		'���Ћ��l�[�̏ꍇ start
		'------------------------------------------------------------------------------
		Response.Write "<h2 class=""csubtitle"">���Ћ��l�[�̑���</h2>"
		Response.Write "<div class=""subcontent"">"

		'�����{�^��
		Response.Write "<p class=""cctrltitle"">���E�Ҍ����E�X�J�E�g���[��</p>"
		Response.Write "<div style=""padding-left:15px;"">"
		Response.Write "<div style=""padding-top:5px;"">"
		Response.Write "<form action=""/staff/person_list.asp"" method=""get"" style=""display:inline;"">"
		Response.Write "<input name=""ordercode"" type=""hidden"" value=""" & sOrderCode & """>"
		Response.Write "<input type=""submit"" value=""���E�҂���������"" style=""width:150px; color:#aa3300;"">"
		Response.Write "</form>"
'		Response.Write "<input type=""button"" value=""���E�҂���������"" style=""width:150px; color:#aa3300;"" onclick=""Go_Edit('10');"">"
		Response.Write "<span style=""font-size:10px; color:#666666;"">�E�E�E���̋��l�[�́A�E��E�Ζ��n�E�ٗp�`�Ԃ𖞂������E�҂��������܂��B</span>"
		Response.Write "</div>"
		Response.Write "<div style=""padding-top:5px;"">"
		Response.Write "<form action=""/staff/person_search_detail.asp"" method=""get"" style=""display:inline;"">"
		Response.Write "<input name=""ordercode"" type=""hidden"" value=""" & sOrderCode & """>"
		Response.Write "<input name=""setdata"" type=""hidden"" value=""1"">"
		Response.Write "<input type=""submit"" value=""���E�҂��ڍ׌���"" style=""width:150px; color:#aa3300;"">"
		Response.Write "</form>"
		Response.Write "<span style=""font-size:10px; color:#666666;"">�E�E�E���̋��l�[����A�ڍׂȌ����������w�肵�ċ��E�҂��������܂��B</span><br>"
		Response.Write "</div>"
		If G_USEFLAG = "0" Then
			Response.Write "<p style=""padding-top:5px; color:#ff0000; font-size:10px;"">�����݃��C�Z���X���؂�Ă��邽�߁A�X�J�E�g�A���l�[�̕ҏW�͂ł��܂���B</p>"
		ElseIf G_PUBLICFLAG = "0" Then
			Response.Write "<p style=""padding-top:5px; color:#ff0000; font-size:10px;"">�����݋��l�[�̌f�ڊ��ԊO�̂��߁A�X�J�E�g�͂ł��܂���B</p>"
		End If
		Response.Write "</div>" & vbCrLf
		'/�����{�^��

		If sHakouFlag = "1" Then
			Response.Write "<br>"

			'���l�[�R�s�[�쐬
			Response.Write "<p class=""cctrltitle"">���l�[�R�s�[�쐬</p>" & vbCrLf
			Response.Write "<div style=""padding:5px 0px;"">"
			Response.Write "<div style=""padding:0px 0px 5px 15px;"">"
			Response.Write "<input type=""button"" value=""���l�[���R�s�["" style=""width:100px; color:#3333ff;"" onclick=""if(confirm('���̋��l�[���R�s�[���āA�V�������l�[���쐬���܂����H')){location.href='" & HTTPS_CURRENTURL & vUserType & "/orderedit/new.asp?copy=" & sOrderCode & "';}"">"
			Response.Write "<span style=""font-size:10px; color:#666666;"">�E�E�E���̋��l�[�����ƂɁA�V�������l�[���쐬���܂��B</span><br>"
			Response.Write "</div>"
			Response.Write "</div>"

			Response.Write "<p class=""cctrltitle"">���l����ҏW����</p>"
			Response.Write "<div style=""padding:5px 0px;"">"
			Response.Write "<div style=""padding:0px 0px 5px 15px;"">"
			Response.Write "<div style=""float:left; width:290px;"">"
			Response.Write "<input type=""button"" value=""���Џ��X�V"" style=""width:100px;"" onclick=""location.href='" & HTTPS_CURRENTURL & vUserType & "/company_reg1.asp';"">"
			Response.Write "<span style=""font-size:10px; color:#666666;"">�E�E�E���Џ����X�V���܂��B</span>"
			Response.Write "</div>"
			Response.Write "<div style=""float:right; width:290px;"">"
			Response.Write "<input type=""button"" value=""��W���ҏW"" style=""width:100px;"" onclick=""location.href='" & HTTPS_CURRENTURL & vUserType & "/orderedit/base.asp?ordercode=" & sOrderCode & "';"">"
			Response.Write "<span style=""font-size:10px; color:#666666;"">�E�E�E��W����ҏW���܂��B</span>"
			Response.Write "</div>"
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "</div>" & vbCrLf

			If G_INTERVIEWFLAG = "1" Then
				Response.Write "<div style=""padding:0px 0px 5px 15px;"">"
				Response.Write "<div style=""float:left; width:290px;"">"
				Response.Write "<form action=""/company/topinterview/reg.asp"" method=""get"" style=""display:inline;"">"
				Response.Write "<input type=""submit"" value=""�g�b�v�C���^�r���["" style=""width:100px;"">"
				Response.Write "</form>"
				Response.Write "<span style=""font-size:10px; color:#666666;"">�E�E�E�g�b�v�C���^�r���[��ҏW���܂��B</span>"
				Response.Write "</div>"
				Response.Write "<div style=""float:right; width:290px;"">"
				Response.Write "<form action=""/company/elderinterview/list.asp"" method=""get"" style=""display:inline;"">"
				Response.Write "<input name=""ordercode"" type=""hidden"" value=""" & sOrderCode & """>"
				Response.Write "<input type=""submit"" value=""��y�C���^�r���["" style=""width:100px;"">"
				Response.Write "</form>"
				Response.Write "<span style=""font-size:10px; color:#666666;"">�E�E�E��y�C���^�r���[��ҏW���܂��B</span>"
				Response.Write "</div>"
				Response.Write "<div style=""clear:both;""></div>"
				Response.Write "</div>"
			End If

			Response.Write "</div>"

			Response.Write "<p class=""cctrltitle"">���[���e���v���[�g</p>"
			Response.Write "<div style=""padding:5px 0px;"">"
			Response.Write "<div style=""padding:0px 0px 5px 15px;"">"

			If iMailTemplateCnt >= 5 Then
				'���[���e���v���[�g��������ɒB���Ă���ꍇ�͐V�K�쐬�ł��Ȃ�
				Response.Write "<p style=""color:#ff0000; font-size:10px;"">���[���e���v���[�g��������ɒB���Ă���̂ŁA����ȏ�쐬�ł��܂���B</p>"
			Else
				'���[���e���v���[�g��������ɒB���Ă��Ȃ��ꍇ�͐V�K�쐬�ł���
				Response.Write "<input type=""button"" value=""�V�K�쐬"" style=""width:100px;"" onclick=""location.href='" & HTTPS_NAVI_CURRENTURL & "mailtemplate/regist.asp?ordercode=" & sOrderCode & "';"">"
				Response.Write "<span style=""font-size:10px; color:#666666;"">�E�E�E���̋��l�̃��[���e���v���[�g��V�K�ɍ쐬���܂��B</span><br>"
			End If

			Response.Write "<p style=""color:#ff0000; font-size:10px;"">�����[���e���v���[�g�͋��l�[���ɍ쐬���܂��B</p>"

			sSQL = "up_GetListMailTemplate '" & G_USERID & "', '" & sOrderCode & "'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then Response.Write "<hr size=""1"">"
			Do While GetRSState(oRS) = True
				sAncMT = "?ordercode=" & oRS.Collect("OrderCode") & "&amp;seq=" & oRS.Collect("SEQ")
				sAncMT = "<a href=""" & HTTPS_NAVI_CURRENTURL & "mailtemplate/regist.asp" & sAncMT & """>" & oRS.Collect("Subject") & "</a>"

				Response.Write "<div style=""width:585px;"">"
				Response.Write "<div style=""float:left; width:85px;"">" & GetDetail("MailTemplateType", oRS.Collect("MailTemplateTypeCode")) & "</div>"
				Response.Write "<div style=""float:left; width:500px;"">" & sAncMT & "</div>"
				Response.Write "<div style=""clear:both;""></div>"
				Response.Write "</div>"

				oRS.MoveNext
			Loop

			Response.Write "</div>"
			Response.Write "</div>"
		End If

		Response.Write "</div>"
		'------------------------------------------------------------------------------
		'���Ћ��l�[�̏ꍇ end
		'******************************************************************************
	ElseIf vUserType = "staff" Then
		'******************************************************************************
		'���O�C�����E�҂̏ꍇ start
		'------------------------------------------------------------------------------
		If rRS.Collect("PublicFlag") = "1" Then
			Response.Write "<div class=""subcontent"" style=""margin-bottom:15px;"">"
			Response.Write "<div style=""padding:5px 0px;"">"
			Response.Write "<p class=""sctrltitle"">����E����E�E�H�b�`���X�g</p>"
			Response.Write "<div style=""padding:0px 0px 5px 15px;"">"
			Response.Write "<div style=""float:left; width:195px;"">"
			Response.Write "<p class=""m0"" style=""margin-right:20px; font-size:10px; color:#666666; text-align:center;"">�����̕�W�։��僁�[���̍쐬</p>"
			Response.Write "<input type=""button"" value=""���僁�[���𑗐M����"" style=""width:180px;"" onclick=""contactCompany('');"">"
			Response.Write "</div>"
			Response.Write "<div align=""center"" style=""float:left; width:195px;"">"
			Response.Write "<p class=""m0"" style=""font-size:10px; color:#666666; text-align:center;"">�����̕�W�֎��⃁�[���̍쐬</p>"
			Response.Write "<input type=""button"" value=""���⃁�[���𑗐M����"" onclick=""contactCompany('1');"">"
			Response.Write "</div>"
			Response.Write "<div style=""float:left; width:195px;"">"
			Response.Write "<p class=""m0"" style=""margin-left:20px; font-size:10px; color:#666666; text-align:center;"">��<a href=""watchlist_info.htm"" onclick=""window.open(this.href, 'mywindow6', 'width=300, height=150, menubar=no, toolbar=no, scrollbars=yes'); return false;"" style=""color:#0045F9;"">�E�H�b�`���X�g</A>�֒ǉ�</p>"

			If flgAddWatchList = True Then
				Response.Write "<p class=""m0"" style=""margin-left:20px; text-align:center; font-weight:bold;"">���ɓo�^�ς݂ł�</p>"
			Else
				Response.Write "<div align=""right""><input type=""button"" value=""���̋��l�[��ǉ�����"" style=""width:180px;"" onclick=""document.forms.frmMain.action='/staff/watchlist_register.asp';document.forms.frmMain.submit();""></div>"
			End If
			Response.Write "</div>"
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "</div>"
			Response.Write "</div>"
			Response.Write "</div>"
		Else
			Response.Write "<div align=""center""><b>���̋��l�[�͌f�ڂ��I�����Ă��܂��B���[�����M�͂ł��܂���B</b></div>"
		End If
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
	MailWin = window.open('<%= HTTPS_NAVI_CURRENTURL %>staff/mailtocompany.asp?' + sQ + 'ordercode=<%= sOrderCode %>','mail','width=480,height=580,resizable=1,scrollbars=no');
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
	sCautionStr = "<p class=""m0"" style=""line-height:11px;text-align:right;font-size:11px;"">�������O�Ɍf�ڏI������ꍇ������܂��B</p>"

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
<div style="width:600px; margin-bottom:10px;">
<%
	'���X�Љ�Č�,�l�މ�ЏЉ�Č��̏ꍇ�́u�]�E���k�Č��v�C���[�W��\��
	If sOrderType = "2" Or (sCompanyKbn = "2" And dbTempOrderFlag = "0" And dbTTPOrderFlag = "0") Then
		Response.Write "<img src=""/img/order/counselable_order.gif"" width=""150"" height=""25"" alt=""�]�E�x�����󂯂ĉ��傷�鋁�l�ł�"">"
	End If
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
			<img src="/ImgQRCode.asp?Code=<%= rRS.Collect("OrderCode") %>" alt="QRCode">
		</div>
		<div style="text-align:right; font-size:11px; padding-top:6px;">
			<a href="<%= HTTPS_NAVI_CURRENTURL %>order/sendmail_jobofferaddress.asp?OrderCode=<% = rRS.Collect("OrderCode") %>&amp;detail=1" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=470');return false;"><img src="/img/staff/mail/mailhei.gif" border="0" align="bottom" alt="���l�[�����[�����M"> ���l�[�����[�����M</a>
		</div>
		<p class="m0" style="text-align:right;padding:4px 0px 0px 0px;">�f�ڊ����F<%= sPublishLimitStr %></p>
		<div style="clear:both;"></div>
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
		<img src="/ImgQRCode.asp?Code=<%= rRS.Collect("OrderCode") %>" alt="QRCode" border="0" align="right">
		<div style="text-align:right; font-size:11px; padding-top:6px;"><a href="<%= HTTPS_NAVI_CURRENTURL %>order/sendmail_jobofferaddress.asp?OrderCode=<% = rRS.Collect("OrderCode") %>&amp;detail=1" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=380');return false;"><img src="/img/staff/mail/mailhei.gif" border="0" align="bottom" alt="���l�[�����[�����M"> ���l�[�����[�����M</a></div>
		<p class="m0" style="text-align:right;padding:4px 0px 0px 0px;">�f�ڊ����F<%= sPublishLimitStr %></p>
		<div style="clear:both;"></div>
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
		<img src="/ImgQRCode.asp?Code=<%= rRS.Collect("OrderCode") %>" alt="QRCode" border="0" align="right">
		<p class="m0" style="text-align:right; width:156px; padding-top:21px;">�f�ڊ����F<%= sPublishLimitStr %></p>
		<div style="clear:both;"></div>
		<%= sCautionStr %>
		<div style="clear:both;"></div>
	</div>
	<div style="clear:both;"></div>
<%
	End If
%>
</div>
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

	Response.Write "<div style=""width:600px; margin-bottom:5px;"">"
	Response.Write "<div style=""float:left; width:350px; margin:0px;"">"
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
	Response.Write "<div class=""clear:both; margin:0px;""></div>"
	Response.Write "</div>"
	Response.Write "<div align=""right"" style=""float:right; width:250px;"">"
	Response.Write "<p class=""m0"">���ԎQ�Ɖ񐔁F" & vAccessCount & "��@�X�V���F" & sUpdateDay & "</p>"
	Response.Write "</div>"
	Response.Write "<div style=""clear:both;""><img src=""/img/order/tab_border.gif"" alt="""" width=""600"" height=""5""></div>"
	Response.Write "</div>" & vbCrLf
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

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")

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

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)

	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"" style=""width:600px;"">"

		Response.Write "<div style=""float:right; width:302px;"">"
		Response.Write "<img class=""big"" src=""" & sImg1 & """ alt="""" border=""1"" width=""300"" height=""225"" style=""border:1px solid #999999;"">"
		Response.Write "</div>"

		Response.Write "<h2>" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"" style=""padding-top:20px;"">" & rRS.Collect("CatchCopy") & "</p><br>"
		Response.Write "<div style=""margin:10px 0px;"">"

		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;padding:5px;"">"
			Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">�y��W�̓����z</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write "</div>"
		Response.Write "<br clear=""all"">"
		Response.Write "</div>"
	Else
		Response.Write "<div id=""catchcopy"" style=""width:600px;"">"
		Response.Write "<h2 style=""width:600px;"">" & rRS.Collect("JobTypeDetail") & "</h2>"
		Response.Write "<p class=""m0"" style=""width:600px;padding-top:20px;"">" & rRS.Collect("CatchCopy") & "</p><br><br>"
		Response.Write "<div style=""margin:10px 0px;"">"

		If sImgSpeciality <> "" Then
			Response.Write "<div style=""border:solid 0px #cccccc;padding:5px;"">"
			Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">�y��W�̓����z</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
		End If

		Response.Write"</div>"
		Response.Write "<div style=""clear:both;""></div>"
		Response.Write "</div>"
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
		Response.Write "<h3>�o�q</h3>"
		Response.Write "<div class=""freeprblock"">"
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
'******************************************************************************
Function DspOrderPictureNow(ByRef rDB, ByRef rRS, ByVal vCategoryCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode
	Dim dbCompanyCode
	Dim dbImageLimit

	Dim sURL

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbCompanyCode = rRS.Collect("CompanyCode")
	dbImageLimit = rRS.Collect("ImageLimit")

	If dbImageLimit > 1 Then
		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', '" & vCategoryCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If Len(oRS.Collect("OptionNo2")) > 0 Or Len(oRS.Collect("OptionNo3")) > 0 Or Len(oRS.Collect("OptionNo4")) > 0 Then
				Response.Write "<div align=""center"" style=""padding:5px 0px 5px 15px; background-color:#e1fbcd; margin-bottom:40px;"">"
				Response.Write "<div style=""width:580px;"">"
				sURL = ""
				If Len(oRS.Collect("OptionNo2")) > 0 Then
					sURL = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo2")
					Response.Write "<div align=""right"" style=""float:left; width:190px;"">"
					Response.Write "<div style=""width:182px; background-color:#ffffff;""><img src=""" & sURL & """ alt=""" & oRS.Collect("Caption2") & """ width=""180"" height=""135"" border=""1"" style=""border:1px solid #999999;""></div>"
					Response.Write "<p class=""m0"" align=""left"" style=""width:182px; font-size:10px;"">" & oRS.Collect("Caption2") & "</p>"
					Response.Write "</div>"
				End If

				sURL = ""
				If Len(oRS.Collect("OptionNo3")) > 0 Then
					sURL = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo3")
					Response.Write "<div align=""right"" style=""float:left; width:190px;"">"
					Response.Write "<div style=""width:182px; background-color:#ffffff;""><img src=""" & sURL & """ alt=""" & oRS.Collect("Caption3") & """ width=""180"" height=""135"" border=""1"" style=""border:1px solid #999999;""></div>"
					Response.Write "<p class=""m0"" align=""left"" style=""width:182px; font-size:10px;"">" & oRS.Collect("Caption3") & "</p>"
					Response.Write "</div>"
				End If

				sURL = ""
				If Len(oRS.Collect("OptionNo4")) > 0 Then
					sURL = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo4")
					Response.Write "<div align=""right"" style=""float:left; width:190px;"">"
					Response.Write "<div style=""width:182px; background-color:#ffffff;""><img src=""" & sURL & """ alt=""" & oRS.Collect("Caption4") & """ width=""180"" height=""135"" border=""1"" style=""border:1px solid #999999;""></div>"
					Response.Write "<p class=""m0"" align=""left"" style=""width:182px; font-size:10px;"">" & oRS.Collect("Caption4") & "</p>"
					Response.Write "</div>"
				End If

				Response.Write "<br clear=""all"">"
				Response.Write "</div>"
				Response.Write "</div>"
			End If
		End If
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
		Response.Write "<h3>�̗p�̔w�i</h3>" & vbCrLf
		Response.Write "<p class=""m0"" style=""padding-left:15px;"">" & dbOrderBackGround & "</p>" & vbCrLf
		DspOrderBackGround = True
	End If

	If DspOrderBackGround = True Then Response.Write "<br>"
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
		Response.Write "<h3>�Ɩ����e</h3>"

		If sBusinessDetail <> "" Then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
			Response.Write "<div class=""category1""><h4>�S���Ɩ�</h4></div>"
			Response.Write "<div class=""value1""><p class=""m0"">" & sBusinessDetail & "</p></div>"
			Response.Write "<div style=""clear:both;""></div>"
		End If

		If (sPlanType = "platinum" Or sPlanType = "gold" Or sPlanType = "old") And sBiz <> "" Then
			If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
			flgLine = True
			Response.Write "<div class=""category1""><h4>�d���̊���</h4></div>"
			'Response.Write "<div class=""value1"">" & sBiz & "</div>"
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
			Response.Write "<div style=""clear:both;""></div>"
		End If
		Response.Write "<br>"
		Response.Write "<br>" & vbCrLf
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

	If sWorkingTime & dbWorkTimeRemark <> "" Then flgDspTime = True
	'</����>

	'<�x��>
	flgDspHoliday = False
	dbWeeklyHolidayType = ChkStr(rRS.Collect("WeeklyHolidayTypeName"))
	dbHolidayRemark = ChkStr(rRS.Collect("HolidayRemark"))

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

		'<�Ζ��n>
		sWorkingPlace = sWorkingPlace & "<div>"
		If iMaxRow > 1 Then sWorkingPlace = sWorkingPlace & "�y�Ζ��n" & dbWorkingPlaceSeq & "�z"
		If dbOrderType = "0" Then
			If dbPlanTypeName = "mail" Then
				sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
			Else
				sWorkingPlace = sWorkingPlace & dbWorkingPlaceAddressAll
				If dbWorkingPlaceSection & dbWorkingPlaceTelephoneNumber <> "" Then
					sWorkingPlace = sWorkingPlace & "("
					If dbWorkingPlaceSection <> "" Then sWorkingPlace = sWorkingPlace & dbWorkingPlaceSection
					If dbWorkingPlaceSection <> "" And dbWorkingPlaceTelephoneNumber <> "" Then sWorkingPlace = sWorkingPlace & "&nbsp;"
					If dbWorkingPlaceTelephoneNumber <> "" Then sWorkingPlace = sWorkingPlace & "TEL:" & dbWorkingPlaceTelephoneNumber
					sWorkingPlace = sWorkingPlace & ")"
				End If
				If dbMapFlag = "1" Then sWorkingPlace = sWorkingPlace & "&nbsp;[<span style=""color:#0045f9;cursor:pointer;"" onclick=""open('" & HTTP_CURRENTURL & "map/showmap.asp?ordercode=" & dbOrderCode & "&wpseq=" & dbWorkingPlaceSeq & "', 'map', 'width=700,height=650');"">�n�}</span>]"
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
				sNearbyRailway = GetNearbyRailway(rDB, oRS3)
			End If
			oRS3.Filter = 0
			'</�Ŋ񉈐�>

			If sNearbyStation <> "" Then
				sWorkingPlace = sWorkingPlace & "<p class=""m0"" style=""padding-left:15px;"">"
				sWorkingPlace = sWorkingPlace & "[�Ŋ�w]"
				sWorkingPlace = sWorkingPlace & sNearbyStation
				sWorkingPlace = sWorkingPlace & "<br>"
				sWorkingPlace = sWorkingPlace & "[����]"
				sWorkingPlace = sWorkingPlace & sNearbyRailway
				sWorkingPlace = sWorkingPlace & "</p>"
			End If
		Else
			sWorkingPlace = sWorkingPlace & dbWorkingPlacePrefectureName & dbWorkingPlaceCity
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
	sHTML = sHTML & "<h3>�Ζ�����</h3>"

	If flgDspWorkingType = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		sHTML = sHTML & "<div class=""category1""><h4>�Ζ��`��</h4></div>"
		sHTML = sHTML & "<div class=""value1"">"
		'<�Ζ��`��>
		If sWorkingType <> "" Then
			sHTML = sHTML & "<p class=""m0"">" & sWorkingType & "</p>"
		End If
		'</�Ζ��`��>
		'<�Љ��̋Ζ��`��>
		If dbTTPOrderFlag = "1" And sAfterWorkingType <> "" Then
			sHTML = sHTML & "<p class=""m0"">" & sAfterWorkingType & "</p>"
		End If
		'</�Љ��̋Ζ��`��>
		'<�A�Ɗ���>
		If sWorkRange <> "" Then
			sHTML = sHTML & "<p class=""m0"">���L���̏ꍇ�F" & sWorkRange & "</p>"
		End If
		'</�A�Ɗ���>
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
	End If

	If flgDspJobType = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		'<�E��>
		sHTML = sHTML & "<div class=""category1""><h4>�E��</h4></div>"
		sHTML = sHTML & "<div class=""value1"">"
		sHTML = sHTML & "<p class=""m0""><strong>" & dbJobTypeDetail & "</strong></p>"
		sHTML = sHTML & "<p class=""m0"">" & sJobType & "</p>"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
		'</�E��>
	End If

	If flgDspSalary = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		sHTML = sHTML & "<div class=""category1""><h4>���^</h4></div>"
		sHTML = sHTML & "<div class=""value1"">"

		If sYearlyIncome <> "" Then
			'<�N��>
			sHTML = sHTML & "<h5>�N��</h5>"
			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sYearlyIncome & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>"
			'</�N��>
		End If

		If sMonthlyIncome <> "" Then
			'<����>
			sHTML = sHTML & "<h5>����</h5>"
			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sMonthlyIncome & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>"
			'</����>
		End If

		If sDailyIncome <> "" Then
			'<����>
			sHTML = sHTML & "<h5>����</h5>"
			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sDailyIncome & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>"
			'</����>
		End If

		If sHourlyIncome <> "" Then
			'<����>
			sHTML = sHTML & "<h5>����</h5>"
			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sHourlyIncome & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>"
			'</����>
		End If

		If dbPercentagePay <> "" Then
			'<������>
			sHTML = sHTML & "<h5>������</h5>"
			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbPercentagePay & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both; margin:0px;""></div>"
			'</������>
		End If

		If sTrafficFee <> "" Then
			'<��ʔ�>
			sHTML = sHTML & "<h5>��ʔ�</h5>"
			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sTrafficFee & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>"
			'</��ʔ�>
		End If

		If dbSalaryRemark <> "" Then
			'<���^���l>
			sHTML = sHTML & "<h5>���^���l</h5>"
			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbSalaryRemark & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both; margin:0px;""></div>"
			'</���^���l>
		End If

		sHTML = sHTML & "<p class=""m0"" style=""font-size:10px;"">"
		sHTML = sHTML & "���Œ�z�͏����Ɋ֌W�Ȃ�������z�ł��B(�N���̍Œ�z�͏����Ɋ֌W�Ȃ������錎���̍��v�ł��B)"
		sHTML = sHTML & "</p>"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
	End If

	If flgDspTime = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		sHTML = sHTML & "<div class=""category1""><h4>����</h4></div>"
		sHTML = sHTML & "<div class=""value1"">"

		If sWorkingTime <> "" Then
			'<�A�Ǝ���>
			sHTML = sHTML & "<h5>�A�Ǝ���</h5>"
			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sWorkingTime & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>"
			'</�A�Ǝ���>
		End If

		If dbWorkTimeRemark <> "" Then
			'<�A�Ǝ��Ԕ��l>
			sHTML = sHTML & "<h5>�A�Ǝ��Ԕ��l</h5>"
			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbWorkTimeRemark & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>"
			'</�A�Ǝ��Ԕ��l>
		End If

		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
	End If

	If flgDspHoliday = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		sHTML = sHTML & "<div class=""category1""><h4>�x��</h4></div>"
		sHTML = sHTML & "<div class=""value1"">"

		If dbWeeklyHolidayType <> "" Then
			'<�x��>
			sHTML = sHTML & "<h5>�x��</h5>"
			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbWeeklyHolidayType & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>"
			'</�x��>
		End If

		If dbHolidayRemark <> "" Then
			'<�x�����l>
			sHTML = sHTML & "<h5>�x�����l</h5>"
			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbHolidayRemark & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>"
			'</�x�����l>
		End If

		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
	End If

	If flgDspHumanNumber = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		'<��W�l��>
		sHTML = sHTML & "<div class=""category1""><h4>��W�l��</h4></div>"
		sHTML = sHTML & "<div class=""value1"">"
		sHTML = sHTML & "<p class=""m0"">" & dbHumanNumber & "</p>"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
		'</��W�l��>
	End If

	If flgDspWorkingPlace = True Then
		If flgLine = True Then sHTML = sHTML & "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		'<�Ζ��n>
		sHTML = sHTML & "<div class=""category1""><h4>�Ζ��n</h4></div>"
		sHTML = sHTML & "<div class=""value1"">"

		If sWorkingPlace <> "" Then
			sHTML = sHTML & "<h5>�Z��</h5>"
			sHTML = sHTML & "<div class=""value2"">"
			sHTML = sHTML & "<p class=""m0"">" & sWorkingPlace & "</p>"
			If sNearbyStationBlock <> "" Then
				sHTML = sHTML & sNearbyStationBlock
			End If
			sHTML = sHTML & "</div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>"
		End If

'<��ֈ�>
'		If sWorkingPlace <> "" Then
'			sHTML = sHTML & "<h5>�Ζ��n</h5>"
'			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sWorkingPlace & "</p></div>"
'			sHTML = sHTML & "<div style=""clear:both;""></div>"
'		End If

'		If sNearbyStation <> "" Then
'			sHTML = sHTML & "<h5>�Ŋ�w</h5>"
'			sHTML = sHTML & "<div class=""value2"">" & sNearbyStation & "</div>"
'			sHTML = sHTML & "<div style=""clear:both;""></div>"
'		End If

'		If sNearbyRailway <> "" Then
'			sHTML = sHTML & "<h5>����</h5>"
'			sHTML = sHTML & "<div class=""value2"">" & sNearbyRailway & "</div>"
'			sHTML = sHTML & "<div style=""clear:both;""></div>"
'		End If
'</��ֈ�>

		If dbTransfer <> "" Then
			sHTML = sHTML & "<h5>�]��</h5>"
			sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbTransfer & "</p></div>"
			sHTML = sHTML & "<div style=""clear:both;""></div>"
		End If

		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
		'</�Ζ��n>
	End If

	sHTML = sHTML & "<br>"

	Response.Write sHTML
	'DspCondition = sHTML
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
	If sOrderType = "2" Or sOrderType = "3" Then
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
		sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[�H]�����ɂ���</a>"
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
	If sFEHistory <> "" Then sFEHistory = sFEHistory & "���ȏ�"
	If sFEHistory <> "" Then DspNeedCondition = True
	'------------------------------------------------------------------------------
	'�w�� end
	'******************************************************************************

	'******************************************************************************
	'���i start
	'------------------------------------------------------------------------------
	sLicense = GetLicense(rDB, rRS)
	sLicenseOther = GetOrderNote(rDB, rRS, "OtherLicense")
	flgLicense = False
	If sLicense & sLicenseOther <> "" Then
		flgLicense = True
		DspNeedCondition = True
	End If
	'------------------------------------------------------------------------------
	'���i end
	'******************************************************************************

	'******************************************************************************
	'�X�L�� start
	'------------------------------------------------------------------------------
	sSkillOS = GetSkill(rDB, rRS, "OS")
	sSkillApp = GetSkill(rDB, rRS, "Application")
	sSkillDL = GetSkill(rDB, rRS, "DevelopmentLanguage")
	sSkillDB = GetSkill(rDB, rRS, "Database")
	sSkillOther = GetOrderNote(rDB, rRS, "OtherSkill")
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
		sOtherNote = GetOrderNote(rDB, rRS, "OtherNote")
		DspNeedCondition = True
	End If
	'------------------------------------------------------------------------------
	'���̑����L���� end
	'******************************************************************************

	flgLine = False

	Response.Write "<h3>�K�v����</h3>" & vbCrLf

	'<�x�X�g�E�x�^�[�p�^�[���o��>
	If dbBestMatchStr & dbBetterMatchStr <> "" Then
		If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		Response.Write "<div class=""category1""><h4>ϯ�ݸ��߲��</h4>[<span style=""color:#0045F9;cursor:pointer;"" onclick=""window.open('" & HTTP_CURRENTURL & "/infomation/matchingpoint.asp','matchingpoint','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=400,height=300');"">�H</span>]</div>" & vbCrLf
		Response.Write "<div class=""value1"">" & vbCrLf

		If dbBestMatchStr <> "" Then
			Response.Write "<h5>�x�X�g</h5>" & vbCrLf
			Response.Write "<div class=""value2"">" & Replace(dbBestMatchStr, vbCrLf, "<br>") & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If dbBetterMatchStr <> "" Then
			If dbBestMatchStr <> "" Then Response.Write "<div class=""line1""></div>"
			Response.Write "<h5>�x�^�[</h5>" & vbCrLf
			Response.Write "<div class=""value2"">" & Replace(dbBetterMatchStr, vbCrLf, "<br>") & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		Response.Write "</div>" & vbCrLf
		Response.Write "<div style=""clear:both;""></div>" & vbCrLf
	End If
	'</�x�X�g�E�x�^�[�p�^�[���o��>

	'<�N��o��>
	If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
	flgLine = True
	Response.Write "<div class=""category1""><h4>�N��</h4></div>" & vbCrLf
	Response.Write "<div class=""value1""><p class=""m0"">" & sAge & "</p></div>" & vbCrLf
	Response.Write "<div style=""clear:both;""></div>" & vbCrLf
	'</�N��o��>

	'<��]�w���o��>
	If sFEHistory <> "" Then
		If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		Response.Write "<div class=""category1""><h4>��]�w��</h4></div>" & vbCrLf
		Response.Write "<div class=""value1""><p class=""m0"">" & sFEHistory & "</p></div>" & vbCrLf
		Response.Write "<div style=""clear:both;""></div>" & vbCrLf
	End If
	'</��]�w���o��>

	'******************************************************************************
	'���i�o�� start
	'------------------------------------------------------------------------------
	sClearSolid = " style=""border-top-width:0px;"""
	If flgLicense = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		Response.Write "<div class=""category1""><h4>���i</h4></div>" & vbCrLf
		Response.Write "<div class=""value1"">" & vbCrLf

		If sLicense <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
			Response.Write "<h5>���i</h5>" & vbCrLf
			Response.Write "<div class=""value2"">" & sLicense & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sLicenseOther <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True

			Response.Write "<h5>���̑����i</h5>" & vbCrLf
			Response.Write "<div class=""value2"">" & sLicenseOther & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		Response.Write "</div>" & vbCrLf
		Response.Write "<div style=""clear:both;""></div>" & vbCrLf
	End If
	'------------------------------------------------------------------------------
	'���i�o�� end
	'******************************************************************************

	'******************************************************************************
	'�X�L���o�� start
	'------------------------------------------------------------------------------
	sClearSolid = " style=""border-top-width:0px;"""
	If flgSkill = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		Response.Write "<div class=""category1""><h4>�X�L��</h4></div>" & vbCrLf
		Response.Write "<div class=""value1"">" & vbCrLf

		If sSkillOS <> "" Then
			Response.Write "<h5>�n�r</h5>" & vbCrLf
			Response.Write "<div class=""value2"">" & sSkillOS & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sSkillApp <> "" Then
			Response.Write "<h5>���ع����</h5>" & vbCrLf
			Response.Write "<div class=""value2"">" & sSkillApp & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sSkillDL <> "" Then
			Response.Write "<h5>�J������</h5>" & vbCrLf
			Response.Write "<div class=""value2"">" & sSkillDL & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sSkillDB <> "" Then
			Response.Write "<h5>�f�[�^�x�[�X</h5>" & vbCrLf
			Response.Write "<div class=""value2"">" & sSkillDB & "</div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sSkillOther <> "" Then
			Response.Write "<h5>���̑��X�L��</h5>" & vbCrLf
			Response.Write "<div class=""value2""><p class=""m0"">" & sSkillOther & "</p></div>" & vbCrLf
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		Response.Write "</div>" & vbCrLf
		Response.Write "<div style=""clear:both;""></div>" & vbCrLf
	End If
	'------------------------------------------------------------------------------
	'�X�L���o�� end
	'******************************************************************************

	'******************************************************************************
	'���̑����L���� start
	'------------------------------------------------------------------------------
	If sOtherNote <> "" Then
		If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		Response.Write "<div class=""category1""><h4>���L����</h4></div>" & vbCrLf
		Response.Write "<div class=""value1""><p class=""m0"">" & sOtherNote & "</p></div>" & vbCrLf
		Response.Write "<div style=""clear:both;""></div>" & vbCrLf

		sClearSolid = ""
	End If
	'------------------------------------------------------------------------------
	'���̑����L���� end
	'******************************************************************************

	If DspNeedCondition = True Then Response.Write "<br>"
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

	Response.Write "<h3>������</h3>" & vbCrLf

	If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
	flgLine = True

	Response.Write "<div class=""category1""><h4>���R�[�h</h4></div>"
	Response.Write "<div class=""value1""><p class=""m0"">" & dbOrderCode & "</p></div>"
	Response.Write "<div style=""clear:both;""></div>" & vbCrLf

	If flgEntryInfo = True Then
		If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		Response.Write "<div class=""category1""><h4>������@</h4></div>"
		Response.Write "<div class=""value1""><p class=""m0"">" & sEntryInfo & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>" & vbCrLf
	End If

	If flgProcess = True Then
		If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		Response.Write "<div class=""category1""><h4>�I�l�菇</h4></div>" & vbCrLf
		Response.Write "<div class=""value1"">" & vbCrLf

		If sProcess1 <> "" Then
			Response.Write "<p class=""m0"" style=""float:left; width:60px; color:#666666; font-weight:bold;"">�X�e�b�v�P</p>"
			Response.Write "<p class=""m0"" style=""float:left; width:400px;"">" & sProcess1 & "</p>"
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sProcess2 <> "" Then
			Response.Write "<p style=""width:60px; color:#666666; text-align:center;"">��</p>"
			Response.Write "<p class=""m0"" style=""float:left; width:60px; color:#666666; font-weight:bold;"">�X�e�b�v�Q</p>"
			Response.Write "<p class=""m0"" style=""float:left; width:400px;"">" & sProcess2 & "</p>"
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sProcess3 <> "" Then
			Response.Write "<p style=""width:60px; color:#666666; text-align:center;"">��</p>"
			Response.Write "<p class=""m0"" style=""float:left; width:60px; color:#666666; font-weight:bold;"">�X�e�b�v�R</p>"
			Response.Write "<p class=""m0"" style=""float:left; width:400px;"">" & sProcess3 & "</p>"
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		If sProcess4 <> "" Then
			Response.Write "<p style=""width:60px; color:#666666; text-align:center;"">��</p>"
			Response.Write "<p class=""m0"" style=""float:left; width:60px; color:#666666; font-weight:bold;"">�X�e�b�v�S</p>"
			Response.Write "<p class=""m0"" style=""float:left; width:400px;"">" & sProcess4 & "</p>"
			Response.Write "<div style=""clear:both;""></div>" & vbCrLf
		End If

		Response.Write "</div>" & vbCrLf
		Response.Write "<div style=""clear:both;""></div>" & vbCrLf
	End If

	If dbWValueURL <> "" Then
		If flgLine = True Then Response.Write "<div class=""line1"" style=""margin-left:15px;""></div>"
		flgLine = True

		Response.Write "<div class=""category1""><h4>���Ѝ̗p<br>�y�[�W</h4></div>"
		Response.Write "<div class=""value1""><a href=""" & dbWValueURL & """ target=""_blank""><img src=""/img/order/btn_wvalue.gif"" border=""0"" alt=""���Ѝ̗p�y�[�W""></a></div>"
		Response.Write "<div style=""clear:both;""></div>" & vbCrLf
	End If

	If DspHowToEntry = True Then Response.Write "<br>" & vbCrLf
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
	Response.Write "<h3 class=""sp"">�S���ҏ��</h3>"
	If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
	Response.Write "<div class=""category1""><h4>�S����</h4></div>"
	Response.Write "<div class=""value1""><p class=""m0"">" & sPerson & "</p></div>"
	Response.Write "<div style=""clear:both;""></div>"
	If sCSectionName <> "" Then
		If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
		Response.Write "<div class=""category1""><h4>�S������</h4></div>"
		Response.Write "<div class=""value1""><p class=""m0"">" & sCSectionName & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>"
	End If

	If dbPlanTypeName <> "mail" Then
		'���[���ۋ��v�����̏ꍇ�͘A������f�ڂ�
		If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
		flgLine = True

		Response.Write "<div class=""category1""><h4>�A����</h4></div>"

		Response.Write "<div class=""value1""><p class=""m0"">" & sContact & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>"
	End If

	Response.Write "<br>"
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
<h3>��y�C���^�r���[</h3>
<div class="freeprblock">
<%
		Do While GetRSState(oRS) = True
			dbSeq = oRS.Collect("Seq")
			dbProfile = oRS.Collect("Profile")
			dbQuestion = oRS.Collect("Question")
			dbAnswer = oRS.Collect("Answer")
			dbPublicFlag = oRS.Collect("PublicFlag")
			dbPictureFlag = oRS.Collect("PictureFlag")
%>
		<h4><%= dbProfile %></h4>
		<div style="clear:both;"></div>
<%
			If dbPictureFlag = "1" Then
				'��y�ʐ^�L��
%>
		<div style="width:580px; margin-left:20px;">
			<div style="float:left; width:182px; padding-top:5px;">
				<img src="/company/elderinterview/picture.asp?ordercode=<%= dbOrderCode %>&amp;seq=<%= dbSeq %>" alt="" border="1" width="180" height="135" style="border:1px solid:#999999;">
			</div>
			<div style="float:left; width:398px;">
				<p style="margin:0px; padding-left:5px;">��<%= dbQuestion %></p>
				<p style="margin:0px; padding-left:5px;"><%= dbAnswer %></p>
			</div>
			<div style="clear:both;"></div>
		</div>
<%
			Else
				'��y�ʐ^����
%>
		<p class="m0">��<%= dbQuestion %></p>
		<p class="m0"><%= dbAnswer %></p>
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
		sBranchName = ChkStr(rRS.Collect("LisDepartment"))
		sTel = ChkStr(rRS.Collect("LisTelephoneNumber"))

		sImg = "<img src=""/consultant/consultantimage.asp?ec=" & sEmployeeCode & """ alt=""���̋��l����S�����Ă���R���T���^���g"" border=""1"" width=""180"" height=""180"" style=""border-color:#666666;"">"
		sComment = Replace(ChkStr(rRS.Collect("ConsultantComment")), vbCrLf, "<br>")
		sComment = Replace(sComment, vbCr, "<br>")
		sComment = Replace(sComment, vbLf, "<br>")
		sConsultantPublicFlag = ChkStr(rRS.Collect("ConsultantPublicFlag"))
		sPictureFlag = ChkStr(rRS.Collect("ConsultantPictureFlag"))

		sConsultantLink = sEmployeeName
		If sConsultantPublicFlag = "1" Then
			sConsultantLink = "<a href=""" & HTTP_NAVI_CURRENTURL & "consultant/consultantdetail.asp?ec=" & sEmployeeCode & """>" & sEmployeeName & "</a>"
		End If
		sConsultantLink = sConsultantLink & "(�l�މ�ЁF���X�������)"
		'------------------------------------------------------------------------------
		'�R���T���^���g end
		'******************************************************************************

		sTitle = "�S���ҘA����"
		If sComment <> "" Then sTitle = "���̋��l����S�����Ă���R���T���^���g�̏���"
		Response.Write "<h3 class=""sp"">" & sTitle & "</h3>"
		Response.Write "<div class=""category1""><h4>�R���T���^���g</h4></div>"
		Response.Write "<div class=""value1""><p class=""m0"">" & sConsultantLink & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>"
		Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
		Response.Write "<div class=""category1""><h4>�S������</h4></div>"
		Response.Write "<div class=""value1""><p class=""m0"">" & sBranchName & "</p></div>"
		Response.Write "<div style=""clear:both;""></div>"
		Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
		Response.Write "<div class=""category1""><h4>�A����</h4></div>"
		Response.Write "<div class=""value1""><p class=""m0"">" & sTel & "<span style=""font-size:10px;"">�@�����₢���킹�̍ہA��L�u���R�[�h�v�Ɓu�����ƃi�r�������v�ƌ����ƃX���[�Y�ł��B</span></p></div>"
		Response.Write "<div style=""clear:both;""></div>"
		If sComment <> "" Then
			Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			Response.Write "<div class=""category1""><h4>����</h4></div>"
			Response.Write "<div class=""value1""><p class=""m0"">" & sComment & "</p></div>"
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "<br>"
		End If
	End If
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
				Case "001": sWorkingType = sWorkingType & "�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/koyoukeitai_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>�h���Ƃ�</a>�z" 
				Case "002","003": sWorkingType = sWorkingType & "�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>�l�ޏЉ�Ƃ�</a>�z" 
				Case "004": sWorkingType = sWorkingType & "�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/syoukaiyotei_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>�Љ�\��h���Ƃ�</a>�z" 
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

	Do While GetRSState(oRS) = True
		sJobType = sJobType & "(" & oRS.Collect("JobTypeName") & ")"
		oRS.MoveNext
		If GetRSState(oRS) = True Then sJobType = sJobType & "<br>"
	Loop
	Call RSClose(oRS)

	GetJobType = sJobType
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
	sSQL = "sp_GetDataWorkingTime '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sWST = ChkStr(oRS.Collect("DspWorkStartTime"))
		sWET = ChkStr(oRS.Collect("DspWorkEndTime"))
		If sWST & sWET <> "" Then
			sWorkingTime = sWorkingTime & sWST & "�`" & sWET
		End If
		oRS.MoveNext
		If GetRSState(oRS) = True And sWST & sWET <> "" Then sWorkingTime = sWorkingTime & "<br>"
	Loop
	Call RSClose(oRS)

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

		sSkill = sSkill & "<p style=""width:50%; float:left;"">" & oRS.Collect("SkillName")
		If ChkStr(oRS.Collect("Period")) <> "" Then
			sSkill = sSkill & "<br>�@<span style=""color:#339933;"">��</span>" & oRS.Collect("Period") & "�N�ȏ�͏���"
		End If
		sSkill = sSkill & "</p>"
		If iSkill Mod SKILLCOL = 0 Then sSkill = sSkill & "<br clear=""all"">"

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

		sLicense = sLicense & "<p style=""width:50%; float:left;"">" & oRS.Collect("LicenseName") & "</p>"
		If iLicense Mod LICENSECOL = 0 Then sLicense = sLicense & "<br clear=""all"">"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'���r���[�ŏI������ꍇ�̒���
	If sLicense <> "" And iLicense Mod LICENSECOL <> 0 Then sLicense = sLicense & "<br clear=""all"">"

	GetLicense = sLicense
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

	sSQL = "EXEC up_DtlOrderTitle '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		rTitle = ChkStr(oRS.Collect("JobTypeDetail")) & "&nbsp;" & ChkStr(oRS.Collect("PrefectureName"))
		rKeywords = "���l���,�]�E," & ChkStr(oRS.Collect("PrefectureName"))
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
			sSQL = "up_SearchRelationAccessOrder '" & CONF_OrderCode & "'"
			sTitle = "���̋��l���������l�͂���ȋ��l�������Ă��܂�"
		Case "2"
			sSQL = "up_SearchHighRelationOrder '" & CONF_OrderCode & "'"
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

	sURL = HTTP_CURRENTURL & "order/order_detail.asp"

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
	sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If sImg = "" And ChkStr(oRS.Collect("OptionNo1")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo1")
		If sImg = "" And ChkStr(oRS.Collect("OptionNo2")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo2")
		If sImg = "" And ChkStr(oRS.Collect("OptionNo3")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo3")
		If sImg = "" And ChkStr(oRS.Collect("OptionNo4")) <> "" Then sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo4")
	End If

	If sImg = "" And dbOrderType = "0" Then
		sSQL = "sp_GetDataPicture '" & dbCompanyCode & "', '1'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			sImg = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
		End If
	End If

	If sImg = "" Then sImg = "/img/nopicture180.gif"
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
				" '" & CONF_CompanyCode & "'" & _
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
	If rRS.Collect("InexperiencedPersonFlag") = "1" Then sHTML = sHTML & "<img src=""/img/no_experience.gif"" alt=""���o���ҁ^���V�����}"" width=""50"" height=""15"">&nbsp;"
	'�t�^�[���E�h�^�[��
	If rRS.Collect("UITurnFlag") = "1" Then sHTML = sHTML & "<img src=""/img/ui_turn.gif"" alt=""�t�^�[���E�h�^�[��"" width=""50"" height=""15"">&nbsp;"
	'��w���������d��
	If rRS.Collect("UtilizeLanguageFlag") = "1" Then sHTML = sHTML & "<img src=""/img/linguistic_job.gif"" alt=""��w���������d��"" width=""50"" height=""15"">&nbsp;"
	'�N�ԋx��120���ȏ�
	If rRS.Collect("ManyHolidayFlag") = "1" Then sHTML = sHTML & "<img src=""/img/year_holidaycnt.gif"" alt=""�N�ԋx��120���ȏ�"" width=""50"" height=""15"">&nbsp;"
	'2006/01/10 M.Hayashi ADD �t���b�N�X�^�C�����x����
	If rRS.Collect("FlexTimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_flextime.gif"" alt=""�t���b�N�X�^�C�����x����"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NearStationFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_nearstation.gif"" alt=""�w��(�k��5���ȓ�)"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NoSmokingFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_nosmoking.gif"" alt=""�։��E����"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NewlyBuiltFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_newlybuilt.gif"" alt=""�V�z�r���E�I�t�B�X(5�N�ȓ�)"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("LandmarkFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_landmark.gif"" alt=""���w(15�K�ȏ�)�r��"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("RenovationFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_renovation.gif"" alt=""���m�x�[�V�����r���E�I�t�B�X(5�N�ȓ�)"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("DesignersFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_designers.gif"" alt=""�f�U�C�i�[�Y�r���E�I�t�B�X"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("CompanyCafeteriaFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_companycafeteria.gif"" alt=""�Ј��H��"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("ShortOvertimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_shortovertime.gif"" alt=""�c��10h/���ȓ�"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("MaternityFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_maternity.gif"" alt=""�Y�x�E��x���т���"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("DressFreeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_dressfree.gif"" alt=""�������R"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("MammyFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_mammy.gif"" alt=""�q��ă}�}���}"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("FixedTimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_fixedtime.gif"" alt=""18���܂łɑގ�"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("ShortTimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_shorttime.gif"" alt=""1��6���Ԉȓ��J��"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("HandicappedFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_handicapped.gif"" alt=""��Q�Ҋ��}"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("RentAllFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_rentallflag.gif"" alt=""�Z���p�S�z�⏕����"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("RentPartFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_rentpartflag.gif"" alt=""�Z���p�ꕔ�⏕����"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MealsFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_mealsflag.gif"" alt=""�H���E�d���t���Č�"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MealsAssistanceFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_mealsassistanceflag.gif"" alt=""�H���⏕���x����"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("TrainingCostFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_trainingcostflag.gif"" alt=""���C������x����"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("EntrepreneurCostFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_entrepreneurcostflag.gif"" alt=""�N�Ƌ@�ޕ⏕���x����"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MoneyFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_moneyflag.gif"" alt=""�����q�E�ᗘ�q�⏕���x����"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("LandShopFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_landshopflag.gif"" alt=""�y�n�E�X�ܓ��񋟐��x����"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("FindJobFestiveFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_findjobfestiveflag.gif"" alt=""�A�E���j�������x����"" width=""50"" height=""15"">&nbsp;"
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
<div align="right" style="width:600px; margin-bottom:5px;">
	<div style="float:right; width:150px;"><a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button3.gif" alt="���O�C�����ĉ���" border="0"></a></div>
	<div style="float:right; width:150px; margin-right:3px;"><a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button1.gif" alt="�������o�^���ĉ���" border="0"></a></div>
	<div style="clear:both;"></div>
<!--
	<div align="center">
	<form action="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>" method="post">
	<input type="submit" value="����̕��͂����炩�牞��ł��܂�">
	</form>
	</div>
-->
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
	<div style="float:right; width:150px;"><a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button3.gif" alt="���O�C�����ĉ���" border="0"></a></div>
	<div style="float:right; width:150px;"><a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button1.gif" alt="�������o�^���ĉ���" border="0"></a></div>
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
<div align="center">
	<hr size="1">
	<p style="color:#ff0000;">
��������o�^����Ή���⎿�₪�\�ɂȂ�܂��I����<BR>
����̂��߂̗������������쐬����܂��B</p>
	<hr size="1">
	<div align="center" style="float:left; width:300px;color:#C51035;">���܂�ID���������łȂ�����<br><a href="<%= HTTPS_NAVI_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button1.gif" alt="�������o�^���ĉ���" border="0"></a></div>
	<div align="center" style="float:right; width:300px;color:#C51035;">�����ł�ID���������̕���<br><a href="<%= HTTPS_NAVI_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button3.gif" alt="���O�C�����ĉ���" border="0"></a></div>
	<div style="clear:both;"></div>
	<br>
</div>
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
	<div align="center" style="float:left; width:300px;color:#C51035;">���܂�ID���������łȂ�����<br><a href="<%= HTTPS_NAVI_CURRENTURL %>resume/staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button1.gif" alt="�������o�^���ĉ���" border="0"></a></div>
	<div align="center" style="float:right; width:300px;color:#C51035;">�����ł�ID���������̕���<br><a href="<%= HTTPS_NAVI_CURRENTURL %>resume/login/login.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/resume/order/order_detail.asp&amp;ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button3.gif" alt="���O�C�����ĉ���" border="0"></a></div>
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
