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
'�@�@�@�FDspOrderShowTypeSwitch		�F���l�[�ڍ׃y�[�W�̉�Џ��E�E����؂�ւ��{�^���ƎQ�Ɖ񐔂��o��
'�@�@�@�FDspOrderCatchCopy			�F���l�[�ڍ׃y�[�W�̃L���b�`�R�s�[�����i�傫���摜�Ȃǁj���o��
'�@�@�@�FDspOrderFreePR				�F���l�[�ڍ׃y�[�W�̃t���[�o�q���o��
'�@�@�@�FDspOrderPictureNow			�F���l�[�ڍ׃y�[�W�̏������摜���o��
'�@�@�@�FDspBusiness				�F���l�[�ڍ׃y�[�W�̋Ɩ����e���o��
'�@�@�@�FDspCondition				�F���l�[�ڍ׃y�[�W�̋Ζ��������o��
'�@�@�@�FDspNeedCondition			�F���l�[�ڍ׃y�[�W�̕K�v�������o��
'�@�@�@�FDspHowToEntry				�F���l�[�ڍ׃y�[�W�̉�������o��
'�@�@�@�FDspContact					�F���l�[�ڍ׃y�[�W�̒S���ҘA������o��
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
'�X�@�V�F2006/05/13 LIS K.Kokubo �쐬
'�@�@�@�F2007/11/22 LIS K.Kokubo up_SearchOrder��K�v�ŏ����̂��̂���������Ă���悤�ɂ������Ƃɂ��ύX�Bsp_GetDetailOrder����f�[�^���擾�B
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

	Dim sOrderCode			'���R�[�h
	Dim sOrderType			'�󒍎��
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
	Dim sYearlyIncome		'�N���\���p
	Dim sDailyIncome		'�����\���p
	Dim sMonthlyIncome		'�����\���p
	Dim sHourlyIncome		'�����\���p
	'��]�Ζ��`�ԁE��]�Ζ��n�A�C�R���@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��
	Dim sWorkingCode
	Dim sWorkingName
	Dim sWorkingPlacePrefectureName
	Dim sBiz
	Dim sBizName1
	Dim sBizName2
	Dim sBizName3
	Dim sBizName4
	Dim sBizPercentage1
	Dim sBizPercentage2
	Dim sBizPercentage3
	Dim sBizPercentage4
	Dim flgAddWatchList
	Dim flgBusiness

	If GetRSState(rRS) = False Then Exit Function

	sOrderCode = rRS.Collect("OrderCode")

	DspOrderListDetail = False

	sSQL = "sp_GetDetailOrder '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	sOrderType = ChkStr(oRS.Collect("OrderType"))

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
	'------------------------------------------------------------------------------
	sStationName = ""
	sSQL = "sp_GetDataNearbyStation '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	If GetRSState(oRS2) = True Then
		sStationName ="�y" & sStationName & GetStrNearbyStation(oRS2.Collect("StationName"), "", "") & "�z"
	End If
	'------------------------------------------------------------------------------
	'�Ŋ�w end
	'******************************************************************************

	'**************************************************************************
	'���[�����M�ς݊m�F start
	'--------------------------------------------------------------------------
	If vUserType = "staff" Then
		sSQL = "sp_GetDataMailHistory '" & vUserID & "', '', '" & sOrderCode & "'"
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
	sImgOrderState = "&nbsp;"
	'�A�N�Z�X����100�𒴂��Ă���΁uHOT�v�\���i���X�����j
	If oRS.Collect("AccessCount") > 100 Then
		sImgOrderState = sImgOrderState & "<img src=""/img/c_HOT_green.gif"" alt=""�l�C"">&nbsp;"
	End If

	'UPDATE�ƍ�������10�����������Łu�V���v�\��(���X����)
	If oRS.Collect("UpdateDay") > NOW()-10 Then
		sImgOrderState = sImgOrderState & "<img src=""/img/c_NEW_green.gif"" alt=""�V��"">&nbsp;"
	End If

	'���o���҂n�j�̏ꍇ�A�킩�΃}�[�N�\��(���X����)
	If oRS.Collect("InexperiencedPersonFlag") = "1" Then
		sImgOrderState = sImgOrderState & "<img src=""/img/no_experience.gif"" alt=""���o���ҁ^���V�����}"">&nbsp;"
	End If

	'�t�^�[���E�h�^�[��
	If oRS.Collect("UITurnFlag") = "1" Then
		sImgOrderState = sImgOrderState & "<img src=""/img/ui_turn.gif"" alt=""�t�^�[���E�h�^�[��"">&nbsp;"
	End If

	'��w���������d��
	If oRS.Collect("UtilizeLanguageFlag") = "1" Then
		sImgOrderState = sImgOrderState & "<img src=""/img/linguistic_job.gif"" alt=""��w���������d��"">&nbsp;"
	End If

	'�N�ԋx��120���ȏ�
	If oRS.Collect("ManyHolidayFlag") = "1" Then
		sImgOrderState = sImgOrderState & "<img src=""/img/year_holidaycnt.gif"" alt=""�N�ԋx��120���ȏ�"">&nbsp;"
	End If

	'�t���b�N�X�^�C�����x���� ------2006/01/10 Hayashi ADD
	If oRS.Collect("FlexTimeFlag") = "1" And oRS.Collect("OrderType") = "0" And oRS.Collect("CompanyKbn") = "1" Then
		sImgOrderState = sImgOrderState & "<img src=""/img/flextime.gif"" alt=""�t���b�N�X�^�C�����x����"">&nbsp;"
	End If

	sSQL = "sp_GetDataWorkingType '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True
		sWorkingCode = oRS2.Collect("WorkingTypeCode")
		sWorkingName = oRS2.Collect("WorkingTypeName")

		sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_w" & sWorkingCode & ".gif"" alt=""" & sWorkingName & """ width=""50"" height=""15"">&nbsp;"

		oRS2.MoveNext
	Loop
	sWorkingPlacePrefectureName = oRS.Collect("WorkingPlacePrefectureName")
	If oRS.Collect("Workingplaceprefecturecode") >= "048" Then
		sWorkingPlacePrefectureName = "�C�O"
	End If

	sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_p" & oRS.Collect("WorkingPlacePrefectureCode") & ".gif"" alt=""" & sWorkingplaceprefecturename & """ width=""50"" height=""15"">&nbsp;"

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
		If sImgSub <> "" Then sImgSub = sImgSub & "<div style=""clear:both;""></div>"
	Else
		If sCompanyPictureFlag = "1" And sOrderType = "0" Then
			sImgMain = "<img src=""/company/imgdsp.asp?companycode=" & oRS2.Collect("CompanyCode") & "&amp;optionno=1"" alt="""" border=""0"" width=""" & PICSIZEW & """ height=""" & PICSIZEH & """>"
			flgImg = True
		End If
	End If

	Call RSClose(oRS2)
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
			sWorkingType = sWorkingType & "�y<a href=""javascript:void(0)"" onclick=""window.open('/staff/s_shokai.htm','count','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=300,height=200')"">�l�ޏЉ�</a>�z"
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
	sWorkingPlace = oRS.Collect("WorkingPlacePrefectureName") & oRS.Collect("WorkingPlaceCity") & "&nbsp;"
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
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	flgAddWatchList = False
	sSQL = "up_GetDataWatchList '" & vUserID & "', '', '', '" & sOrderCode & "', ''"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	If GetRSState(oRS2) = False Then
		flgAddWatchList = True
	End If
	Call RSClose(oRS2)
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'******************************************************************************
	'���l�[�f�ڊ��� start
	'------------------------------------------------------------------------------
	'��ƃ��O�C�����ȊO�̂Ƃ��Ɍf�ڊ�����\��
	sPublishLimitStr = GetDateStr(oRS.Collect("riyotodate"), "/")

	If sPublishLimitStr = "" Then
		sPublishLimitStr = "�펞��W��" 
	End If

	sPublishLimitStr = sPublishLimitStr & "&nbsp;"
	'------------------------------------------------------------------------------
	'���l�[�f�ڊ��� end
	'******************************************************************************

	'******************************************************************************
	'�d���̊��� start�@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��
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
	'------------------------------------------------------------------------------
	'�d���̊��� end
	'******************************************************************************


%>
<input type="hidden" name="CONF_OrderCodes" value="<%= oRS.Collect("OrderCode") %>">
<table border="0" class="old">
	<tbody>
	<tr>
		<td class="old11" style="padding-left:0px; width:600px;" valign="middle">
<%
	If vUserType = "" Or vUserType = "staff" Then
		'�񃍃O�C�����A�X�^�b�t���O�C����

		If G_USERID <> "" And G_FLGRESUME = False or G_FLGRESUME = False Then
			'�����ƃi�r�̋��l�[�ꗗ�̏ꍇ�͈ȉ���\��
			'�E���l�[�t�q�k�����[�����M
			'�E�E�H�b�`���X�g�֕ۑ�
%>
			<div style="float:left;width:420px;">
			<img src="/img/list_companyicon.gif" alt="" align="left"><%= sTitleCompanyName %>
			<h3 style="margin-left:5px;">��<a href="<%= HTTP_NAVI_CURRENTURL %>order/order_detail.asp?OrderCode=<% = oRS.Collect("OrderCode") %>"><%= sTitleJobName %></a><%= sImgMail %></h3>
			</div>
			<div align="right" style="float:right;font-size:11px;width:113px;">
			<a href="<%= HTTPS_NAVI_CURRENTURL %>order/sendmail_jobofferaddress.asp?OrderCode=<% = oRS.Collect("OrderCode") %>" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;"><img src="/img/order/ordermail.gif" style="margin-bottom:6px;" border="0" alt="���l�������[�����M" align="top"></a>
			<a href="<%= HTTPS_NAVI_CURRENTURL %>order/sendmail_jobofferaddress.asp?OrderCode=<% = oRS.Collect("OrderCode") %>" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;"><img src="/img/order/orderwachlist.gif" border="0" alt="�E�H�b�`���X�g�ɒǉ�" align="top"></a>
			</div>
			<div style="clear:both;"></div>
<%
		Else
			'���������̋��l�[�ꗗ�̏ꍇ�͈ȉ���\�����Ȃ��I
			'�E�t�q�k�����[�����M
			'�E�E�H�b�`���X�g�֕ۑ�
%>
			<p class="m0"><img src="/img/list_companyicon.gif" alt="" align="left"><%= sTitleCompanyName %></p>
			<h3 style="margin-left:5px;">��<a href="../order/order_detail.asp?OrderCode=<% = oRS.Collect("OrderCode") %>"><%= sTitleJobName %></a><%= sImgMail %></h3>
<%
		End If

	ElseIf vUserType = "company" Then
		'��ƃ��O�C����
%>
			<p class="m0"><img src="/img/list_companyicon.gif" alt="" align="left"><%= sTitleCompanyName %></p>
			<h3 style="margin-left:5px;">��<a href="../order/order_detail.asp?OrderCode=<% = oRS.Collect("OrderCode") %>"><%= sTitleJobName %></a><%= sImgMail %></h3>
<%
	End If
%>
		</td>
	</tr>
	<tr>
		<td class="old12">
			<div style="float:left;"><%= sImgOrderState %></div>
			<div align="right" style="font-size:10px;line-height:14px;">�f�ڊ����F<%= sPublishLimitStr %></div>
			<div style="clear:both;"></div>
			<table border="0" class="old2">
<%
	If sCatchCopy <> "" Then
%>
				<caption><%= sCatchCopy %></caption>
<%
	End If
%>
				<tbody>
				<tr>
					<td rowspan="2" valign="top">
<%
	If flgImg = True Then
		'�摜���L��ꍇ�̃��C�A�E�g
%>
					<div class="old21" valign="top" style="margin:0px 12px;">
					<b>�y�S���Ɩ��̐����z</b><br><%= sBusinessDetail %>
					</div>
					<div class="old21" valign="top" style="width:240px; float:left; margin:0px 5px;">
						<a href="../order/order_detail.asp?OrderCode=<% = oRS.Collect("OrderCode") %>" title="<%= sTitleCompanyName %>"><%= sImgMain %></a>
						<%= sImgSub %>
					</div>
<%
	Else
		'�摜�������ꍇ�̃��C�A�E�g
%>
					<div class="old21" valign="top" style="width:239px; float:left; margin:0px 5px;">
					<b>�y�S���Ɩ��̐����z</b><br><%= sBusinessDetail %>
					</div><br>
<%
	End If
%>
					<table style="width:330px; margin-left:3px;">
						<tr>
							<td style="font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;">
							�Ζ��`��
							</td>
							<td style="background-color:#eeeeee; padding:5px 0px 5px 10px; border-bottom:solid 3px #ffffff;">
							<%= sWorkingType %>
							</td>
						</tr>
						<tr>
							<td style="font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;">
							�Ζ��n
							</td>
							<td style="background-color:#eeeeee; padding-left:10px; border-bottom:solid 3px #ffffff;">
							<%= sWorkingPlace %><%= sStationName %>
							</td>
						</tr>
						<tr>
							<td style="font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;">
							���^
							</td>
							<td style="background-color:#eeeeee; padding:5px 0px 5px 10px; border-bottom:solid 3px #ffffff;">
<%
			If sYearlyIncome <> "" Then
%>
							<p>�N�� <%= sYearlyIncome %></p>
<%
			End If

			If sMonthlyIncome <> "" Then
%>
							<p>���� <%= sMonthlyIncome %></p>
<%
			End If

			If sDailyIncome <> "" Then
%>
							<p>���� <%= sDailyIncome %></p>
<%
			End If

			If sHourlyIncome <> "" Then
%>
							<p>���� <%= sHourlyIncome %></p>
<%
			End If
%>
							</td>
						</tr>
<%
	If sBizName1 <> "" Then
%>

						<tr>
							<td style="font-weight:bold; background-color:#E1FBCD; width:70px; border-bottom:solid 3px #ffffff; text-align:center;">
							�d���̊���
							</td>
							<td style="background-color:#eeeeee; border-bottom:solid 3px #ffffff; padding-left:0px; line-height:14px;">
								<table>
									<tr>
										<td style="padding:5px 0px 5px 7px;">
										<script type="text/javascript" language="javascript">
											viewWorkAvg(<%= sBizPercentage1 %>, <%= sBizPercentage2 %>, <%= sBizPercentage3 %>, <%= sBizPercentage4 %>)
										</script>
										</td>
										<td>
<%
		If sBizName1 <> "" Then Response.Write "<p style=""font-size:10px; line-height:12px;""><span style=""color:#ff9999;"">��</span>" & sBizPercentage1 & "%�@" & sBizName1 & "</p>"
		If sBizName2 <> "" Then Response.Write "<p style=""font-size:10px; line-height:12px;""><span style=""color:#9999ff;"">��</span>" & sBizPercentage2 & "%�@" & sBizName2 & "</p>"
		If sBizName3 <> "" Then Response.Write "<p style=""font-size:10px; line-height:12px;""><span style=""color:#99ff99;"">��</span>" & sBizPercentage3 & "%�@" & sBizName3 & "</p>"
		If sBizName4 <> "" Then Response.Write "<p style=""font-size:10px; line-height:12px;""><span style=""color:#ffff99;"">��</span>" & sBizPercentage4 & "%�@" & sBizName4 & "</p>"
%>
										</td>
									</tr>
								</table>
							</td>
						</tr>
<%
	End If
%>
					</table>
						<div align="right" style="margin:3px 5px;">
							<a href="../order/order_detail.asp?OrderCode=<% = oRS.Collect("OrderCode") %>">
								<img src="/img/detail_button2.gif" border="0" alt="">
							</a>
						</div>
					</td>
				</tr>
			</table>
			<div style="clear:both;"></div>
		</td>
	</tr>
<%
	If oRS.Collect("CompanyCode") = vUserID And vMyOrder = "1" Then
%>
	<tr>
		<td class="old13">
			<table class="old3">
				<tbody>
				<tr>
					<td class="old31">���R�[�h(<% = oRS.Collect("OrderCode") %>)</td>
					<td class="old32">���</td>
					<td class="old33">
						<%= sProgress %>
						<select name="CONF_PublicFlags" <%= sPublicListDsp %>>
						<option value="1"<% If oRS.Collect("PublicFlag") = "1" Then Response.Write(" selected") %>>�f��</option>
						<option value="0"<% If oRS.Collect("PublicFlag") = "0" Then Response.Write(" selected") %>>��f��</option>
						</select>
					</td>
					<td class="old34">�f�ړ�<br>�o�^��</td>
					<td class="old35"><%= sPublicDay %><br><%= sRegistDay %></td>
					<td class="old36"><input type="checkbox" name="CONF_DeleteFlags" value="<%= oRS.Collect("OrderCode") %>">�폜</td>
				</tr>
				</tbody>
			</table>
		</td>
	</tr>
<%
	End If
%>
	<tr>
		<td class="old14"></td>
	</tr>
</table>
<%
	DspOrderListDetail = True
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
'�X�@�V�F
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
	Dim sStation			'�Ŋ�w
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

	sSQL = "sp_GetDetailOrder '" & rRS.Collect("OrderCode") & "'"
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

	'******************************************************************************
	'�Ŋ�w start
	'------------------------------------------------------------------------------
	sStation = ""
	sSQL = "sp_GetDataNearbyStation '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True
		sStation = sStation & GetStrNearbyStation(oRS2.Collect("StationName"), oRS2.Collect("ToStationTime"), oRS2.Collect("ToStationRemark"))
		oRS2.MoveNext
		If GetRSState(oRS2) = True Then sStation = sStation & "<br>"
	Loop
	'------------------------------------------------------------------------------
	'�Ŋ�w end
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
		<tr>
			<th>�Ŋ�w</th>
			<td><%= sStation %></td>
		</tr>
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
'�X�@�V�F
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
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvOrderCode		�F
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/company_order.asp
'******************************************************************************
Function DspCompanyInfo(ByRef rDB, ByRef rRS, ByVal vOrderCode, ByVal vUserType, ByVal vUserID)
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
	Dim sClass				'�g�p����X�^�C���V�[�g�̃N���X�@�摜�̗L���ŕω�
	Dim sLineClass			'
	Dim flgLine				'�������t���O
	Dim sAddTitle			'�h����Ƃ̏��̏ꍇ�́u�h���v�����ږ��ɕt����

	If GetRSState(rRS) = False Then Exit Function

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

	If sCompanyPictureFlag = "1" Then
		sClass = "value1"
		sLineClass = "odline2"
	Else
		sClass = "value2"
		sLineClass = "odline1"
	End If

	flgLine = False
%>
<div class="companyblock">
	<h3><%= sAddTitle %>��Ə��</h3>
<%
	If sCompanyPictureFlag = "1" Then
%>
	<div style="width:302px; float:right;"><img id="imgcompany" src="<%= HTTPS_NAVI_CURRENTURL %>company/imgdsp.asp?companycode=<%= sCompanyCode %>&amp;optionno=1" alt="�C���[�W�ʐ^" width="300" height="225" style="border:1px solid #999999;"></div>
	<div style="float:left; width:295px;">
<%
	End If

	If sCompanyCode <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>��ƃR�[�h</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sCompanyCode %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sEstablishYear <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>�ݗ��N�x</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sEstablishYear %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sCapitalAmount <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>���{�z</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sCapitalAmount %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sListClass <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>�������J</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sListClass %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sEmployeeNum <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>�Ј���</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sEmployeeNum %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sIndustryType <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>�Ǝ�</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sIndustryType %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sAddress <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>�{�ЏZ��</h4></div>
		<div class="<%= sClass %>"><p class="m0"><%= sAddress %></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sHomePage <> "" Then
		If flgLine = True Then Response.Write "<table class=""" & sLineClass & """ border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
		<div class="category"><h4>�z�[���y�[�W</h4></div>
		<div class="<%= sClass %>"><p class="m0"><a href="<%= sHomePage %>" target="_blank">���̊�Ƃ̃z�[���y�[�W</a></p></div>
		<div style="clear:both;"></div>
<%
	End If

	If sCompanyPictureFlag = "1" Then
%>
	</div>
	<div style="clear:both;"></div>
<%
	End If
%>
</div>
<%
End Function

'******************************************************************************
'�T�@�v�F��Ə��̂o�q�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvOrderCode		�F
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/company_order.asp
'******************************************************************************
Function DspCompanyPR(ByRef rDB, ByRef rRS, ByVal vOrderCode, ByVal vUserType, ByVal vUserID)
	Const WELFARECOL = "3"	'���������̂P�s������̗�

	Dim sOrderType			'�󒍎��
	Dim sCompanyKbn			'��Ƌ敪
	Dim sBusiness			'���Ɠ��e
	Dim sPR					'��ƏЉ�
	Dim sWelfare			'��������
	Dim iWelfare			'���������J�E���g
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

	If ChkStr(rRS.Collect("FlexTimeFlag")) = "1" Then
		iWelfare = iWelfare + 1
		If iWelfare Mod WELFARECOL = 1 Then sWelfare = sWelfare & "<tr>"
		sWelfare = sWelfare & "<td class=""welfare""><p class=""m0"">�t���b�N�X�^�C��</p></td>"
		If iWelfare Mod WELFARECOL = 0 Then sWelfare = sWelfare & "</tr>"
	End If

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
	'------------------------------------------------------------------------------
	'�������� end
	'******************************************************************************

	flgPR = False
	If sBusiness & sPR & sWelfare <> "" Then flgPR = True

	flgLine = False
	sClass = "value2"

	If flgPR = True Then
%>
<div class="companyblock">
	<h3><%= sAddTitle %>�o�q���</h3>
<%
		If sBusiness <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
	<div class="category"><h4>���Ɠ��e</h4></div>
	<div class="<%= sClass %>"><p class="m0"><%= sBusiness %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sPR <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
	<div class="category"><h4>��Ђo�q</h4></div>
	<div class="<%= sClass %>"><p class="m0"><%= sPR %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sWelfare <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
	<div class="category"><h4>��������</h4></div>
	<div class="<%= sClass %>"><p class="m0"><%= sWelfare %></p></div>
	<div style="clear:both;"></div>
<%
		End If
%>
</div>
<br>
<%
	End If
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̃��X�̏Љ��E�h�����Ə����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
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
		'�������J end
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
	End If

	flgLine = False

	If sListClass & sIndustryType & sPR <> "" Then
		DspLisOrderCompanyInfo = True
%>
<h3><%= sIntrDisp %>��Ə��</h3>
<%
		If sListClass <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>�������J</h4></div>
<div class="value1"><p class="m0"><%= sListClass %></p></div>
<div style="clear:both;"></div>
<%
		End If

		If sIndustryType <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>�Ǝ�</h4></div>
<div class="value1"><p class="m0"><%= sIndustryType %></p></div>
<div style="clear:both;"></div>
<%
		End If


		If sPR <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
			

%>
<div class="category1"><h4>���Ɠ��e</h4></div>
<div class="value1"><p class="m0"><%= sPR %></p></div>
<div style="clear:both;"></div>
<% End If %>
				<p class="m0" style="font-size:10px;margin:0 0 20px 20px;">
				���l��<%= left(sIntrDisp,2) %>�ł��ē����邨�d���̂��߁A�ڂ�����Џ��͉��̃{�^���₨�d�b�ȂǂŒ��ڂ��⍇�����������B
		</p>
<%
	End If
End Function

'******************************************************************************
'�T�@�v�F�h����Ƃ̔h�����Ə����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvMyOrder		�F���Ћ��l�[�t���O
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/company_order.asp
'******************************************************************************
Function DspTempOrderCompanyInfo(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder)
	Dim sCompanyCode		'��ƃR�[�h
	Dim sCompanyName		'��Ж�
	Dim sCompanyName_F		'��Ж��J�i
	Dim sAddress			'�Z��
	Dim sTel				'�d�b�ԍ�
	Dim sIndustryType		'�Ǝ�
	Dim sCapitalAmount		'���{�z
	Dim sListClass			'�������J
	Dim sEmployeeNum		'�Ј���
	Dim flgLine				'�������t���O
	Dim flgData				'�o�̓f�[�^�̗L���t���O

	flgData = False
	If GetRSState(rRS) = False Then Exit Function

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
	sCompanyName = ChkStr(rRS.Collect("TempCompanyName"))
	sCompanyName_F = ChkStr(rRS.Collect("TempCompanyName_F"))
	
	If sCompanyName_F <> "" Then sCompanyName = sCompanyName & "(" & sCompanyName_F & ")"
	'------------------------------------------------------------------------------
	'��Ж� end
	'******************************************************************************

	'******************************************************************************
	'�Z�� start
	'------------------------------------------------------------------------------
	sAddress = ""
	If rRS.Collect("TempPost_U") & rRS.Collect("TempPost_L") <> "" Then
		sAddress = "��" & rRS.Collect("TempPost_U") & "-" & rRS.Collect("TempPost_L") & "<br>"
	End If
	sAddress = sAddress & rRS.Collect("TempPrefectureName") & rRS.Collect("TempCity") & rRS.Collect("TempTown") & rRS.Collect("TempAddress")
	'------------------------------------------------------------------------------
	'�Z�� end
	'******************************************************************************

	'******************************************************************************
	'�d�b�ԍ� start
	'------------------------------------------------------------------------------
	sTel = ChkStr(rRS.Collect("TempTelephoneNumber"))
	'------------------------------------------------------------------------------
	'�d�b�ԍ� end
	'******************************************************************************

	'******************************************************************************
	'�Ǝ� start
	'------------------------------------------------------------------------------
	sIndustryType = ChkStr(rRS.Collect("TempIndustryTypeName"))
	If sIndustryType <> "" Then flgData = True
	'------------------------------------------------------------------------------
	'�Ǝ� end
	'******************************************************************************

	'******************************************************************************
	'���{�z start
	'------------------------------------------------------------------------------
	sCapitalAmount = ChkStr(rRS.Collect("TempCapitalAmount"))
	sCapitalAmount = GetJapaneseYen(sCapitalAmount)
	If sCapitalAmount <> "" Then flgData = True
	'------------------------------------------------------------------------------
	'���{�z end
	'******************************************************************************

	'******************************************************************************
	'�������J start
	'------------------------------------------------------------------------------
	sListClass = ChkStr(rRS.Collect("TempListClass"))
	If sListClass <> "" Then flgData = True
	'------------------------------------------------------------------------------
	'�������J end
	'******************************************************************************

	'******************************************************************************
	'�Ј��� start
	'------------------------------------------------------------------------------
	sEmployeeNum = ChkStr(rRS.Collect("TempAllEmployeeNumber"))
	If sEmployeeNum <> "" Then sEmployeeNum = sEmployeeNum & "�l"
	If sEmployeeNum <> "" Then flgData = True
	'------------------------------------------------------------------------------
	'�Ј��� end
	'******************************************************************************

	flgLine = False

	If flgData = True Then
%>
<h3>�h�����Ə��</h3>
<%
		If vMyOrder = "1" Then
%>
<p class="m0" style="margin:0px 0px 10px 20px;">����Ɩ��A�Z���A�d�b�ԍ��͔���J���ł��B</p>
<%
			If sCompanyName <> "" Then
				If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
				flgLine = True
%>
<div class="category1"><h4>��Ɩ�</h4></div>
<div class="value1"><p class="m0"><%= sCompanyName %></p></div>
<div style="clear:both;"></div>
<%
			End If

			If sAddress <> "" Then
				If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
				flgLine = True
%>
<div class="category1"><h4>�Z��</h4></div>
<div class="value1"><p class="m0"><%= sAddress %></p></div>
<div style="clear:both;"></div>
<%
			End If

			If sTel <> "" Then
				If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
				flgLine = True
%>
<div class="category1"><h4>�d�b�ԍ�</h4></div>
<div class="value1"><p class="m0"><%= sTel %></p></div>
<div style="clear:both;"></div>
<%
			End If
		End If

		If sIndustryType <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>�Ǝ�</h4></div>
<div class="value1"><p class="m0"><%= sIndustryType %></p></div>
<div style="clear:both;"></div>
<%
		End If

		If sIndustryType <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>���{�z</h4></div>
<div class="value1"><p class="m0"><%= sCapitalAmount %></p></div>
<div style="clear:both;"></div>
<%
		End If

		If sIndustryType <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>�������J</h4></div>
<div class="value1"><p class="m0"><%= sListClass %></p></div>
<div style="clear:both;"></div>
<%
		End If

		If sIndustryType <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>�Ј���</h4></div>
<div class="value1"><p class="m0"><%= sEmployeeNum %></p></div>
<div style="clear:both;"></div>
<%
		End If

		Response.Write "<br>"
	End If
End Function

'******************************************************************************
'�T�@�v�F���l�[�R���g���[���{�^��
'���@���FrDB				�F�ڑ�����DBConnection
'�@�@�@�FrRS				�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType			�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID			�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvMyOrder			�F���Ћ��l�[���ۂ� ["1"]���Ћ��l�[ ["0"]���Ћ��l�[�łȂ�
'�@�@�@�FvJobTypeLimitFlag	�F�E�퐔���������z���Ă��Ȃ��� ["1"]OK ["0"]NO
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'******************************************************************************
Function DspOrderControlButton(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder, ByVal vJobTypeLimitFlag)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError
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
	sSQL = "up_GetDataWatchList '" & vUserID & "', '', '', '" & sOrderCode & "', ''"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = False Then
		flgAddWatchList = True
	End If
	Call RSClose(oRS2)
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	If vMyOrder = "1" Then
		'******************************************************************************
		'���Ћ��l�[�̏ꍇ start
		'------------------------------------------------------------------------------
		If sHakouFlag = "1" Then
%>
<h2 class="csubtitle">���Ћ��l�[�̑���</h2>
<div class="subcontent">
<%
			If sPermitFlag = "1" And sRiyoFlag = "0" Then
%>
	<p class="cctrltitle">���E�Ҍ����E�X�J�E�g���[��</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
			<p style="color:#ff0000;">���̋��l�[�͂܂��f�ڂ���Ă���܂���i�f�ڊJ�n���O�ł��j�B���̂��߁A���E�҂̌����͗��p�ł��܂���B</p>
		</div>
	</div>
<%
			ElseIf sPermitFlag = "0" Then
%>
	<p class="cctrltitle">���E�Ҍ����E�X�J�E�g���[��</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
			<p style="color:#ff0000;">���̋��l�[�͂܂��f�ڂ���Ă���܂���i�R�����ł��j�B���̂��߁A���E�҂̌����͗��p�ł��܂���B</p>
		</div>
	</div>
<%
			ElseIf sPermitFlag = "1" And sPublicFlag = "1" And sRiyoFlag = "1" Then
%>
	<p class="cctrltitle">���E�Ҍ����E�X�J�E�g���[��</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
			<input type="button" value="���E�҂���������" style="width:150px; color:#aa3300;" onclick="Go_Edit('10');">
			<span style="font-size:10px; color:#666666;">�E�E�E���̋��l�[�́A�E��E�Ζ��n�E�ٗp�`�Ԃ𖞂������E�҂��������܂��B</span>
		</div>
		<div style="padding:0px 0px 5px 15px;">
			<input type="button" value="���E�҂��ڍ׌���" style="width:150px; color:#aa3300;" onclick="Go_Edit('11');">
			<span style="font-size:10px; color:#666666;">�E�E�E���̋��l�[����A�ڍׂȌ����������w�肵�ċ��E�҂��������܂��B</span><br>
		</div>
	</div>
<%
			Else
%>
	<p class="cctrltitle">���E�Ҍ����E�X�J�E�g���[��</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
			<p style="color:#ff0000;">�f�ڂ���Ă��Ȃ����l�[����̃X�J�E�g�͂ł��܂���B</p>
		</div>
	</div>
<%
			End If

			If vJobTypeLimitFlag = True Then
				'�E�퐔���������z���Ă��Ȃ���΁u���l�[�R�s�[�쐬�v�{�^���̕\��
%>
	<p class="cctrltitle">���l�[�R�s�[�쐬</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
			<input type="button" value="���l�[���R�s�[" style="width:100px; color:#3333ff;" onclick="Go_Edit('4');">
			<span style="font-size:10px; color:#666666;">�E�E�E���̋��l�[�����ƂɁA�V�������l�[���쐬���܂��B</span><br>
		</div>
	</div>
<%
			End If
%>
	<p class="cctrltitle">���l����ҏW����</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
			<div style="float:left; width:290px;">
				<input type="button" value="���Џ��X�V" style="width:100px;" onclick="Go_Edit('1');">
				<span style="font-size:10px; color:#666666;">�E�E�E���Џ����X�V���܂��B</span>
			</div>
			<div style="float:right; width:290px;">
				<input type="button" value="�摜�o�^" style="width:100px;" onclick="Go_Edit('5');">
				<span style="font-size:10px; color:#666666;">�E�E�E�摜���f�ڂ��܂��B</span><br>
			</div>
			<div style="clear:both;"></div>
		</div>
		<div style="padding:0px 0px 5px 15px;">
			<div style="float:left; width:290px; margin:0px;">
				<input type="button" value="��W���ҏW" style="width:100px;" onclick="Go_Edit('2');">
				<span style="font-size:10px; color:#666666;">�E�E�E�o�q�E��W�v����ҏW���܂��B</span>
			</div>
			<div style="float:right; width:290px;">
				<input type="button" value="�X�L�������ҏW" style="width:100px;" onclick="Go_Edit('3');">
				<span style="font-size:10px; color:#666666;">�E�E�E�K�v�X�L���E���i��ҏW���܂��B</span><br>
			</div>
			<div style="clear:both;"></div>
		</div>
	</div>

	<p class="cctrltitle">���[���e���v���[�g</p>
	<div style="padding:5px 0px;">
		<div style="padding:0px 0px 5px 15px;">
<%
			If iMailTemplateCnt >= 5 Then
				'���[���e���v���[�g��������ɒB���Ă���ꍇ�͐V�K�쐬�ł��Ȃ�
%>
			<p style="color:#ff0000; font-size:10px;">���[���e���v���[�g��������ɒB���Ă���̂ŁA����ȏ�쐬�ł��܂���B</p>
<%
			Else
				'���[���e���v���[�g��������ɒB���Ă��Ȃ��ꍇ�͐V�K�쐬�ł���
%>
			<input type="button" value="�V�K�쐬" style="width:100px;" onclick="location.href = '<%= HTTPS_NAVI_CURRENTURL %>mailtemplate/regist.asp?ordercode=<%= sOrderCode %>';">
			<span style="font-size:10px; color:#666666;">�E�E�E���̋��l�̃��[���e���v���[�g��V�K�ɍ쐬���܂��B</span><br>
<%
			End If
%>
			<p style="color:#ff0000; font-size:10px;">�����[���e���v���[�g�͋��l�[���ɍ쐬���܂��B</p>
<%
			sSQL = "up_GetListMailTemplate '" & G_USERID & "', '" & sOrderCode & "'"
			flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
			If GetRSState(oRS2) = True Then Response.Write "<hr size=""1"">"
			Do While GetRSState(oRS2) = True
				sAncMT = "?ordercode=" & oRS2.Collect("OrderCode") & "&amp;seq=" & oRS2.Collect("SEQ")
				sAncMT = "<a href=""" & HTTPS_NAVI_CURRENTURL & "mailtemplate/regist.asp" & sAncMT & """>" & oRS2.Collect("Subject") & "</a>"
%>
			<div style="width:585px;">
				<div style="float:left; width:85px;"><%= GetDetail("MailTemplateType", oRS2.Collect("MailTemplateTypeCode")) %></div>
				<div style="float:left; width:500px;"><%= sAncMT %></div>
				<div style="clear:both;"></div>
			</div>
<%
				oRS2.MoveNext
			Loop
%>
		</div>
	</div>
</div>
<%
		End If
		'------------------------------------------------------------------------------
		'���Ћ��l�[�̏ꍇ end
		'******************************************************************************
	ElseIf vUserType = "staff" Then
		'******************************************************************************
		'���O�C�����E�҂̏ꍇ start
		'------------------------------------------------------------------------------
		If rRS.Collect("PublicFlag") = "1" Then
%>
<div class="subcontent" style="margin-bottom:15px;">
	<div style="padding:5px 0px;">
		<p class="sctrltitle">����E����E�E�H�b�`���X�g</p>
		<div style="padding:0px 0px 5px 15px;">
			<div style="float:left; width:195px;">
				<p class="m0" style="margin-right:20px; font-size:10px; color:#666666; text-align:center;">�����̕�W�։��僁�[���̍쐬</p>
				<input type="button" value="���僁�[���𑗐M����" style="width:180px;" onclick="contactCompany('');">
			</div>
			<div align="center" style="float:left; width:195px;">
				<p class="m0" style="font-size:10px; color:#666666; text-align:center;">�����̕�W�֎��⃁�[���̍쐬</p>
				<input type="button" value="���⃁�[���𑗐M����" onclick="contactCompany('1');">
			</div>
			<div style="float:left; width:195px;">
				<p class="m0" style="margin-left:20px; font-size:10px; color:#666666; text-align:center;">��<a href="watchlist_info.htm" onclick="window.open(this.href, 'mywindow6', 'width=300, height=150, menubar=no, toolbar=no, scrollbars=yes'); return false;" style="color:#0045F9;">�E�H�b�`���X�g</A>�֒ǉ�</p>
<%
			If flgAddWatchList = True Then
%>
				<div align="right"><input type="button" value="���̋��l�[��ǉ�����" style="width:180px;" onclick="document.forms.frmMain.action='../staff/watchlist_register.asp';document.forms.frmMain.submit();"></div>
<%
			Else
%>
				<p class="m0" style="margin-left:20px; text-align:center; font-weight:bold;">���ɓo�^�ς݂ł�</p>
<%
			End If
%>
			</div>
			<div style="clear:both;"></div>
		</div>
	</div>
</div>
<%
		Else
%>
	<div align="center"><b>���̋��l�[�͌f�ڂ��I�����Ă��܂��B���[�����M�͂ł��܂���B</b></div>
<%
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
'�@�@�@�FrRS				�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
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

	If vMyOrder = "1" Then
		'******************************************************************************
		'���Ћ��l�[�̏ꍇ start
		'------------------------------------------------------------------------------
%>
<script type="text/javascript">
<!--
function Go_Edit(pNo){
	switch (pNo){
		case '1':
			//�u��Џ��X�V�v��
			location.href = '<%= HTTPS_NAVI_CURRENTURL & vUserType %>/company_reg1.asp';
			return true;
		case '2':
			//�u��W���ҏW�v��
			document.forms.frmMain.mode.value="edit"
			document.forms.frmMain.action="<%= HTTPS_NAVI_CURRENTURL & vUserType %>/company_reg2.asp";
			break;
		case '3':
			//�u�X�L���v��
			document.forms.frmMain.mode.value="edit"
			document.forms.frmMain.action="<%= HTTPS_NAVI_CURRENTURL & vUserType %>/company_reg3.asp";
			break;
		case '4':
			//�u�R�s�[���ċ��l�[�̍쐬�v��
			document.forms.frmMain.mode.value="copy"
			document.forms.frmMain.action="<%= HTTPS_NAVI_CURRENTURL & vUserType %>/company_reg2.asp";
			break;
		case '5':
			//�u���l�[�ʐ^�o�^�v��
			location.href = '<%= HTTP_NAVI_CURRENTURL %>company/order_img_listnow.asp?ordercode=<%= sOrderCode %>';
			return true;
		case '10':
			//��������
			document.forms.frmMain.action="<%= HTTP_NAVI_CURRENTURL %>staff/person_list.asp";
			break;
		case '11':
			//�ڍ׌���
			document.forms.frmMain.action="<%= HTTP_NAVI_CURRENTURL %>staff/person_search_detail.asp";
			break;
		default:
			return false;
	}
	document.forms.frmMain.submit();
}
//-->
</script>
<%
		'------------------------------------------------------------------------------
		'���Ћ��l�[�̏ꍇ end
		'******************************************************************************
	ElseIf vUserType = "staff" Then
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
'�@�@�@�FrRS				�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
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
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'******************************************************************************
Function DspOrderCompanyName(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType
	Dim sCompanyCode		'��ƃR�[�h
	Dim sCompanyName		'��Ɩ���
	Dim sCompanyNameF		'��Ɩ��̃J�i
	Dim sCompanyKbn			'��Ƌ敪
	Dim sCompanySpeciality	'��Ɠ���
	Dim sPublishLimitStr	'�f�ڊ����\���p������
	Dim sCautionStr			'�f�ڊ����\�����ӕ���������
	Dim flgNowPublic		'���݌f�ڒ��̋��l�[���� '[True]�f�ڒ� [False]��f��

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'��Ж� start
	'------------------------------------------------------------------------------
	sCompanyName = rRS.Collect("CompanyName")
	sCompanyNameF = rRS.Collect("CompanyName_F")
	sCompanyKbn = rRS.Collect("CompanyKbn")
	sCompanySpeciality = rRS.Collect("CompanySpeciality")
	sOrderType = rRS.Collect("OrderType")

	Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
	'------------------------------------------------------------------------------
	'��Ж� end
	'******************************************************************************

	'******************************************************************************
	'���l�[�f�ڊ��� start
	'------------------------------------------------------------------------------
	sCautionStr = "<p style=""line-height:11px;text-align:right;font-size:11px;"">�������O�Ɍf�ڏI������ꍇ������܂��B</p>"

	'�f�ڒ� or ��f��
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'�ЊO�Č��Ȃ�riyotodate���A�Г��Č��Ȃ�PublicLimitDay��\��
	'�ЊO�Č� OrderType = 0
	'�Г��Č� OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(rRS.Collect("riyotodate"), "/")
	Else
		sPublishLimitStr = rRS.Collect("PublicLimitDay")
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = 0 Then
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
	If sOrderType = "2" Then
		'���X�Љ�Č��̏ꍇ�́u�]�E���k�Č��v�C���[�W��\��
%>
	<img src="/img/order/counselable_order.gif" width="150" height="25" alt="�]�E�x�����󂯂ĉ��傷�鋁�l�ł�">
<%
	End If

	If vUserType = "" Or vUserType = "staff" Then
		'�񃍃O�C�����A�X�^�b�t���O�C����

		If G_USERID <> "" And G_FLGRESUME = False And flgNowPublic = True Then
			'�����ƃi�r�Ƀ��O�C�����̏ꍇ�́A��Ɩ��{�f�ڊ����{���l�[�t�q�k���[�����M
%>
	<div class="m0" style="width:420px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="float:right; padding:0px;"><img src="../ImgQRCode.asp?Code=<%= rRS.Collect("OrderCode") %>" alt="QRCode"></div>
	<div style="text-align:right; font-size:11px; padding-top:6px;"><a href="../order/sendmail_jobofferaddress.asp?OrderCode=<% = rRS.Collect("OrderCode") %>&detail=1" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=470');return false;"><img src="/img/staff/mail/mailhei.gif" border="0" align="bottom" alt="���l�[�����[�����M"> ���l�[�����[�����M</a></div>
	<p style="text-align:right;padding:4px 0px 0px 0px;">�f�ڊ����F<%= sPublishLimitStr %></p>
	<div style="clear:both;"></div>
	<%= sCautionStr %>
	<div style="clear:both;"></div>
<%
		ElseIf G_FLGRESUME = False And flgNowPublic = True Then
			'�����ƃi�r�ɔ񃍃O�C���̏ꍇ�́A��Ɩ��{�f�ڊ����{���l�[�t�q�k���[�����M
%>
	<div class="m0" style="width:420px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="float:right; padding:0px;"><img src="../ImgQRCode.asp?Code=<%= rRS.Collect("OrderCode") %>" alt="QRCode"></div>
	<div style="text-align:right; font-size:11px; padding-top:6px;"><a href="../order/sendmail_jobofferaddress.asp?OrderCode=<% = rRS.Collect("OrderCode") %>&detail=1" onclick="window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=380');return false;"><img src="/img/staff/mail/mailhei.gif" border="0" align="bottom" alt="���l�[�����[�����M"> ���l�[�����[�����M</a></div>
	<p style="text-align:right;padding:4px 0px 0px 0px;">�f�ڊ����F<%= sPublishLimitStr %></p>
	<div style="clear:both;"></div>
	<%= sCautionStr %>
	<div style="clear:both;"></div>
<%
		Else
			'���������̋��l�[�̏ꍇ�́A��Ɩ��{�f�ڊ����̂�
%>
	<div class="m0" style="width:420px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="float:right; padding:0px;"><img src="../ImgQRCode.asp?Code=<%= rRS.Collect("OrderCode") %>" alt="QRCode"></div>
	<p style="text-align:right;padding-top:21px;">�f�ڊ����F<%= sPublishLimitStr %></p>
	<div style="clear:both;"></div>
	<%= sCautionStr %>
	<div style="clear:both;"></div>
<%
		End If
	Else
%>
	<div class="m0" style="width:420px; float:left;">
		<div style="font-size:14px; font-weight:bold;"><%= sCompanyName %></div>
		<div style="font-size:10px; color:#666666;"><%= sCompanyNameF %></div>
	</div>
	<div style="float:right; padding:0px;"><img src="../ImgQRCode.asp?Code=<%= rRS.Collect("OrderCode") %>" alt="QRCode"></div>
	<p style="text-align:right;padding-top:21px;">�f�ڊ����F<%= sPublishLimitStr %></p>
	<div style="clear:both;"></div>
	<%= sCautionStr %>
	<div style="clear:both;"></div>
<%
	End If
%>
</div>
<%
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̉�Џ��E�E����؂�ւ��{�^���ƎQ�Ɖ񐔂��o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvType			�F�\�������̎�� ["0"]�E���� ["1"]��Џ��
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

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	sOrderType = rRS.Collect("OrderType")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'��̓I�E�햼
	sJobTypeDetail = rRS.Collect("JobTypeDetail")
	'�X�V��
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	If sJobTypeDetail <> "" Then sJobTypeDetail = sJobTypeDetail & "�̂��d�����ڍ�"
%>
<div style="width:600px; margin-bottom:5px;">
	<div style="float:left; width:350px; margin:0px;">
<%
	If vType = "0" Then
		'�d������\�����̏ꍇ
%>
		<div style="float:left; width:93px; margin:0px;"><img src="/img/order/tab_orderdetail_on.gif" alt="<%= sJobTypeDetail %>" border="0" width="93" height="22"></div>
<%
		If sOrderType = "0" Then
			'��ʂ̋��l�L���̏ꍇ�͉�Џ��ւ̃����N��\��
%>
		<div style="float:left; width:93px; margin:0px;"><a href="./company_order.asp?poc=<%= sOrderCode %>" title="��Џ��"><img src="/img/order/tab_companyinfo_off.gif" alt="��Џ��" border="0" width="93" height="22"></a></div>
<%
		End If
	ElseIf vType = "1" Then
		'��Џ���\�����̏ꍇ
%>
		<div style="float:left; width:93px; margin:0px;"><a href="./order_detail.asp?ordercode=<%= sOrderCode %>"><img src="/img/order/tab_orderdetail_off.gif" alt="<%= sJobTypeDetail %>" border="0" width="93" height="22"></a></div>
<%
		If sOrderType = "0" Then
			'��ʂ̋��l�L���̏ꍇ�͉�Џ��ւ̃����N��\��
%>
		<div style="float:left; width:93px; margin:0px;"><img src="/img/order/tab_companyinfo_on.gif" alt="��Џ��" border="0" width="93" height="22"></div>
<%
		End If
	End If
%>
		<div class="clear:both; margin:0px;"></div>
	</div>
	<div align="right" style="float:right; width:250px;">
		<p class="m0">���ԎQ�Ɖ񐔁F<%= vAccessCount %>��@�X�V���F<%= sUpdateDay %></p>
	</div>
	<div style="clear:both;"><img src="/img/order/tab_border.gif" alt="" width="600" height="5"></div>
</div>
<%
End Function

'******************************************************************************
'�T�@�v�F���l�[�̃L���b�`�R�s�[�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/company_order.asp
'�@�@�@�F�����ƃi�r/order/order_detail_entity.asp
'******************************************************************************
Function DspOrderCatchCopy(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType
	Dim sCompanyCode
	Dim sOrderCode

	Dim sOptionNo			'�傫���ʐ^�̔ԍ�
	Dim sCompanyPictureFlag	'��Ǝʐ^�t���O ["1"]�L ["0"]��
	Dim sImg1
	Dim sClass

	If GetRSState(rRS) = False Then Exit Function

	sOrderType = rRS.Collect("OrderType")
	sOrderCode = rRS.Collect("OrderCode")
	sCompanyCode = rRS.Collect("CompanyCode")

	'******************************************************************************
	'�傫���摜 start
	'------------------------------------------------------------------------------
	sOptionNo = ""
	sImg1 = ""
	sSQL = "up_GetListOrderPictureNow '" & sCompanyCode & "', '" & sOrderCode & "', 'orderpicture'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
			sOptionNo = oRS.Collect("OptionNo1")
			sImg1 = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=" & sOptionNo
		End If
	End If

	If sImg1 = "" And sOrderType = "0" Then
		sSQL = "sp_GetDataPicture '" & sCompanyCode & "', '1'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			sImg1 = "/company/imgdsp.asp?companycode=" & sCompanyCode & "&amp;optionno=1"
		End If
	End If
	'------------------------------------------------------------------------------
	'�傫���摜 end
	'******************************************************************************

	If sImg1 <> "" Then
%>
<div id="catchcopy" style="width:600px;">
	<div style="float:right; width:300px;"><img class="big" src="<%= sImg1 %>" alt="" border="1" width="300" height="225" style="border:1px solid #999999;"></div>
	<h2><%= rRS.Collect("JobTypeDetail") %></h2>
	<div style="margin:10px 0px;"><%= GetImgOrderSpeciality(rDB, rRS) %></div>
	<p class="m0"><%= rRS.Collect("CatchCopy") %></p>
	<br clear="all">
</div>
<%
	Else
%>
<div id="catchcopy" style="width:600px;">
	<h2 style="width:600px;"><%= rRS.Collect("JobTypeDetail") %></h2>
	<div style="margin:10px 0px;"><%= GetImgOrderSpeciality(rDB, rRS) %></div>
	<p class="m0" style="width:600px;"><%= rRS.Collect("CatchCopy") %></p>
	<div style="clear:both;"></div>
</div>
<%
	End If
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̃t���[�o�q���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/company_order.asp
'�@�@�@�F�����ƃi�r/order/order_detail_entity.asp
'******************************************************************************
Function DspOrderFreePR(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sPRTitle1			'�o�q�^�C�g��1
	Dim sPRTitle2			'�o�q�^�C�g��2
	Dim sPRTitle3			'�o�q�^�C�g��3
	Dim sPRContents1		'�o�q��1
	Dim sPRContents2		'�o�q��2
	Dim sPRContents3		'�o�q��3
	Dim flgPR				'�o�q�L���t���O [True]�L [False]��

	If GetRSState(rRS) = False Then Exit Function

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
	<h3>�o�q</h3>
	<div class="freeprblock">
<%
		If sPRTitle1 <> "" Or sPRContents1 <> "" Then
%>
		<h4><%= sPRTitle1 %></h4>
		<div style="clear:both;"></div>
		<p class="m0"><%= sPRContents1 %></p>
<%
		End If

		If sPRTitle2 <> "" Or sPRContents2 <> "" Then
%>
		<h4><%= sPRTitle2 %></h4>
		<div style="clear:both;"></div>
		<p class="m0"><%= sPRContents2 %></p>
<%
		End If

		If sPRTitle3 <> "" Or sPRContents3 <> "" Then
%>
		<h4><%= sPRTitle3 %></h4>
		<div style="clear:both;"></div>
		<p class="m0"><%= sPRContents3 %></p>
<%
		End If
%>
	</div>
<%
	End If
End Function

'******************************************************************************
'�T�@�v�F���l��Ɖ摜�ꗗ�\���g�s�l�k�\��
'�쐬�ҁFLis Kokubo
'�쐬���F2006/12/27
'���@���FvCompanyCode	�F��ƃR�[�h
'�@�@�@�FvOrderCode		�F���R�[�h
'�@�@�@�FvCategoryCode	�F�J�e�S���R�[�h
'�g�p��F
'���@�l�F
'******************************************************************************
Function DspOrderPictureNow(ByVal vCompanyCode, ByVal vOrderCode, ByVal vCategoryCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sURL
	Dim flgPicture

	flgPicture = False
	sSQL = "up_GetListOrderPictureNow '" & vCompanyCode & "', '" & vOrderCode & "', '" & vCategoryCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = True Then
		If Len(oRS.Collect("OptionNo2")) > 0 Or Len(oRS.Collect("OptionNo3")) > 0 Or Len(oRS.Collect("OptionNo4")) > 0 Then
%>
<div align="center" style="padding:5px 15px; background-color:#e1fbcd; margin-bottom:40px;">
<div style="width:570px;">
<%
			sURL = ""
			If Len(oRS.Collect("OptionNo2")) > 0 Then
				sURL = "/company/imgdsp.asp?companycode=" & vCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo2")
%>
<div align="right" style="float:left; width:190px;">
	<div style="width:182px; background-color:#ffffff;"><img src="<%= sURL %>" alt="<%= oRS.Collect("Caption2") %>" width="180" height="135" border="1" style="border:1px solid #999999;"></div>
	<p class="m0" align="left" style="width:182px; font-size:10px;"><%= oRS.Collect("Caption2") %></p>
</div>
<%
			End If

			sURL = ""
			If Len(oRS.Collect("OptionNo3")) > 0 Then
				sURL = "/company/imgdsp.asp?companycode=" & vCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo3")
%>
<div align="right" style="float:left; width:190px;">
	<div style="width:182px; background-color:#ffffff;"><img src="<%= sURL %>" alt="<%= oRS.Collect("Caption3") %>" width="180" height="135" border="1" style="border:1px solid #999999;"></div>
	<p class="m0" align="left" style="width:182px; font-size:10px;"><%= oRS.Collect("Caption3") %></p>
</div>
<%
			End If

			sURL = ""
			If Len(oRS.Collect("OptionNo4")) > 0 Then
				sURL = "/company/imgdsp.asp?companycode=" & vCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo4")
%>
<div align="right" style="float:left; width:190px;">
	<div style="width:182px; background-color:#ffffff;"><img src="<%= sURL %>" alt="<%= oRS.Collect("Caption4") %>" width="180" height="135" border="1" style="border:1px solid #999999;"></div>
	<p class="m0" align="left" style="width:182px; font-size:10px;"><%= oRS.Collect("Caption4") %></p>
</div>
<%
			End If

			Response.Write "<br clear=""all"">"
%>
</div>
</div>
<%
		End If
	End If
End Function

'******************************************************************************
'�T�@�v�F���l�[�̋Ɩ����e���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'******************************************************************************
Function DspBusiness(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode			'���R�[�h
	Dim sCompanyCode		'��ƃR�[�h
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
	If GetRSState(rRS) = True Then

		'******************************************************************************
		'��ƃR�[�h start
		'------------------------------------------------------------------------------
		sOrderCode = rRS.Collect("OrderCode")
		sCompanyCode = rRS.Collect("CompanyCode")
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
	End If

	flgLine = False
	If flgBusiness = True Then
%>
<h3>�Ɩ����e</h3>
<%
		If sBusinessDetail <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>�S���Ɩ�</h4></div>
<div class="value1"><p class="m0"><%= sBusinessDetail %></p></div>
<div style="clear:both;"></div>
<%
		End If

		If sBiz <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>�d���̊���</h4></div>
<%'<div class="value1"><%= sBiz % ></div>%>
<div class="value1">
	<table border="0">
		<tbody>
		<tr>
			<td>
<script type="text/javascript" language="javascript">
	viewWorkAvg(<%= sBizPercentage1 %>, <%= sBizPercentage2 %>, <%= sBizPercentage3 %>, <%= sBizPercentage4 %>)
</script>
			</td>
			<td style="padding-left:5px; vertical-align:middle;">
				<table border="0">
					<tbody>
<%
			If sBizName1 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#ff9999; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage1 & "%</td><td>" & sBizName1 & "</td></tr>"
			If sBizName2 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#9999ff; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage2 & "%</td><td>" & sBizName2 & "</td></tr>"
			If sBizName3 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#99ff99; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage3 & "%</td><td>" & sBizName3 & "</td></tr>"
			If sBizName4 <> "" Then Response.Write "<tr><td style=""width:16px; background-color:#ffff99; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & sBizPercentage4 & "%</td><td>" & sBizName4 & "</td></tr>"
%>
					</tbody>
				</table>
			</td>
		</tr>
		</tbody>
	</table>
</div>
<div style="clear:both;"></div>
<%
		End If
%>
<br>
<%
	End If
End Function

'******************************************************************************
'�T�@�v�F���l�[�̋Ζ��������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'******************************************************************************
Function DspCondition(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode			'���R�[�h
	Dim sCompanyCode		'��ƃR�[�h
	Dim sOrderType			'���l�[���
	Dim sCompanyKbn			'��Ƌ敪
	Dim sJobTypeDetail		'�E��ڍ�
	Dim sSalary				'���^
	Dim sYearlyIncome		'�N��
	Dim sYearlyIncomeMin	'�N��
	Dim sYearlyIncomeMax	'�N��
	Dim sMonthlyIncome		'����
	Dim sMonthlyIncomeMin	'����
	Dim sMonthlyIncomeMax	'����
	Dim sDailyIncome		'����
	Dim sDailyIncomeMin		'����
	Dim sDailyIncomeMax		'����
	Dim sHourlyIncome		'����
	Dim sHourlyIncomeMin	'����
	Dim sHourlyIncomeMax	'����
	Dim sPercentagePay		'������
	Dim sSalaryRemark		'���^���l
	Dim sTrafficFee			'��ʔ�
	Dim sTrafficFeeType		'
	Dim sTrafficFeeMonth	'��ʔ�^�P����
	Dim sTime				'����
	Dim sWorkRange			'�A�Ɗ���
	Dim sWorkStartDay		'�A�ƊJ�n��
	Dim sWorkEndDay			'�A�ƏI����
	Dim sWorkUpdate			'�A�Ɗ��Ԃ̍X�V�L��
	Dim sWorkingTime		'�A�Ǝ���
	Dim sWorkTimeRemark		'�A�Ǝ��Ԕ��l
	Dim sHoliday			'�x��
	Dim sWeeklyHolidayType	'�T�x
	Dim sHolidayRemark		'�x�����l
	Dim sWorkingPlace		'�A�Əꏊ
	Dim sWPSection			'�Ζ��敔��
	Dim sWPTel				'�Ζ���d�b�ԍ�
	Dim sWPAddress			'�Ζ���Z��
	Dim sMAP				'�n�}���
	Dim sNearbyStation		'�Ŋ�w
	Dim sNearbyRailway		'�Ŋ񉈐�
	Dim sTransfer
	Dim sClearSolid
	Dim flgSalary
	Dim flgTime
	Dim flgHoliday
	Dim flgWorkingPlace
	Dim flgLine
	Dim flgLine2

	DspCondition = False

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'��ƃR�[�h start
		'------------------------------------------------------------------------------
		sOrderCode = rRS.Collect("OrderCode")
		sCompanyCode = rRS.Collect("CompanyCode")
		sOrderType = rRS.Collect("OrderType")
		sCompanyKbn = rRS.Collect("CompanyKbn")
		'------------------------------------------------------------------------------
		'��ƃR�[�h end
		'******************************************************************************

		'******************************************************************************
		'�E��ڍ� start
		'------------------------------------------------------------------------------
		sJobTypeDetail = rRS.Collect("JobTypeDetail")
		'------------------------------------------------------------------------------
		'�E��ڍ� end
		'******************************************************************************

		'******************************************************************************
		'���^ start
		'------------------------------------------------------------------------------
		sYearlyIncomeMin = ChkStr(rRS.Collect("YearlyIncomeMin"))
		sYearlyIncomeMax = ChkStr(rRS.Collect("YearlyIncomeMax"))
		If sYearlyIncomeMin = "0" Then sYearlyIncomeMin = ""
		If sYearlyIncomeMax = "0" Then sYearlyIncomeMax = ""
		If sYearlyIncomeMin <> "" Then sYearlyIncomeMin = GetJapaneseYen(sYearlyIncomeMin)
		If sYearlyIncomeMax <> "" Then sYearlyIncomeMax = GetJapaneseYen(sYearlyIncomeMax)
		If sYearlyIncomeMin & sYearlyIncomeMax <> "" Then
			If sYearlyIncomeMin <> "" Then sYearlyIncome = sYearlyIncome & sYearlyIncomeMin
			sYearlyIncome = sYearlyIncome & "&nbsp;�`&nbsp;"
			If sYearlyIncomeMax <> "" Then sYearlyIncome = sYearlyIncome & sYearlyIncomeMax
		End If

		sMonthlyIncomeMin = ChkStr(rRS.Collect("MonthlyIncomeMin"))
		sMonthlyIncomeMax = ChkStr(rRS.Collect("MonthlyIncomeMax"))
		If sMonthlyIncomeMin = "0" Then sMonthlyIncomeMin = ""
		If sMonthlyIncomeMax = "0" Then sMonthlyIncomeMax = ""
		If sMonthlyIncomeMin <> "" Then sMonthlyIncomeMin = GetJapaneseYen(sMonthlyIncomeMin)
		If sMonthlyIncomeMax <> "" Then sMonthlyIncomeMax = GetJapaneseYen(sMonthlyIncomeMax)
		If sMonthlyIncomeMin & sMonthlyIncomeMax <> "" Then
			If sMonthlyIncomeMin <> "" Then sMonthlyIncome = sMonthlyIncome & sMonthlyIncomeMin
			sMonthlyIncome = sMonthlyIncome & "&nbsp;�`&nbsp;"
			If sMonthlyIncomeMax <> "" Then sMonthlyIncome = sMonthlyIncome & sMonthlyIncomeMax
		End If

		sDailyIncomeMin = ChkStr(rRS.Collect("DailyIncomeMin"))
		sDailyIncomeMax = ChkStr(rRS.Collect("DailyIncomeMax"))
		If sDailyIncomeMin = "0" Then sDailyIncomeMin = ""
		If sDailyIncomeMax = "0" Then sDailyIncomeMax = ""
		If sDailyIncomeMin <> "" Then sDailyIncomeMin = GetJapaneseYen(sDailyIncomeMin)
		If sDailyIncomeMax <> "" Then sDailyIncomeMax = GetJapaneseYen(sDailyIncomeMax)
		If sDailyIncomeMin & sDailyIncomeMax <> "" Then
			If sDailyIncomeMin <> "" Then sDailyIncome = sDailyIncome & sDailyIncomeMin
			sDailyIncome = sDailyIncome & "&nbsp;�`&nbsp;"
			If sDailyIncomeMax <> "" Then sDailyIncome = sDailyIncome & sDailyIncomeMax
		End If

		sHourlyIncomeMin = ChkStr(rRS.Collect("HourlyIncomeMin"))
		sHourlyIncomeMax = ChkStr(rRS.Collect("HourlyIncomeMax"))
		If sHourlyIncomeMin = "0" Then sHourlyIncomeMin = ""
		If sHourlyIncomeMax = "0" Then sHourlyIncomeMax = ""
		If sHourlyIncomeMin <> "" Then sHourlyIncomeMin = GetJapaneseYen(sHourlyIncomeMin)
		If sHourlyIncomeMax <> "" Then sHourlyIncomeMax = GetJapaneseYen(sHourlyIncomeMax)
		If sHourlyIncomeMin & sHourlyIncomeMax <> "" Then
			If sHourlyIncomeMin <> "" Then sHourlyIncome = sHourlyIncome & sHourlyIncomeMin
			sHourlyIncome = sHourlyIncome & "&nbsp;�`&nbsp;"
			If sHourlyIncomeMax <> "" Then sHourlyIncome = sHourlyIncome & sHourlyIncomeMax
		End If

'		sYearlyIncome = GetMoneyRange(ChkStr(rRS.Collect("YearlyIncomeMin")), ChkStr(rRS.Collect("YearlyIncomeMax")), 1)
'		sMonthlyIncome = GetMoneyRange(ChkStr(rRS.Collect("MonthlyIncomeMin")), ChkStr(rRS.Collect("MonthlyIncomeMax")), 1)
'		sDailyIncome = GetMoneyRange(ChkStr(rRS.Collect("DailyIncomeMin")), ChkStr(rRS.Collect("DailyIncomeMax")), 1)
'		sHourlyIncome = GetMoneyRange(ChkStr(rRS.Collect("HourlyIncomeMin")), ChkStr(rRS.Collect("HourlyIncomeMax")), 1)
		sPercentagePay = ChkStr(rRS.Collect("PercentagePayFlag"))
		sSalaryRemark = Replace(ChkStr(rRS.Collect("IncomeRemark")), vbCrLf, "<br>")
		sSalaryRemark = Replace(sSalaryRemark, vbCr, "<br>")
		sSalaryRemark = Replace(sSalaryRemark, vbLf, "<br>")
		sTrafficFee = ""
		sTrafficFeeType = ChkStr(rRS.Collect("TrafficFeeType"))
		sTrafficFeeMonth = ChkStr(rRS.Collect("MonthTrafficFee"))
		flgSalary = False

		'���^
		sSalary = ""
		If sYearlyIncome <> "" Then
			sSalary = sSalary & sYearlyIncome
			flgSalary = True
		End If
		If sMonthlyIncome <> "" Then
			If sSalary <> "" Then sSalary = sSalary & "<br>"
			sSalary = sSalary & sMonthlyIncome
			flgSalary = True
		End If
		If sDailyIncome <> "" Then
			If sSalary <> "" Then sSalary = sSalary & "<br>"
			sSalary = sSalary & sDailyIncome
			flgSalary = True
		End If
		If sHourlyIncome <> "" Then
			If sSalary <> "" Then sSalary = sSalary & "<br>"
			sSalary = sSalary & sHourlyIncome
			flgSalary = True
		End If

		'������
		If sPercentagePay <> "" Then
			If sPercentagePay = "1" Then
				sPercentagePay = "����"
			ElseIf sPercentagePay = "0" Then
				sPercentagePay = "�Ȃ�"
			End If
			flgSalary = True
		End If

		'��ʔ�
		If ChkStr(rRS.Collect("NaviTrafficPayFlag")) = "1" Then 
			sTrafficFee = "��ʔ�x������" & sTrafficFeeType
			If IsNumber(sTrafficFeeMonth, 0, False) = True Then
				sTrafficFee = sTrafficFee & "(" & FormatCanma(sTrafficFeeMonth) & "�~�^��)"
			End If
			flgSalary = True
		End If

		If flgSalary = True Then DspCondition = True
		'------------------------------------------------------------------------------
		'���^ end
		'******************************************************************************

		'******************************************************************************
		'���� start
		'------------------------------------------------------------------------------
		sWorkRange = ""
		sWorkStartDay = ChkStr(rRS.Collect("WorkStartDay"))
		sWorkEndDay = ChkStr(rRS.Collect("WorkEndDay"))
		sWorkingTime = GetWorkingTime(rDB, rRS)
		sWorkTimeRemark = ChkStr(rRS.Collect("WorkTimeRemark"))
		flgTime = False

		'�A�Ɗ���
		If sWorkStartDay & sWorkEndDay <> "" Then
			If sWorkStartDay <> "" Then sWorkRange = sWorkRange & GetDateStr(sWorkStartDay, "/")
			If sWorkRange <> "" Then sWorkRange = sWorkRange & "�`"
			If sWorkEndDay <> "" Then sWorkRange = sWorkRange & GetDateStr(sWorkEndDay, "/")
		End If
		If sOrderType = "1" Then
			If rRS.Collect("WorkUpdateFlag") = "1" Then
				sWorkUpdate = "�L"
			Else
				sWorkUpdate = "��"
			End If
			sWorkRange = sWorkRange & "(�X�V" & sWorkUpdate & ")"
		End If

		If sWorkRange & sWorkingTime & sWorkTimeRemark <> "" Then
			flgTime = True
			DspCondition = True
		End If
		'------------------------------------------------------------------------------
		'���� end
		'******************************************************************************

		'******************************************************************************
		'�x�� start
		'------------------------------------------------------------------------------
		sWeeklyHolidayType = ChkStr(rRS.Collect("WeeklyHolidayTypeName"))
		sHolidayRemark = ChkStr(rRS.Collect("HolidayRemark"))
		flgHoliday = False

		If sWeeklyHolidayType & sHolidayRemark <> "" Then
			flgHoliday = True
			DspCondition = True
		End If
		'------------------------------------------------------------------------------
		'�x�� end
		'******************************************************************************

		'******************************************************************************
		'�Ζ��� start
		'------------------------------------------------------------------------------
		sWorkingPlace = ""
		sWPSection = ""
		sWPTel = ""
		sWPAddress = ""
		sMAP = ""
		sNearbyStation = GetNearbyStation(rDB, rRS)
		sNearbyRailway = GetNearbyRailway(rDB, rRS)
		flgWorkingPlace = False

		If sOrderType = "0" Then
			sWPSection = ChkStr(rRS.Collect("WorkingPlaceSection"))
			sWPTel = ChkStr(rRS.Collect("WorkingPlaceTelephoneNumber"))
			sWPAddress = ChkStr(rRS.Collect("WorkingPlaceAddressAll"))
		Else
			sWPAddress = ChkStr(rRS.Collect("WorkingPlacePrefectureName")) & ChkStr(rRS.Collect("WorkingPlaceCity"))
		End If
		If ChkStr(rRS.Collect("ExistsMap")) = "1" Then sMAP = "<div style=""margin:5px 0px;""><input type=""button"" value=""�n�}�m�F"" onclick=""open('/map/showmap.asp?mapOrderCode=" & sOrderCode & "', 'map', 'width=700,height=650');""></div>"

		'�]��
		If (sOrderType = "0" Or sOrderType = "2") And sCompanyKbn <> "4" Then
			'ؽ�̔h�����l�[ �܂��� �h����Ђ̋��l�[�̏ꍇ�͕\�����Ȃ�

			sTransfer = ChkStr(rRS.Collect("Transfer"))
			If sTransfer <> "" Then
				If sTransfer = "1" Then
					sTransfer = "����"
				Else
					sTransfer = "�Ȃ�"
				End If

				sWorkingPlace = sWorkingPlace & "<tr><td><img src=""/img/order/transfer.gif"" alt=""�]��""></td><td style=""padding-left:5px;""><p class=""m0"">" & sTransfer & "</p></td></tr>"
			End If
		End If

		flgWorkingPlace = True
		DspCondition = True
		'------------------------------------------------------------------------------
		'�Ζ��� end
		'******************************************************************************
	End If

	flgLine = False
%>
<h3>�Ζ�����</h3>
<%
	If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
%>
<div class="category1"><h4>�Ζ��`��</h4></div>
<div class="value1"><p class="m0"><%= GetWorkingType(rDB, rRS) %></p></div>
<div style="clear:both;"></div>
<%
	If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
%>
<div class="category1"><h4>�E��</h4></div>
<div class="value1">
	<p class="m0"><strong><%= sJobTypeDetail %></strong></p>
	<p class="m0"><%= GetJobType(rDB, rRS) %></p>
</div>
<div style="clear:both;"></div>
<%
	If flgSalary = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>���^</h4></div>
<div class="value1">
<%
		If sYearlyIncome <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>�N��</h5>
	<div class="value2"><p class="m0"><%= sYearlyIncome %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sMonthlyIncome <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>����</h5>
	<div class="value2"><p class="m0"><%= sMonthlyIncome %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sDailyIncome <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>����</h5>
	<div class="value2"><p class="m0"><%= sDailyIncome %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sHourlyIncome <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>����</h5>
	<div class="value2"><p class="m0"><%= sHourlyIncome %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sSalaryRemark <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>���^���l</h5>
	<div class="value2"><p class="m0"><%= sSalaryRemark %></p></div>
	<div style="clear:both; margin:0px;"></div>
<%
		End If

		If sTrafficFee <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>��ʔ�</h5>
	<div class="value2"><p class="m0"><%= sTrafficFee %></p></div>
	<div style="clear:both;"></div>
<%
		End If
%>
	<p class="m0" style="font-size:10px;">
		���Œ�z�͏����Ɋ֌W�Ȃ�������z�ł��B(�N���̍Œ�z�͏����Ɋ֌W�Ȃ������錎���̍��v�ł��B)
	</p>
</div>
<div style="clear:both;"></div>
<%
	End If

	If flgTime = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>����</h4></div>
<div class="value1">
<%
		If sWorkRange <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>�A�Ɗ���</h5>
	<div class="value2"><p class="m0"><%= sWorkRange %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sWorkingTime <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>�A�Ǝ���</h5>
	<div class="value2"><p class="m0"><%= sWorkingTime %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sWorkTimeRemark <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>�A�Ǝ��Ԕ��l</h5>
	<div class="value2"><p class="m0"><%= sWorkTimeRemark %></p></div>
	<div style="clear:both;"></div>
<%
		End If
%>
</div>
<div style="clear:both;"></div>
<%
	End If

	If flgHoliday = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>�x��</h4></div>
<div class="value1">
<%
		If sWeeklyHolidayType <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>�x��</h5>
	<div class="value2"><p class="m0"><%= sWeeklyHolidayType %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sHolidayRemark <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>�x�����l</h5>
	<div class="value2"><p class="m0"><%= sHolidayRemark %></p></div>
	<div style="clear:both;"></div>
<%
			sClearSolid = ""
		End If
%>
</div>
<div style="clear:both;"></div>
<%
	End If

	If flgWorkingPlace = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>�Ζ���</h4></div>
<div class="value1">
<%
		If sWPSection <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>������</h5>
	<div class="value2"><p class="m0"><%= sWPSection %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sWPTel <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>�d�b�ԍ�</h5>
	<div class="value2"><p class="m0"><%= sWPTel %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sWPAddress <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>�Ζ��n</h5>
	<div class="value2"><p class="m0"><%= sWPAddress %></p><%= sMAP %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sNearbyStation <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>�Ŋ�w</h5>
	<div class="value2"><%= sNearbyStation %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sNearbyRailway <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>����</h5>
	<div class="value2"><%= sNearbyRailway %></div>
	<div style="clear:both;"></div>
<%
		End If
%>
</div>
<div style="clear:both;"></div>
<%
	End If

	If DspCondition = True Then Response.Write "<br>"
End Function














'******************************************************************************
'�T�@�v�F���l�[�̕K�v�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail_entity.asp
'******************************************************************************
Function DspNeedCondition(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode			'���R�[�h
	Dim sCompanyCode		'��ƃR�[�h
	Dim sOrderType			'���l�[���
	Dim sCompanyKbn			'��Ƌ敪
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

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'��ƃR�[�h start
		'------------------------------------------------------------------------------
		sOrderCode = rRS.Collect("OrderCode")
		sCompanyCode = rRS.Collect("CompanyCode")
		sOrderType = rRS.Collect("OrderType")
		sCompanyKbn = rRS.Collect("CompanyKbn")
		'------------------------------------------------------------------------------
		'��ƃR�[�h end
		'******************************************************************************

		'******************************************************************************
		'�N�� start
		'------------------------------------------------------------------------------
		sAge = ""
		sAgeMin = ChkStr(rRS.Collect("AgeMin"))
		sAgeMax = ChkStr(rRS.Collect("AgeMax"))
		sAgeReasonFlag = ChkStr(rRS.Collect("AgeReasonFlag"))
		sAgeReason = ChkStr(rRS.Collect("AgeReason"))
		sAgeReasonDetail = Replace(ChkStr(rRS.Collect("AgeReasonDetail")), vbCrLf, "<br>")

		If sAgeReasonFlag = "0" Or sAgeReasonFlag = "" Or (sAgeMin & sAgeMax = "") Then
			sAge = "�N��s��<br>"
			sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[�H]�����ɂ���</a>"
		ElseIf sOrderType = "1" Or (sOrderType = "0" And sCompanyKbn = "4") Then
			sAge = "�h���Č��̂��߁A�N��f�ڂ��Ă��܂���B<br>"
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
		sSkillOther = GetSkill(rDB, rRS, "OtherSkill")
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
	End If

	flgLine = False
%>
<h3>�K�v����</h3>
<%
	If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
%>
<div class="category1"><h4>�N��</h4></div>
<div class="value1"><p class="m0"><%= sAge %></p></div>
<div style="clear:both;"></div>
<%
	If sFEHistory <> "" Then
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>��]�w��</h4></div>
<div class="value1"><p class="m0"><%= sFEHistory %></p></div>
<div style="clear:both;"></div>
<%
	End If

	'******************************************************************************
	'���i�o�� start
	'------------------------------------------------------------------------------
	sClearSolid = " style=""border-top-width:0px;"""
	If flgLicense = True Then
		flgLine2 = False
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>���i</h4></div>
<div class="value1">
<%
		If sLicense <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>���i</h5>
	<div class="value2"><%= sLicense %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sLicenseOther <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>���̑����i</h5>
	<div class="value2"><p class="m0"><%= sLicenseOther %></p></div>
	<div style="clear:both;"></div>
<%
		End If
%>
</div>
<div style="clear:both;"></div>
<%
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
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>�X�L��</h4></div>
<div class="value1">
<%
		If sSkillOS <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>�n�r</h5>
	<div class="value2"><%= sSkillOS %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sSkillApp <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>���ع����</h5>
	<div class="value2"><%= sSkillApp %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sSkillDL <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>�J������</h5>
	<div class="value2"><%= sSkillDL %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sSkillDB <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>�f�[�^�x�[�X</h5>
	<div class="value2"><%= sSkillDB %></div>
	<div style="clear:both;"></div>
<%
		End If

		If sSkillOther <> "" Then
'			If flgLine2 = True Then Response.Write "<table class=""odline2"" border=""0""><tr><td></td></tr></table>"
'			flgLine2 = True
%>
	<h5>���̑��X�L��</h5>
	<div class="value2"><%= sSkillOther %></div>
	<div style="clear:both;"></div>
<%
		End If
%>
</div>
<div style="clear:both;"></div>
<%
	End If
	'------------------------------------------------------------------------------
	'�X�L���o�� end
	'******************************************************************************

	'******************************************************************************
	'���̑����L���� start
	'------------------------------------------------------------------------------
	If sOtherNote <> "" Then
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>���L����</h4></div>
<div class="value1"><p class="m0"><%= sOtherNote %></p></div>
<div style="clear:both;"></div>
<%
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
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/company_order.asp
'******************************************************************************
Function DspHowToEntry(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sOrderCode			'���R�[�h
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
	Dim flgEntryInfo		'�����񂪗L�邩������ [True]���� [False]�Ȃ�
	Dim flgProcess			'�I�l�菇���L�邩������ [True]���� [False]�Ȃ�
	Dim sClearSolid
	Dim flgLine				'�������t���O

	DspHowToEntry = False

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'��ƃR�[�h start
		'------------------------------------------------------------------------------
		sOrderType = ChkStr(rRS.Collect("OrderType"))
		sOrderCode = ChkStr(rRS.Collect("OrderCode"))
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
		'��ƃR�[�h start
		'------------------------------------------------------------------------------
		sCSectionName = ChkStr(rRS.Collect("LisDepartment"))
		sCPersonName = ChkStr(rRS.Collect("EmployeeName"))
		sCTel = ChkStr(rRS.Collect("LisTelephoneNumber"))
		sLis = sCPersonName & "�m���X�������" & sCSectionName & "�n�@" & sCTel & "<br>(���̈Č��̓��X������Ђ����܂Ƃ߂Ă��܂��B)"
		DspHowToEntry = True
		'------------------------------------------------------------------------------
		'��ƃR�[�h end
		'******************************************************************************

	End If


	flgLine = False
%>
<h3>������</h3>
<%
	If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
%>
<div class="category1"><h4>���R�[�h</h4></div>
<div class="value1"><p class="m0"><%= sOrderCode %></p></div>
<div style="clear:both;"></div>
<%
	If flgEntryInfo = True Then
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>������@</h4></div>
<div class="value1"><p class="m0"><%= sEntryInfo %></p></div>
<div style="clear:both;"></div>
<%
	End If

	If flgProcess = True Then
		If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>�I�l�菇</h4></div>
<div class="value1">
<%
		If sProcess1 <> "" Then
%>
	<h5>�X�e�b�v�P</h5>
	<div class="value2"><p class="m0"><%= sProcess1 %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sProcess2 <> "" Then
%>
	<h5>�X�e�b�v�Q</h5>
	<div class="value2"><p class="m0"><%= sProcess2 %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sProcess3 <> "" Then
%>
	<h5>�X�e�b�v�R</h5>
	<div class="value2"><p class="m0"><%= sProcess3 %></p></div>
	<div style="clear:both;"></div>
<%
		End If

		If sProcess4 <> "" Then
%>
	<h5>�X�e�b�v�S</h5>
	<div class="value2"><p class="m0"><%= sProcess4 %></p></div>
	<div style="clear:both;"></div>
<%
		End If
%>
</div>
<div style="clear:both;"></div>
<%
	End If

	If DspHowToEntry = True Then Response.Write "<br>"
End Function

'******************************************************************************
'�T�@�v�F���l�[�̒S���ҘA������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/11
'���@�l�F
'�g�p���F�����ƃi�r/order/company_order.asp
'******************************************************************************
Function DspContact(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
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
	Dim flgLine				'�������t���O

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'��ƃR�[�h start
		'------------------------------------------------------------------------------
		sCompanyCode = rRS.Collect("CompanyCode")
		sOrderType = rRS.Collect("OrderType")
		If sOrderType <> "0" Then Exit Function
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

		Call SetOrderCompanyName(sCompanyName, sCompanyNameF, sOrderType, sCompanyKbn, sCompanySpeciality)
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

			If sCompanyKbn = "2" Or sCompanyKbn = "4" Then
				'�l�މ�Ђ̋��l�[�̏ꍇ�́u���O�v�{�u�l�މ�Ж��v
				sPerson = sCPersonName & "(" & sCompanyName & ")"
			Else
				'��ʊ�Ƃ̋��l�[�̏ꍇ�́u���O�v�{�u�J�i�v
				sPerson = sCPersonName
				If sCPersonNameF <> "" Then sPerson = sPerson & "(" & sCPersonNameF & ")"
			End If
'		Else
'			'���X�󒍕[�̏ꍇ�́u���X�S���Җ��v�{�u���X�S���҃J�i�v
'			sCSectionName = ChkStr(rRS.Collect("LisDepartment"))
'			sCPersonName = ChkStr(rRS.Collect("EmployeeName"))
'			sCTel = ChkStr(rRS.Collect("LisTelephoneNumber"))
'			sPerson = sCPersonName
'			If sPerson <> "" Then sPerson = sPerson & "(�l�މ�ЁF���X�������)"
		End If

		sContact = ""
		If sCTel <> "" Then sContact = sContact & sCTel & "	<SPAN style='font-size:10px;'>�@���d�b���ł̂��₢���킹�̍ہA�u�����ƃi�r�������v�ƌ����ƃX���[�Y�ł��B</SPAN>"
		If sContact <> "" Then sContact = sContact & "<br>"
		If sCMail <> "" Then sContact = sContact & sCMail
		'------------------------------------------------------------------------------
		'�d���̘A����
		'******************************************************************************
	End If

	flgLine = False
%>
<h3 class="sp">�S���ҘA����</h3>

<%
	If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
%>
<div class="category1"><h4>�S����</h4></div>
<div class="value1"><p class="m0"><%= sPerson %></p></div>
<div style="clear:both;"></div>
<%
	If sCSectionName <> "" Then
		If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
%>
<div class="category1"><h4>�S������</h4></div>
<div class="value1"><p class="m0"><%= sCSectionName %></p></div>
<div style="clear:both;"></div>
<%
	End If

	If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
%>
<div class="category1"><h4>�A����</h4></div>

<div class="value1"><p class="m0"><%= sContact %></p></div>
<div style="clear:both;"></div>
<br>
<%
End Function

'******************************************************************************
'�T�@�v�F���X�̈Č��S���ҁA�R���T���������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
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

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sCompanyCode = rRS.Collect("CompanyCode")
	sOrderType = rRS.Collect("OrderType")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

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
%>
<h3 class="sp"><%= sTitle %></h3>
<div class="category1"><h4>�R���T���^���g</h4></div>
<div class="value1"><p class="m0"><%= sConsultantLink %></p></div>
<div style="clear:both;"></div>
<%
	Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
%>
<div class="category1"><h4>�S������</h4></div>
<div class="value1"><p class="m0"><%= sBranchName %></p></div>
<div style="clear:both;"></div>
<%
	Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
%>
<div class="category1"><h4>�A����</h4></div>
<div class="value1"><p class="m0"><%= sTel %><SPAN style='font-size:10px;'>�@�����₢���킹�̍ہA��L�u���R�[�h�v�Ɓu�����ƃi�r�������v�ƌ����ƃX���[�Y�ł��B</SPAN></p>	</div>
<div style="clear:both;"></div>
<%
	If sComment <> "" Then
		Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
%>
<div class="category1"><h4>����</h4></div>
<div class="value1"><p class="m0"><%= sComment %></p></div>
<div style="clear:both;"></div>
<br>
<%
	End If
End Function

'******************************************************************************
'�T�@�v�F�ŐV���[�����o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
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
		sSQL = "sp_GetDataMailHistory '" & vUserID & "', '" & rRS.Collect("CompanyCode") & "', '" & rRS.Collect("OrderCode") & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			sDateTime = GetDateStr(oRS.Collect("SendDay"), "/") & "�@" & GetTimeStr(oRS.Collect("SendDay"), ":")
			sSubject = ChkStr(oRS.Collect("Subject"))
			sDetail = Replace(ChkStr(oRS.Collect("Body")), vbCrLf, "<br>")
			sDetail = Replace(sDetail, vbCr, "<br>")
			sDetail = Replace(sDetail, vbLf, "<br>")
%>
<h3 class="sp">�ŐV�̑��M�ς݃��[��</h3>
<%
			If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>���M����</h4></div>
<div class="value1"><p class="m0"><%= sDateTime %></p></div>
<div style="clear:both;"></div>
<%
			If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>�T�u�W�F�N�g</h4></div>
<div class="value1"><p class="m0"><%= sSubject %></p></div>
<div style="clear:both;"></div>
<%
			If flgLine = True Then Response.Write "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>���e</h4></div>
<div class="value1"><p class="m0"><%= sDetail %></p></div>
<div style="clear:both;"></div>
<br>
<%
		End If
	End If

	Call RSClose(oRS)

	DspNewMail = True
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̋Ζ��`�ԕ���
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
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
			sWorkingType = sWorkingType & "�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=300,height=200"")'>�l�ޏЉ�</a>�z" 
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
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
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
		sJobType = sJobType & oRS.Collect("JobTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then sJobType = sJobType & "<br>"
	Loop
	Call RSClose(oRS)

	GetJobType = sJobType
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̋Ζ��`�ԕ���
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
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
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�쐬�ҁFLis Kokubo
'�쐬���F2006/05/08
'���@�l�F
'�g�p���F
'******************************************************************************
Function GetNearbyStation(ByRef rDB, ByRef rRS)
	Const STATIONCOL = 2

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim idx
	Dim sStation
	Dim sToStation
	Dim iStation

	If GetRSState(rRS) = False Then Exit Function

	iStation = 0
	sStation = ""
	sSQL = "sp_GetDataNearbyStation '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iStation = iStation + 1

		sToStation = ""
		If ChkStr(oRS.Collect("ToStationTime")) <> "" Then sToStation = oRS.Collect("ToStationTime") & "��"
		If ChkStr(oRS.Collect("ToStationRemark")) <> "" Then sToStation = oRS.Collect("ToStationRemark") & sToStation
		If sToStation <> "" Then sToStation = "(" & sToStation & ")"

		sStation = sStation & "<p style=""width:50%; float:left;"">" & oRS.Collect("StationName") & "�w" & sToStation & "</p>"
		If iStation Mod STATIONCOL = 0 Then sStation = sStation & "<br clear=""all"">"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'���r���[�ŏI������ꍇ�̒���
	If sStation <> "" And iStation Mod STATIONCOL <> 0 Then sStation = sStation & "<br clear=""all"">"

	GetNearbyStation = sStation
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̍Ŋ񉈐�����
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�쐬�ҁFLis Kokubo
'�쐬���F2006/05/08
'���@�l�F
'�g�p���F
'******************************************************************************
Function GetNearbyRailway(ByRef rDB, ByRef rRS)
	Const RAILWAYCOL = 2

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim idx
	Dim sRailway
	Dim iRailway

	If GetRSState(rRS) = False Then Exit Function

	iRailway = 0
	sRailway = ""
	sSQL = "sp_GetDataNearbyRailwayLine '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		iRailway = iRailway + 1

		sRailway = sRailway & "<p style=""width:50%; float:left;"">" & oRS.Collect("RailwayLineName2") & "</p>"
		If iRailway Mod RAILWAYCOL = 0 Then sRailway = sRailway & "<br clear=""all"">"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'���r���[�ŏI������ꍇ�̒���
	If sRailway <> "" And iRailway Mod RAILWAYCOL <> 0 Then
		sRailway = sRailway & "<br clear=""all"">"
	End If

	GetNearbyRailway = sRailway
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̃X�L������
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
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
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
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
'�@�@�@�FrRS			�Fsp_GetDetailOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
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
Function GetOrderTitle(ByRef rDB, ByVal vOrderCode, ByRef rTitle, ByRef rDescription)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sWorkingType

	sSQL = "up_GetOrderTitle '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		rTitle = ChkStr(oRS.Collect("JobTypeDetail"))
		rDescription = ChkStr(oRS.Collect("BusinessDetail"))
	End If
	Call RSClose(oRS)

	sSQL = "sp_GetListWorkingType '" & vOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	sWorkingType = ""
	Do While GetRSState(oRS) = True
		If sWorkingType <> "" Then sWorkingType = sWorkingType & ","
		sWorkingType = sWorkingType & oRS.Collect("WorkingTypeName")
		oRS.MoveNext
	Loop
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
'�X�@�V�F
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
'�@�@�@�F[OUTPUT]rJobTypeDetail		�F��̓I�E�햼
'�@�@�@�F[OUTPUT]rCompanyName		�F��Ɩ�
'�@�@�@�F[OUTPUT]rImg				�F��ƃC���[�W
'�@�@�@�F[OUTPUT]rWorkingTypeIcon	�F�Ζ��`�ԃA�C�R��
'�@�@�@�F[OUTPUT]rWorkingPlace		�F�Ζ��n
'�@�@�@�F[OUTPUT]rStation			�F�Ŋ�w
'�@�@�@�F[OUTPUT]rYearlyIncome		�F�N��
'�@�@�@�F[OUTPUT]rMonthlyIncome		�F����
'�@�@�@�F[OUTPUT]rDailyIncome		�F����
'�@�@�@�F[OUTPUT]rHourlyIncome		�F����
'�߂�l�F
'�쐬���F2007/05/31
'�쐬�ҁFLis Kokubo
'���@�l�F
'�X�@�V�F
'******************************************************************************
Function GetRecommendValues(ByRef rDB, ByRef rRS, ByVal vRCMD, ByRef rJobTypeDetail, ByRef rCompanyName, ByRef rImg, ByRef rWorkingTypeIcon, ByRef rWorkingPlace, ByRef rStation, ByRef rYearlyIncome, ByRef rMonthlyIncome, ByRef rDailyIncome, ByRef rHourlyIncome)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim sOrderCode			'���R�[�h
	Dim sCompanyCode		'��ƃR�[�h
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
	Dim sWorkingPlace		'�Ζ��n
	Dim sStation			'�Ŋ�w
	Dim sImg				'�摜URL

	Dim sURL				'���l�[�ڍׂ�URL
	Dim sAlign				'�g�� [vCols = 1]left [vCols = vMaxCols]right [����ȊO]center

	If GetRSState(rRS) = False Then Exit Function

	sURL = HTTP_CURRENTURL & "order/order_detail.asp"

	sSQL = "sp_GetDetailOrder '" & rRS.Collect("OrderCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	'���R�[�h
	sOrderCode = ChkStr(oRS.Collect("OrderCode"))
	'��ƃR�[�h
	sCompanyCode = ChkStr(oRS.Collect("CompanyCode"))
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
			Case "001": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/haken.gif"" alt=""�h��"" style=""margin-right:1px;"">"
			Case "002": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/seishain.gif"" alt=""���Ј�"" style=""margin-right:1px;"">"
			Case "003": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/keiyaku.gif"" alt=""�_��Ј�"" style=""margin-right:1px;"">"
			Case "004": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/syoha.gif"" alt=""�Љ�\��h��"" style=""margin-right:1px;"">"
			Case "005": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/arbeit.gif"" alt=""�A���o�C�g�E�p�[�g"" style=""margin-right:1px;"">"
			Case "006": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/soho.gif"" alt=""SOHO"" style=""margin-right:1px;"">"
			Case "007": sWorkingTypeIcon = sWorkingTypeIcon & "<img src=""/img/fc.gif"" alt=""FC"" style=""margin-right:1px;"">"
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
	'sImg = "<img src=""" & sImg & """ alt=""" & sCompanyName & """ width=""156"" height=""117"">"
	sImg = "<img src=""" & sImg & """ alt=""" & sCompanyName & """ width=""88"" height=""66"" border=""0"" align=""left"" style=""margin:0px; padding:0px;"">"
	'------------------------------------------------------------------------------
	'�摜 end
	'******************************************************************************

	'******************************************************************************
	'�Ζ��n start
	'------------------------------------------------------------------------------
	sWorkingPlace = ""
	If sOrderType = "0" Then
		sWorkingPlace = ChkStr(oRS.Collect("WorkingPlaceAddressAll"))
	Else
		sWorkingPlace = ChkStr(oRS.Collect("WorkingPlacePrefectureName")) & ChkStr(oRS.Collect("WorkingPlaceCity"))
	End If
	'------------------------------------------------------------------------------
	'�Ŋ�w end
	'******************************************************************************

	'******************************************************************************
	'�Ŋ�w start
	'------------------------------------------------------------------------------
	sStation = ""
	sSQL = "sp_GetDataNearbyStation '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
	Do While GetRSState(oRS2) = True
		sStation = sStation & GetStrNearbyStation(oRS2.Collect("StationName"), oRS2.Collect("ToStationTime"), oRS2.Collect("ToStationRemark"))
		oRS2.MoveNext
		If GetRSState(oRS2) = True Then sStation = sStation & "<br>"
	Loop
	'------------------------------------------------------------------------------
	'�Ŋ�w end
	'******************************************************************************

	rJobTypeDetail = "<a href=""" & sURL & "?ordercode=" & sOrderCode & "&amp;rcmd=" & vRCMD & """>" & sViewJobTypeDetail & "</a>"
	rCompanyName = sCompanyName
	rImg = "<a href=""" & sURL & "?ordercode=" & sOrderCode & "&amp;rcmd=" & vRCMD & """>" & sImg & "</a>"
	rWorkingTypeIcon = sWorkingTypeIcon
	rWorkingPlace = sWorkingPlace
	rStation = sStation
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
'�쐬�ҁFLis Kokubo
'�쐬���F2007/02/14
'�߂�l�F
'�@�@�@�F
'�g�p���F�����ƃi�r/order/order_detail.asp
'���@�l�F
'******************************************************************************
Function GetImgOrderSpeciality(ByRef rDB, ByRef rRS)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sWorkingCode
	Dim sOrderType
	Dim sCompanyKbn

	If GetRSState(rRS) = False Then Exit Function

	sOrderType = rRS.Collect("OrderType")
	sCompanyKbn = rRS.Collect("CompanyKbn")

	GetImgOrderSpeciality = ""
	'�A�N�Z�X����100�𒴂��Ă���΁uHOT�v�\���i���X�����j
	If rRS.Collect("AccessCount") > 100 Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/c_HOT_green.gif"" alt=""�l�C"" width=""50"" height=""15"">&nbsp;"
	End If

	'UPDATE�ƍ�������10�����������Łu�V���v�\��(���X����)
	If rRS.Collect("Updateday") > NOW()-10 Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/c_NEW_green.gif"" alt=""�V��"" width=""50"" height=""15"">&nbsp;"
	End If

	'���o���҂n�j�̏ꍇ�A�킩�΃}�[�N�\��(���X����)
	If rRS.Collect("InexperiencedPersonFlag") = "1" Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/no_experience.gif"" alt=""���o���ҁ^���V�����}"" width=""50"" height=""15"">&nbsp;"
	End If

	'�t�^�[���E�h�^�[��
	If rRS.Collect("UITurnFlag") = "1" Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/ui_turn.gif"" alt=""�t�^�[���E�h�^�[��"" width=""50"" height=""15"">&nbsp;"
	End If

	'��w���������d��
	If rRS.Collect("UtilizeLanguageFlag") = "1" Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/linguistic_job.gif"" alt=""��w���������d��"" width=""50"" height=""15"">&nbsp;"
	End If

	'�N�ԋx��120���ȏ�
	If rRS.Collect("ManyHolidayFlag") = "1" Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/year_holidaycnt.gif"" alt=""�N�ԋx��120���ȏ�"" width=""50"" height=""15"">&nbsp;"
	End If

	'�t���b�N�X�^�C�����x���� ------2006/01/10 Hayashi ADD
	If rRS.Collect("FlexTimeFlag") = "1" And sOrderType = "0" And sCompanyKbn = "1" Then
		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/flextime.gif"" alt=""�t���b�N�X�^�C�����x����"" width=""50"" height=""15"">&nbsp;"
	End If

'����Yahoo!�̌������炨�d�����ڍ׃y�[�W�֗���l�փA�C�R���\��
if G_FLGRESUME = False Then
	if InStr(Request.ServerVariables("HTTP_REFERER"),"search.yahoo.co.jp/") <> 0 Then

	sSQL = "sp_GetDataWorkingType '" & sOrderCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sWorkingcode = oRS.Collect("WorkingTypecode")

		GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/order_detail_icon/icon_w" & sWorkingcode & ".gif"" alt=""�h���Ј�"" width=""50"" height=""15"">&nbsp;"

		oRS.MoveNext
	Loop

	GetImgOrderSpeciality = GetImgOrderSpeciality & "<img src=""/img/order_detail_icon/icon_p" & rRS.Collect("Workingplaceprefecturecode") & ".gif"" alt=""�k�C��"" width=""50"" height=""15"">&nbsp;"
	End if
End if
'/����Yahoo!�̌������炨�d�����ڍ׃y�[�W�֗���l�փA�C�R���\��

	If GetImgOrderSpeciality <> "" Then GetImgOrderSpeciality = "<div>" & GetImgOrderSpeciality & "</div>"

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
	<div style="float:right; width:150px;"><a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= vOrderCode %>"><img src="/img/order/btn_reg_button1.gif" alt="�������o�^���ĉ���" border="0"></a></div>
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
'�쐬�ҁFLis Kokubo
'�쐬���F2007/05/08
'���@�l�F
'�g�p���Forder/order_detail_entity.asp
'******************************************************************************
Function AccessHistoryOrder(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	If vUserType = "staff" Then
		sSQL = "up_Reg_LOG_AccessHistoryOrder '" & vOrderCode & "', '" & vUserID & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Call RSClose(oRS)
	ElseIf IsRE(Request.Cookies("id_memory"), "^S\d\d\d\d\d\d\d$", True) = True Then
		sSQL = "up_Reg_LOG_AccessHistoryOrder '" & vOrderCode & "', '" & Request.Cookies("id_memory") & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Call RSClose(oRS)
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
	If GetRSState(oRS2) = True Then
		AccessCountUp = oRS.Collect("AccessCount")
	End If
	Call RSClose(oRS)
End Function

'*******************************************************************************
'�T�@�v�F�S�p���p����������������̃o�C�g���𐳊m�ɕԂ�(Web����̈��p)
'���@���Fstring		:�Ώە�����
'�߂�l�FInterger	:�Ώە�����̃o�C�g��
'�쐬���F2007/05/23 Lis Sotome
'�X�@�V�F
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
'�X�@�V�F
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
