<%
'**********************************************************************************************************************
'�T�@�v�F���[���ڍ׃y�[�W /staff/mailhistory_person_entity.asp
'�@�@�@�F��L�y�[�W�ŏo�͗p�̊֐��Q�����̃t�@�C���ɗp�ӂ���B
'�@�@�@�F
'�@�@�@�F�������@�O������@������
'�@�@�@�F�v���O�C���N���[�h
'�@�@�@�F/config/personel.asp
'�@�@�@�F/include/commonfunc.asp
'��@���F�������@���[���ڍ׃y�[�W�o�͗p�@������
'�@�@�@�FDspNoticeMailLink
'�@�@�@�FDspMailReturnBtn	�F�ԐM�{�^���o��
'�@�@�@�FDspMailDetail		�F���[���ڍׂ��o��
'�@�@�@�FDspNoMailDetail	�F���[���������ꍇ�̕����o��
'�@�@�@�F�������@���[���ڍ׃y�[�W�X�V�n�@������
'�@�@�@�FRegNoticeScoutMailUnRead_OpenFlag	�F���J���ʒm���[�����O�̊J���t���O�𗧂Ă鏈��
'**********************************************************************************************************************

'******************************************************************************
'�T�@�v�F�X�P�W���[���ʒm�T�[�r�X�����N���o��
'���@���FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�@�@�@�FvAnswerNG	�F�ԐM�� ["0"]�ԐM�� ["1"]�ԐM�s��
'���@�l�F
'�g�p���F/staff/maildetail_person.asp
'�X�@�V�F2007/03/02 LIS K.Kokubo 
'******************************************************************************
Function DspNoticeMailLink(ByVal vMode, ByVal vAnswerNG)
	If vMode <> "1" And vAnswerNG = "0" Then
		Response.Write "<div style=""padding:0px;"">"
		Response.Write "<input type=""button"" value=""�X�P�W���[���ʒm�T�[�r�X�ɓo�^"" style=""width:190px; color:#aa3300;"" onclick=""window.open('" & HTTPS_NAVI_CURRENTURL & "staff/notification_mail_service.asp?popup=1','notification_mail_service','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=630');return false;"">"
		Response.Write "<span style=""font-size:10px;"">...�ʐڂȂǂ̓��������܂�����A�X�P�W���[���ʒm�T�[�r�X�ɓo�^���ĖY�p�h�~�I</span>"
		Response.Write "</div>"
		Response.Write "<hr size=""1"">"
	End If
End Function

'******************************************************************************
'�T�@�v�F�ԐM�{�^�����o��
'���@���FrRS		�F���[���ڍׂ̃��R�[�h�Z�b�g(up_GetDetailMail)
'�@�@�@�FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�@�@�@�FvAnswerNG	�F�ԐM�� ["0"]�ԐM�� ["1"]�ԐM�s��
'���@�l�F
'�g�p���F�i�r/staff/maildetail_person_entity.asp
'���@���F2007/03/02 LIS K.Kokubo �쐬
'�@�@�@�F2009/03/27 LIS K.Kokubo up_ChkScoutMail��ReceiveFlag��ǉ������Ή�
'******************************************************************************
Function DspMailReturnBtn(ByRef rRS, ByVal vMode, ByVal vAnswerNG)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbID
	Dim dbScoutMailFlag
	Dim dbReceiveFlag

	If GetRSState(rRS) = False Then Exit Function

	dbID = rRS.Collect("ID")
	dbScoutMailFlag = "0"
	dbReceiveFlag = "0"

	sSQL = "EXEC up_ChkScoutMail '" & dbID & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbScoutMailFlag = ChkStr(oRS.Collect("ScoutMailFlag"))
		dbReceiveFlag = ChkStr(oRS.Collect("ReceiveFlag"))
	End If
	Call RSClose(oRS)

	If vMode <> "1" Then
		If vAnswerNG = "0" Then
			Response.Write "<div align=""center"" style=""padding:5px 0px;"">"
			If dbScoutMailFlag = "0" Or dbReceiveFlag = "1" Then
				Response.Write "<input type=""button"" value=""�ԁ@�M"" onclick=""SendAnswer();""><br><span style=""font-size:10px;"">����Ƃփ��[����ԐM���܂��B</span>"
			Else
				Response.Write "<fieldset style=""width:175px; float:left;"">"
				Response.Write "<legend>���[���{������͂��ĕԐM</legend>"
				Response.Write "<input type=""button"" value=""�ԁ@�M"" onclick=""SendAnswer();""><br>"
				Response.Write "<p class=""m0"" style=""font-size:10px;"">�����[����ԐM���܂��B(����,����Ȃ�)</p>"
				Response.Write "</fieldset>" & vbCrLf

				Response.Write "<fieldset style=""width:375px; float:right;"">"
				Response.Write "<legend>�J���^���ԐM(���[���̍쐬�͕s�v)</legend>"
				Response.Write "<div style=""float:left; width:50%;"">"
				Response.Write "<form action=""mail_detail_person.asp?id=" & dbID & """ method=""post"" onsubmit=""return confirm('�����ۗ��i�����j���܂����H');"">"
				Response.Write "<input name=""frmmailtype"" type=""hidden"" value=""1"">"
				Response.Write "<input type=""submit"" value=""�ہ@��""><br>"
				Response.Write "<p class=""m0"" style=""font-size:10px;"">�������ۗ�(����)���܂��B[<span style=""color:#0045f9; cursor:pointer;"" onclick=""window.open('/infomation/horyubutton.asp','autologin','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=400,height=220');""><u>�H</u></span>]</p>"
				Response.Write "</form>"
				Response.Write "</div>" & vbCrLf

				Response.Write "<div style=""float:left; width:50%;"">"
				Response.Write "<form action=""mail_detail_person.asp?id=" & dbID & """ method=""post"" onsubmit=""return confirm('�X�J�E�g�����ނ��܂����H');"">"
				Response.Write "<input name=""frmmailtype"" type=""hidden"" value=""2"">"
				Response.Write "<input type=""submit"" value=""���@��""><br>"
				Response.Write "<p class=""m0"" style=""font-size:10px;"">���X�J�E�g�����ނ��܂��B</p>"
				Response.Write "</form>"
				Response.Write "</div>" & vbCrLf
				Response.Write "<div style=""clear:both;""></div>"
				Response.Write "</fieldset>"
			End If
			Response.Write "<div style=""clear:both;""></div>"
			Response.Write "</div>"
		Else
			Response.Write "<b>���̋��l�[�͌f�ڂ��I�����Ă���A�A������邱�Ƃ��ł��܂���B</b><br><br>"
		End If
	End If
End Function

'******************************************************************************
'�T�@�v�F���[���ڍׂ��o��
'���@���FvMode		�F���݂̕\�����[�h ["0"]��M���[���ꗗ ["1"]���M���[���ꗗ
'�@�@�@�FvAnswerNG	�F�ԐM�� ["0"]�ԐM�� ["1"]�ԐM�s��
'���@�l�F
'�g�p���F�i�r/staff/maildetail_person_entity.asp
'�X�@�V�F2007/03/02 LIS K.kokubo �쐬
'�@�@�@�F2008/05/08 LIS K.Kokubo ���W�b�N�ύX
'�@�@�@�F2008/08/20 Lis �� �����t���O�̒ǉ��ƃt���b�N�X�ړ�
'******************************************************************************
Function DspMailDetail(ByRef rRS, ByVal vMode, ByVal vAnswerNG)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode	'���R�[�h
	Dim dbCompanyCode	'��ƃR�[�h
	Dim dbWorkingPlacePrefectureCode
	Dim dbWorkingPlacePrefectureName

	Dim sJobTypeDetail
	Dim sAccessCount
	Dim sUpdateDay
	Dim sInexperiencedPersonFlag
	Dim sUITurnFlag
	Dim sUtilizeLanguageFlag
	Dim sManyHolidayFlag
	Dim sFlexTimeFlag
	'**TOP 08/08/19 Lis�� ADD
	Dim sNearStationFlag,sNoSmokingFlag,sNewlyBuiltFlag,sLandmarkFlag
	Dim sRenovationFlag,sDesignersFlag,sCompanyCafeteriaFlag,sShortOvertimeFlag,sMaternityFlag
	Dim sDressFreeFlag,sMammyFlag,sFixedTimeFlag,sShortTimeFlag,sHandicappedFlag
	'**BTM 08/08/19 Lis�� ADD
	Dim sOrderType
	Dim sCompanyKbn
	Dim sImgOrderState
	Dim sWorkingPlacePrefectureCode
	Dim sWorkingcode
	Dim sWorkingname
	Dim sWorkingPlacePrefectureName

	Dim dbSecretFlag
	Dim sYearlyIncomeMin
	Dim sYearlyIncomeMax
	Dim sMonthlyIncomeMin
	Dim sMonthlyIncomeMax
	Dim sDailyIncomeMin
	Dim sDailyIncomeMax
	Dim sDailyIncome
	Dim sHourlyIncomeMin
	Dim sHourlyIncomeMax
	Dim sYearlyIncome
	Dim sMonthlyIncome
	Dim sHourlyIncome
	Dim sImgMain
	Dim sImgSub
	Dim sCompanyPictureFlag
	Dim flgImg
	Dim idx

	'��̓I�E�햼�̎擾
	dbOrderCode = rRS.Collect("OrderCode")
	sSQL = "select OrderType,JobTypeDetail,AccessCount,UpdateDay,Companycode from C_info where ordercode = '" & dbOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sJobTypeDetail = oRS.Collect("JobTypeDetail")
		sAccessCount = oRS.Collect("AccessCount")
		sUpdateDay = oRS.Collect("UpdateDay")
		dbCompanyCode = oRS.Collect("Companycode")
		sOrderType = oRS.Collect("OrderType")
	End if

	'**TOP 08/08/19 Lis�� REP
	'sSQL = "select InexperiencedPersonFlag,UITurnFlag,UtilizeLanguageFlag,ManyHolidayFlag from C_SupplementInfo where ordercode = '" & dbOrderCode & "'"
	sSQL = "select InexperiencedPersonFlag,UITurnFlag,UtilizeLanguageFlag,ManyHolidayFlag"
	sSQL = sSQL & ",FlexTimeFlag,NearStationFlag,NoSmokingFlag,NewlyBuiltFlag,LandmarkFlag"
	sSQL = sSQL & ",RenovationFlag,DesignersFlag,CompanyCafeteriaFlag,ShortOvertimeFlag,MaternityFlag"
	sSQL = sSQL & ",DressFreeFlag,MammyFlag,FixedTimeFlag,ShortTimeFlag,HandicappedFlag"
	sSQL = sSQL & " from C_SupplementInfo where ordercode = '" & dbOrderCode & "'"
	'**BTM 08/08/19 Lis�� REP
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sInexperiencedPersonFlag = oRS.Collect("InexperiencedPersonFlag")
		sUITurnFlag = oRS.Collect("UITurnFlag")
		sUtilizeLanguageFlag = oRS.Collect("UtilizeLanguageFlag")
		sManyHolidayFlag = oRS.Collect("ManyHolidayFlag")
		'**TOP 08/08/19 Lis�� REP
		sFlexTimeFlag = oRS.Collect("FlexTimeFlag")
		sNearStationFlag = oRS.Collect("NearStationFlag")
		sNoSmokingFlag = oRS.Collect("NoSmokingFlag")
		sNewlyBuiltFlag = oRS.Collect("NewlyBuiltFlag")
		sLandmarkFlag = oRS.Collect("LandmarkFlag")
		sRenovationFlag = oRS.Collect("RenovationFlag")
		sDesignersFlag = oRS.Collect("DesignersFlag")
		sCompanyCafeteriaFlag = oRS.Collect("CompanyCafeteriaFlag")
		sShortOvertimeFlag = oRS.Collect("ShortOvertimeFlag")
		sMaternityFlag = oRS.Collect("MaternityFlag")
		sDressFreeFlag = oRS.Collect("DressFreeFlag")
		sMammyFlag = oRS.Collect("MammyFlag")
		sFixedTimeFlag = oRS.Collect("FixedTimeFlag")
		sShortTimeFlag = oRS.Collect("ShortTimeFlag")
		sHandicappedFlag = oRS.Collect("HandicappedFlag")
		'**BTM 08/08/19 Lis�� REP
	End if

	'**TOP 08/08/19 Lis�� REP
	'sSQL = "select FlexTime,CompanyKbn from Companyinfo where companycode = '" & dbCompanyCode & "'"
	sSQL = "select CompanyKbn from Companyinfo where companycode = '" & dbCompanyCode & "'"
	'**BTM 08/08/19 Lis�� REP
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		'sFlexTimeFlag = oRS.Collect("FlexTime")		'08/08/19 Lis�� DEL
		sCompanyKbn = oRS.Collect("CompanyKbn")
	End if


	sSQL = "up_DtlOrder '" & rRS.Collect("OrderCode") & "', ''"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbSecretFlag = oRS.Collect("SecretFlag")

		'******************************************************************************
		'���^ start�@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��
		'------------------------------------------------------------------------------

		'�N��
		If GetRSState(oRS) = True Then
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
		End If

		'------------------------------------------------------------------------------
		'���^ end
		'******************************************************************************

		'**************************************************************************
		'�摜 start
		'--------------------------------------------------------------------------
		sSQL = "select ordercode,CASE WHEN ISNULL(CONVERT(VARBINARY, OP.Picture), 0x00) > 0x00 THEN '1' ELSE '0' END AS CompanyPictureFlag from c_info as CI LEFT JOIN OptionPicture AS OP ON CI.CompanyCode = OP.CompanyCode AND OptionNo = 1 Where CI.ordercode='" & dbOrderCode & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			sCompanyPictureFlag = ChkStr(oRS.Collect("CompanyPictureFlag"))
		End If

		flgImg = False
		sImgMain = ""
		sImgSub = ""

		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("OptionNo1")) <> "" Or (sOrderType = "0" And sCompanyPictureFlag = "1") Then
				sImgMain = "<img src=""/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo1") & """ alt="""" border=""0"" width=""100"" height=""75"" style=""float:left; margin-right:5px;"">"
				flgImg = True
			End If
			If sImgSub <> "" Then sImgSub = sImgSub & "<div style=""clear:both;""></div>"
		Else
			If sCompanyPictureFlag = "1" And sOrderType = "0" Then
				sImgMain = "<img src=""/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"" alt="""" border=""0"" width=""100"" height=""75"">"
				flgImg = True
			End If
		End If

		Call RSClose(oRS)
		'--------------------------------------------------------------------------
		'�摜 end
		'**************************************************************************

		'**************************************************************************
		'���img start
		'--------------------------------------------------------------------------
		sImgOrderState = ""
		'�A�N�Z�X����100�𒴂��Ă���΁uHOT�v�\���i���X�����j
		If sAccessCount > 100 Then
			sImgOrderState = sImgOrderState & "<img src=""/img/c_HOT_green.gif"" alt=""�l�C"">&nbsp;"
		End If

		'UPDATE�ƍ�������10�����������Łu�V���v�\��(���X����)
		If sUpdateDay > NOW()-10 Then
			sImgOrderState = sImgOrderState & "<img src=""/img/c_NEW_green.gif"" alt=""�V��"">&nbsp;"
		End If

		'���o���҂n�j�̏ꍇ�A�킩�΃}�[�N�\��(���X����)
		If sInexperiencedPersonFlag = "1" Then
			sImgOrderState = sImgOrderState & "<img src=""/img/no_experience.gif"" alt=""���o���ҁ^���V�����}"">&nbsp;"
		End If

		'�t�^�[���E�h�^�[��
		If sUITurnFlag = "1" Then
			sImgOrderState = sImgOrderState & "<img src=""/img/ui_turn.gif"" alt=""�t�^�[���E�h�^�[��"">&nbsp;"
		End If

		'��w���������d��
		If sUtilizeLanguageFlag = "1" Then
			sImgOrderState = sImgOrderState & "<img src=""/img/linguistic_job.gif"" alt=""��w���������d��"">&nbsp;"
		End If

		'�N�ԋx��120���ȏ�
		If sManyHolidayFlag = "1" Then
			sImgOrderState = sImgOrderState & "<img src=""/img/year_holidaycnt.gif"" alt=""�N�ԋx��120���ȏ�"">&nbsp;"
		End If

		'�t���b�N�X�^�C�����x���� ------2006/01/10 Hayashi ADD
		'**TOP 08/08/19 Lis�� REP
		'If sFlexTimeFlag = "ON" And sOrderType = "0" And sCompanyKbn = "1" Then
		'	sImgOrderState = sImgOrderState & "<img src=""/img/flextime.gif"" alt=""�t���b�N�X�^�C�����x����"">&nbsp;"
		'End If
		If sFlexTimeFlag = "1" Then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_flextime.gif"" alt=""�t���b�N�X�^�C�����x����"">&nbsp;"
		End If
		if sNearStationFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_nearstation.gif"" alt=""�w��(�k��5���ȓ�)"">&nbsp;"
		end if
		if sNoSmokingFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_nosmoking.gif"" alt=""�։��E����"">&nbsp;"
		end if
		if sNewlyBuiltFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_newlybuilt.gif"" alt=""�V�z�r���E�I�t�B�X(5�N�ȓ�)"">&nbsp;"
		end if
		if sLandmarkFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_landmark.gif"" alt=""���w(15�K�ȏ�)�r��"">&nbsp;"
		end if
		if sRenovationFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_renovation.gif"" alt=""���m�x�[�V�����r���E�I�t�B�X(5�N�ȓ�)"">&nbsp;"
		end if
		if sDesignersFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_designers.gif"" alt=""�f�U�C�i�[�Y�r���E�I�t�B�X"">&nbsp;"
		end if
		if sCompanyCafeteriaFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_companycafeteria.gif"" alt=""�Ј��H��"">&nbsp;"
		end if
		if sShortOvertimeFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_shortovertime.gif"" alt=""�c��10h/���ȓ�"">&nbsp;"
		end if
		if sMaternityFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_maternity.gif"" alt=""�Y�x�E��x���т���"">&nbsp;"
		end if
		if sDressFreeFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_dressfree.gif"" alt=""�������R"">&nbsp;"
		end if
		if sMammyFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_mammy.gif"" alt=""�q��ă}�}���}"">&nbsp;"
		end if
		if sFixedTimeFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_fixedtime.gif"" alt=""18���܂łɑގ�"">&nbsp;"
		end if
		if sShortTimeFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_shorttime.gif"" alt=""1��6���Ԉȓ��J��"">&nbsp;"
		end if
		if sHandicappedFlag = "1" then
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/oc_handicapped.gif"" alt=""��Q�Ҋ��}"">&nbsp;"
		end if
		'**BTM 08/08/19 Lis�� REP

		'�V�[�N���b�g���l ------2008/05/08 LIS K.Kokubo ADD
		If dbSecretFlag = "1" Then
			sImgOrderState = sImgOrderState & "<img src=""/img/order/secret.gif"" alt=""�X�J�E�g���󂯂��l�������{���ł��鋁�l���"" width=""50"" height=""15"">&nbsp;"
		End If

		'<��]�Ζ��`�ԃA�C�R���@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��>
		sSQL = "sp_GetDataWorkingType '" & dbOrderCode & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			sWorkingcode = oRS.Collect("WorkingTypecode")
			sWorkingname = GetDetail("WorkingType", sWorkingcode)

			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_w" & sWorkingcode & ".gif"" alt=""" & sWorkingname & """ width=""50"" height=""15"">&nbsp;"

			oRS.MoveNext
		Loop
		Call RSClose(oRS)
		'</��]�Ζ��`�ԃA�C�R���@10��1���ꗗ�ύX�p�ɕ\���ǉ�_�V��>

		'<�Ζ��n�A�C�R��>
		idx = 0
		sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True And idx < 3
			dbWorkingPlacePrefectureCode = ChkStr(oRS.Collect("WorkingPlacePrefectureCode"))
			dbWorkingPlacePrefectureName = ChkStr(oRS.Collect("WorkingPlacePrefectureName"))
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_p" & dbWorkingPlacePrefectureCode & ".gif"" alt=""" & dbWorkingPlacePrefectureName & """ width=""50"" height=""15"">&nbsp;"

			oRS.MoveNext
			idx = idx + 1
		Loop
		Call RSClose(oRS)
		'</�Ζ��n�A�C�R��>
		'--------------------------------------------------------------------------
		'���img end
		'**************************************************************************
	End If

	Response.Write "<table class=""pattern1"" border=""0"" style=""width:600px; table-layout:fixed;"">" & vbCrLf
	Response.Write "<thead>" & vbCrLf
	Response.Write "<tr>" & vbCrLf
	Response.Write "<th style=""font-size:16px; text-align:left; padding:4px 0px 2px 10px;""><span style=""color:#66cc33;"">�� </span>" & rRS.Collect("Subject") & "</th>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	Response.Write "</thead>" & vbCrLf
	Response.Write "<tbody>" & vbCrLf
	Response.Write "<tr>" & vbCrLf

	If vMode = "1" Then
		'���M��ʂ̏ꍇ�͈���
	Else
		'�J���ς݂ɂ���
		sSQL = "sp_Reg_MailOpenDay '" & rRS.Collect("ID") & "'"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
		'��M��ʂ̏ꍇ�͍��o�l
	End If

	Response.Write "<td>"

	If sCompanyPictureFlag <> "" Then
		If sAnswerNG = "0" And Trim(dbOrderCode) <> "" Then
			Response.Write "<a href=""" & HTTPS_CURRENTURL & "order/order_detail.asp?ordercode=" & rRS.Collect("OrderCode") & """ title=""" & sJobTypeDetail & """>" & sImgMain & "</a>"
		ElseIf sAnswerNG = "1" Then
			Response.Write sImgMain
		End If
	End If

	Response.Write rRS.Collect("CompanyName")
	If sAnswerNG = "0" And dbOrderCode <> "" Then
		Response.Write "<br>[<a href=""" & HTTPS_CURRENTURL & "order/order_detail.asp?OrderCode=" & dbOrderCode & """>" & sJobTypeDetail & "�i" & dbOrderCode & "�j</a>]"
	ElseIf sAnswerNG = "1" Then
		Response.Write "<br>" & rRS.Collect("OrderCode")
	End If

	If sImgOrderState <> "" Then
		Response.Write "<div style=""margin-top:5px;width:480px;"">" & sImgOrderState & "</div>"
	End if

	If sYearlyIncome <> "" Then Response.Write "�N���y" & sYearlyIncome & "�z"
	If sMonthlyIncome <> "" Then Response.Write "�����y" & sMonthlyIncome & "�z"
	If sDailyIncome <> "" Then Response.Write "�����y" & sDailyIncome & "�z"
	If sHourlyIncome <> "" Then Response.Write "�����y" & sHourlyIncome & "�z"

	Response.Write "</td>"
	Response.Write "</tr>" & vbCrLf
	Response.Write "<tr>" & vbCrLf
	Response.Write "<td>" & vbCrLf
	Response.Write "<div readonly style=""border:solid 1px #cccccc; overflow:visible; width:576px; background-color:#fffff5; padding:5px;"">" & vbCrLf
	Response.Write Replace(Replace(Replace(rRS.Collect("Body"),vbCrLf,"<br>"), vbCr, "<br>"), vbLf, "<br>") & vbCrLf
	Response.Write "</div>" & vbCrLf
	Response.Write "</td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	Response.Write "</tbody>" & vbCrLf
	Response.Write "</table>" & vbCrLf
	Response.Write "<br>" & vbCrLf
End Function

'******************************************************************************
'�T�@�v�F���[���������ꍇ�̕����o��
'���@���F
'���@�l�F
'�g�p���F/staff/maildetail_person.asp
'�X�@�V�F2007/03/02 LIS K.Kokubo 
'******************************************************************************
Function DspNoMailDetail()
	Response.Write "<b>�w�肳�ꂽ���[���͑��݂��Ȃ����A�폜����Ă��܂��B</b><br><br>"
End Function

'******************************************************************************
'�T�@�v�F���J���ʒm���[�����O�̊J���t���O�𗧂Ă鏈��
'���@���FrDB	�F�ڑ����c�a�R�l�N�V����
'�@�@�@�FrRS	�F���[���ڍׂ̃��R�[�h�Z�b�g(up_GetDetailMail)
'���@�l�F
'�g�p���F/staff/maildetail_person.asp
'�X�@�V�F2007/03/02 LIS K.Kokubo 
'******************************************************************************
Function RegNoticeScoutMailUnRead_OpenFlag(ByRef rDB, ByRef rRS, ByVal vMode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbMailID

	If GetRSState(rRS) = False Then Exit Function
	If vMode = "1" Then Exit Function

	dbMailID = rRS.Collect("ID")

	sSQL = "/*�i�r�F���J���ʒm���[�����O�̊J���t���O�𗧂Ă�*/" & vbCrLf
	sSQL = sSQL & "EXEC up_UpdLOG_Notice_ScoutMailUnRead_OpenFlag '" & dbMailID & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Call RSClose(oRS)
End Function
%>
