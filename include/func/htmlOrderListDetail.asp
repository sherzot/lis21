<%
'******************************************************************************
'�T�@�v�F���l�[�ꗗ�y�[�W�̊e���l�[���ڂ�\��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_SearchOrder or ���l�[�ڍ׌���SQL �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvMyOrder		�F���p�����[�U�̎��Ћ��l�[���ۂ� ["1"]���Ћ��l�[ ["0"]���Ћ��l�[�łȂ�
'�@�@�@�FvHTMLType		�F[xhtml]XHTML�`�� [html]HTML�`��
'�g�p���F/rss/job.asp
'���@�l�F
'���@���F2011/09/21 LIS K.Kokubo �쐬
'******************************************************************************
Function htmlOrderListDetail(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vMyOrder, ByVal vHTMLType)
	Const PICSIZEW = 240
	Const PICSIZEH = 180
	Const PICSIZESUBW = 72
	Const PICSIZESUBH = 56

	Dim sHTML
	Dim sSlash

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

	If GetRSState(rRS) = False Then Exit Function

	If LCase(vHTMLType) = "xhtml" Then sSlash = " /"

	dbOrderCode = rRS.Collect("OrderCode")

	If G_USEFLAG = "0" And vMyOrder = "1" And G_OLDAPPLICATIONCODE <> "" Then
		sSQL = "EXEC up_DtlOrder '" & rRS.Collect("OrderCode") & "', '" & G_OLDAPPLICATIONCODE & "';"
	Else
		sSQL = "EXEC up_DtlOrder '" & rRS.Collect("OrderCode") & "', '';"
	End If

	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

	dbCompanyCode = oRS.Collect("CompanyCode")
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
			sImgMail = "<img src=""/img/s_contact.gif"" alt=""���[�����M�ς�""" & sSlash & ">"
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
	sImgOrderState = htmlOrderSpecialityImg(rDB, oRS, vHTMLType)
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
		sImgMain = "<img src=""" & sImg1 & """ alt="""" border=""0"" width=""" & PICSIZEW & """ height=""" & PICSIZEH & """" & sSlash & ">"
	End If

	If sImg2 <> "" Then
		sImgSub = sImgSub & "<div align=""center"" style=""float:left; width:80px;"">" & _
			"<img src=""" & sImg2 & """ border=""1"" width=""" & PICSIZESUBW & """ height=""" & PICSIZESUBH & """ style=""border:1px solid #666666;""" & sSlash & "><br" & sSlash & ">"
		sImgSub = sImgSub & "</div>"
		flgImg = True
	End If
	If sImg3 <> "" Then
		sImgSub = sImgSub & "<div align=""center"" style=""float:left; width:80px;"">" & _
			"<img src=""" & sImg3 & """ border=""1"" width=""" & PICSIZESUBW & """ height=""" & PICSIZESUBH & """ style=""border:1px solid #666666;""" & sSlash & "><br" & sSlash & ">"
		sImgSub = sImgSub & "</div>"
		flgImg = True
	End If
	If sImg4 <> "" Then
		sImgSub = sImgSub & "<div align=""center"" style=""float:left; width:80px;"">" & _
			"<img src=""" & sImg4 & """ border=""1"" width=""" & PICSIZESUBW & """ height=""" & PICSIZESUBH & """ style=""border:1px solid #666666;""" & sSlash & "><br" & sSlash & ">"
		sImgSub = sImgSub & "</div>"
	End If

	If sImgSub <> "" Then sImgSub = "<div style=""padding-top:1px;"">" & sImgSub & "<div style=""clear:both;""></div></div>"
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
		If (oRS.Collect("OrderType") ="0" And oRS.Collect("Companykbn") = "2") Or oRS.Collect("OrderType") ="1" Or oRS.Collect("OrderType") ="2" Or oRS.Collect("OrderType") ="3" Then
			Select Case oRS2.Collect("WorkingTypeCode")
				Case "001": sWorkingType = sWorkingType & "�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/koyoukeitai_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>�h���Ƃ�</a>�z" 
				Case "002","003": sWorkingType = sWorkingType & "�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>�l�ޏЉ�Ƃ�</a>�z" 
				Case "004": sWorkingType = sWorkingType & "�y<a href=""javascript:void(0)"" onclick='window.open(""/staff/syoukaiyotei_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>�Љ�\��h���Ƃ�</a>�z" 
			End Select
		End If
		sWorkingType = sWorkingType & "<br" & sSlash & ">"
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
			sImgOrderState = sImgOrderState & "<img src=""/img/order_detail_icon/icon_p" & dbWorkingPlacePrefectureCode & ".gif"" alt=""" & dbWorkingPlacePrefectureName & """ width=""50"" height=""15""" & sSlash & ">&nbsp;"
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
			If sBizName1 <> "" And sBizPercentage1 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName1 & "</td><td class=""biz2"">" & sBizPercentage1 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage1) * 3 & """ height=""20""" & sSlash & "></td></tr>"
			If sBizName2 <> "" And sBizPercentage2 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName2 & "</td><td class=""biz2"">" & sBizPercentage2 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage2) * 3 & """ height=""20""" & sSlash & "></td></tr>"
			If sBizName3 <> "" And sBizPercentage3 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName3 & "</td><td class=""biz2"">" & sBizPercentage3 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage3) * 3 & """ height=""20""" & sSlash & "></td></tr>"
			If sBizName4 <> "" And sBizPercentage4 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & sBizName4 & "</td><td class=""biz2"">" & sBizPercentage4 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(sBizPercentage4) * 3 & """ height=""20""" & sSlash & "></td></tr>"
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

	sHTML = sHTML & "<input type=""hidden"" name=""CONF_OrderCodes"" value=""" & oRS.Collect("OrderCode") & """>"
	sHTML = sHTML & "<table border=""0"" class=""old"">"
	sHTML = sHTML & "<tbody>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td class=""old11"" style=""padding-left:0px; width:600px;"" valign=""middle"">"

	If vUserType = "" Or vUserType = "staff" Then
		'�񃍃O�C�����A�X�^�b�t���O�C����

		'�E���l�[�t�q�k�����[�����M
		'�E�E�H�b�`���X�g�֕ۑ�
		sHTML = sHTML & "<div style=""float:left;width:420px;"">"
		sHTML = sHTML & "<img src=""/img/list_companyicon.gif"" alt="""" align=""left""" & sSlash & ">" & sTitleCompanyName
		sHTML = sHTML & "<h3 style=""margin-left:5px;"">��<a href=""" & HTTP_CURRENTURL & "order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """>" & sTitleJobName & "</a>" & sImgMail & "</h3>"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div align=""right"" style=""float:right;font-size:11px;width:113px;"">"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "order/sendmail_jobofferaddress.asp?OrderCode=" & oRS.Collect("OrderCode") & """ onclick=""window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;""><img src=""/img/order/ordermail.gif"" style=""margin-bottom:6px;"" border=""0"" alt=""���l�������[�����M"" align=""top""" & sSlash & "></a>"
		sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "order/sendmail_jobofferaddress.asp?OrderCode=" & oRS.Collect("OrderCode") & """ onclick=""window.open(this.href,'sendmail_jobofferaddress','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=640,height=490');return false;""><img src=""/img/order/orderwachlist.gif"" border=""0"" alt=""�E�H�b�`���X�g�ɒǉ�"" align=""top""" & sSlash & "></a>"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
	ElseIf vUserType = "company" Then
		'��ƃ��O�C����
		sHTML = sHTML & "<p class=""m0""><img src=""/img/list_companyicon.gif"" alt="""" align=""left""" & sSlash & ">" & sTitleCompanyName & "</p>"
		sHTML = sHTML & "<h3 style=""margin-left:5px;"">��<a href=""../order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """>" & sTitleJobName & "</a>" & sImgMail & "</h3>"
	End If

	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td class=""old12"">"
	'**TOP 08/08/19 Lis�� REP
	'sHTML = sHTML & "<div style=""float:left;"">" & sImgOrderState & "</div>"
	'sHTML = sHTML & "<div align=""right"" style=""font-size:10px;line-height:14px;"">�f�ڊ����F" & sPublishLimitStr & "</div>"
	'sHTML = sHTML & "<div style=""clear:both;""></div>"
	sHTML = sHTML & "<table style='width:600px;'><tr><td style='width:500px;padding-left:5px;'>" & sImgOrderState & "</td>"
	sHTML = sHTML & "<td style='width:100px;vertical-align:top;font-size:10px;text-align:right;'>�f�ڊ����F"
	sHTML = sHTML & sPublishLimitStr & "</td></tr></table>"
	'**BTM 08/08/19 Lis�� REP
	sHTML = sHTML & "<table border=""0"" class=""old2"">"

	If sCatchCopy <> "" Then
		sHTML = sHTML & "<caption>" & sCatchCopy & "</caption>"
	End If

	sHTML = sHTML & "<tbody>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td rowspan=""2"" valign=""top"">"

	If flgImg = True Then
		'�摜���L��ꍇ�̃��C�A�E�g
		sHTML = sHTML & "<div class=""old21"" style=""margin:0px 12px;"">"
		sHTML = sHTML & "<b>�y�S���Ɩ��̐����z</b><br" & sSlash & ">" & sBusinessDetail
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div class=""old21"" style=""width:240px; float:left; margin:0px 5px;"">"
		sHTML = sHTML & "<a href=""" & HTTP_NAVI_CURRENTURL & "order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """ title=""" & sTitleCompanyName & """>" & sImgMain & "</a>"
		sHTML = sHTML & sImgSub
		sHTML = sHTML & "</div>"
	Else
		'�摜�������ꍇ�̃��C�A�E�g
		sHTML = sHTML & "<div class=""old21"" style=""width:239px; float:left; margin:0px 5px;"">"
		sHTML = sHTML & "<b>�y�S���Ɩ��̐����z</b><br" & sSlash & ">" & sBusinessDetail
		sHTML = sHTML & "</div><br" & sSlash & ">"
	End If

	sHTML = sHTML & "<table style=""width:330px; margin-left:3px;"">"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td style=""font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;"">"
	sHTML = sHTML & "�Ζ��`��"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "<td style=""background-color:#eeeeee; padding:5px 0px 5px 10px; border-bottom:solid 3px #ffffff;"">"
	sHTML = sHTML & sWorkingType
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td style=""font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;"">"
	sHTML = sHTML & "�Ζ��n"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "<td style=""background-color:#eeeeee; padding-left:10px; border-bottom:solid 3px #ffffff;"">"
	sHTML = sHTML & sWorkingPlace & "" & sStationName
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	If sYearlyIncome & sMonthlyIncome & sDailyIncome & sHourlyIncome <> "" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td style=""font-weight:bold; background-color:#E1FBCD; width:70px; text-align:center; line-height:30px; border-bottom:solid 3px #ffffff;"">"
		sHTML = sHTML & "���^"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td style=""background-color:#eeeeee; padding:5px 0px 5px 10px; border-bottom:solid 3px #ffffff;"">"

		If sYearlyIncome <> "" Then
			sHTML = sHTML & "<p>�N��&nbsp;" & sYearlyIncome & "</p>"
		End If

		If sMonthlyIncome <> "" Then
			sHTML = sHTML & "<p>����&nbsp;" & sMonthlyIncome & "</p>"
		End If

		If sDailyIncome <> "" Then
			sHTML = sHTML & "<p>����&nbsp;" & sDailyIncome & "</p>"
		End If

		If sHourlyIncome <> "" Then
			sHTML = sHTML & "<p>����&nbsp;" & sHourlyIncome & "</p>"
		End If

		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	If sBizName1 <> "" Then

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td style=""font-weight:bold; background-color:#E1FBCD; width:70px; border-bottom:solid 3px #ffffff; text-align:center;"">"
		sHTML = sHTML & "�d���̊���"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td style=""background-color:#eeeeee; border-bottom:solid 3px #ffffff; padding-left:0px; line-height:14px;"">"
		sHTML = sHTML & "<table>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td style=""padding:5px 0px 5px 7px;"">"
		sHTML = sHTML & "<script type=""text/javascript"" language=""javascript"">"
		sHTML = sHTML & "viewWorkAvg(" & sBizPercentage1 & ", " & sBizPercentage2 & ", " & sBizPercentage3 & ", " & sBizPercentage4 & ")"
		sHTML = sHTML & "</script>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td>"

		If sBizName1 <> "" Then sHTML = sHTML & "<p style=""font-size:10px; line-height:12px;""><span style=""color:#ff9999;"">��</span>" & sBizPercentage1 & "%�@" & sBizName1 & "</p>"
		If sBizName2 <> "" Then sHTML = sHTML & "<p style=""font-size:10px; line-height:12px;""><span style=""color:#9999ff;"">��</span>" & sBizPercentage2 & "%�@" & sBizName2 & "</p>"
		If sBizName3 <> "" Then sHTML = sHTML & "<p style=""font-size:10px; line-height:12px;""><span style=""color:#99ff99;"">��</span>" & sBizPercentage3 & "%�@" & sBizName3 & "</p>"
		If sBizName4 <> "" Then sHTML = sHTML & "<p style=""font-size:10px; line-height:12px;""><span style=""color:#ffff99;"">��</span>" & sBizPercentage4 & "%�@" & sBizName4 & "</p>"

		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</table>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	sHTML = sHTML & "</table>"
	sHTML = sHTML & "<div align=""right"" style=""margin:3px 5px;"">"

	If dbWValueURL <> "" Then
		sHTML = sHTML & "<a href=""" & dbWValueURL & """ target=""_blank""><img src=""/img/order/btn_wvalue.gif"" border=""0"" alt=""���l���:" & sTitleCompanyName & "�̎��Ѝ̗p�y�[�W""" & sSlash & "></a>"
	End If

	If dbTopInterviewFlag = "1" Then
		sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "order/order_interview.asp?ordercode=" & dbOrderCode & """><img src=""/img/order/interview_icon.gif"" border=""0"" alt=""���l���:�g�b�v�C���^�r���[""" & sSlash & "></a>"
	End If

	sHTML = sHTML & "<a href=""" & HTTP_CURRENTURL & "order/order_detail.asp?OrderCode=" & oRS.Collect("OrderCode") & """><img src=""/img/detail_button2.gif"" border=""0"" alt=""���l���:�ڍ�""" & sSlash & "></a>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</table>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	If oRS.Collect("CompanyCode") = vUserID And vMyOrder = "1" And G_USEFLAG = "1" Then
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td class=""old13"">"
		sHTML = sHTML & "<table class=""old3"">"
		sHTML = sHTML & "<tbody>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td class=""old31"">���R�[�h(" & oRS.Collect("OrderCode") & ")</td>"
		sHTML = sHTML & "<td class=""old32"">���</td>"
		sHTML = sHTML & "<td class=""old33"">"
		sHTML = sHTML & sProgress & "&nbsp;"
		sHTML = sHTML & "<select name=""CONF_PublicFlags"" " & sPublicListDsp & ">"
		If oRS.Collect("PublicFlag") = "1" Then
			sHTML = sHTML & "<option value=""1"" selected>�f��</option>"
			sHTML = sHTML & "<option value=""0"">��f��</option>"
		Else
			sHTML = sHTML & "<option value=""1"">�f��</option>"
			sHTML = sHTML & "<option value=""0"" selected>��f��</option>"
		End If
		sHTML = sHTML & "</select>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td class=""old34"">�f�ړ�<br" & sSlash & ">�o�^��</td>"
		sHTML = sHTML & "<td class=""old35"">" & sPublicDay & "<br" & sSlash & ">" & sRegistDay & "</td>"
		'sHTML = sHTML & "<td class=""old36""><input type=""checkbox"" name=""CONF_DeleteFlags"" value=""" & oRS.Collect("OrderCode") & """>�폜</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
	End If

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td class=""old14""></td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</table>"

	htmlOrderListDetail = sHTML
End Function
%>
