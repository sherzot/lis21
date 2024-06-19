<%
'******************************************************************************
'�T�@�v�F���l�[�̓���
'���@���FrDB		�F
'�@�@�@�FrRS		�F
'�@�@�@�FvHTMLType	�F
'�߂�l�F
'���@�l�F
'���@���F2011/09/21 LIS K.Kokubo �쐬
'******************************************************************************
Function htmlOrderSpecialityImg(ByRef rDB, ByRef rRS, ByVal vHTMLType)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode
	Dim dbWorkingPlacePrefectureCode
	Dim dbWorkingPlacePrefectureName

	Dim sHTML
	Dim sSlash
	Dim sWorkingCode

	If GetRSState(rRS) = False Then Exit Function

	If LCase(vHTMLType) = "xhtml" Then sSlash = " /"

	dbOrderCode = rRS.Collect("OrderCode")

	sHTML = ""
	'�A�N�Z�X����100�𒴂��Ă���΁uHOT�v�\���i���X�����j
	If rRS.Collect("AccessCount") > 100 Then sHTML = sHTML & "<img src=""/img/c_HOT_green.gif"" alt=""�l�C"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'UPDATE�ƍ�������10�����������Łu�V���v�\��(���X����)
	If rRS.Collect("Updateday") > NOW()-10 Then sHTML = sHTML & "<img src=""/img/c_NEW_green.gif"" alt=""�V��"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'���o���҂n�j�̏ꍇ�A�킩�΃}�[�N�\��(���X����)
	If rRS.Collect("InexperiencedPersonFlag") = "1" Then sHTML = sHTML & "<img src=""/img/no_experience.gif"" alt=""���o���ҁ^���V�����}"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'�t�^�[���E�h�^�[��
	If rRS.Collect("UITurnFlag") = "1" Then sHTML = sHTML & "<img src=""/img/ui_turn.gif"" alt=""�t�^�[���E�h�^�[��"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'��w���������d��
	If rRS.Collect("UtilizeLanguageFlag") = "1" Then sHTML = sHTML & "<img src=""/img/linguistic_job.gif"" alt=""��w���������d��"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'�N�ԋx��120���ȏ�
	If rRS.Collect("ManyHolidayFlag") = "1" Then sHTML = sHTML & "<img src=""/img/year_holidaycnt.gif"" alt=""�N�ԋx��120���ȏ�"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2006/01/10 M.Hayashi ADD �t���b�N�X�^�C�����x����
	If rRS.Collect("FlexTimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_flextime.gif"" alt=""�t���b�N�X�^�C�����x����"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NearStationFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_nearstation.gif"" alt=""�w��(�k��5���ȓ�)"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NoSmokingFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_nosmoking.gif"" alt=""�։��E����"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NewlyBuiltFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_newlybuilt.gif"" alt=""�V�z�r���E�I�t�B�X(5�N�ȓ�)"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("LandmarkFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_landmark.gif"" alt=""���w(15�K�ȏ�)�r��"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("RenovationFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_renovation.gif"" alt=""���m�x�[�V�����r���E�I�t�B�X(5�N�ȓ�)"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("DesignersFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_designers.gif"" alt=""�f�U�C�i�[�Y�r���E�I�t�B�X"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("CompanyCafeteriaFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_companycafeteria.gif"" alt=""�Ј��H��"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("ShortOvertimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_shortovertime.gif"" alt=""�c��10h/���ȓ�"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("MaternityFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_maternity.gif"" alt=""�Y�x�E��x���т���"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("DressFreeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_dressfree.gif"" alt=""�������R"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("MammyFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_mammy.gif"" alt=""�q��ă}�}���}"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("FixedTimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_fixedtime.gif"" alt=""18���܂łɑގ�"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("ShortTimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_shorttime.gif"" alt=""1��6���Ԉȓ��J��"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("HandicappedFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_handicapped.gif"" alt=""��Q�Ҋ��}"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("RentAllFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_rentallflag.gif"" alt=""�Z���p�S�z�⏕����"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("RentPartFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_rentpartflag.gif"" alt=""�Z���p�ꕔ�⏕����"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MealsFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_mealsflag.gif"" alt=""�H���E�d���t���Č�"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MealsAssistanceFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_mealsassistanceflag.gif"" alt=""�H���⏕���x����"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("TrainingCostFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_trainingcostflag.gif"" alt=""���C������x����"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("EntrepreneurCostFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_entrepreneurcostflag.gif"" alt=""�N�Ƌ@�ޕ⏕���x����"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MoneyFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_moneyflag.gif"" alt=""�����q�E�ᗘ�q�⏕���x����"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("LandShopFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_landshopflag.gif"" alt=""�y�n�E�X�ܓ��񋟐��x����"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("FindJobFestiveFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_findjobfestiveflag.gif"" alt=""�A�E���j�������x����"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/12/01 LIS K.Kokubo ADD 
	If rRS.Collect("AppointmentFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_appointmentflag.gif"" alt=""���Ј��o�p���x����"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2009/12/01 LIS K.Kokubo ADD 
	If rRS.Collect("SocietyInsuranceFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_societyinsuranceflag.gif"" alt=""�Еۊ���"" width=""50"" height=""15""" & sSlash & ">&nbsp;"
	'2008/05/08 LIS K.Kokubo ADD �V�[�N���b�g���l
	If rRS.Collect("SecretFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order/secret.gif"" alt=""�X�J�E�g���󂯂��l�������{���ł��鋁�l���"" width=""50"" height=""15""" & sSlash & ">&nbsp;"

	'����Yahoo!�̌������炨�d�����ڍ׃y�[�W�֗���l�փA�C�R���\��
	If InStr(Request.ServerVariables("HTTP_REFERER"),"search.yahoo.co.jp/") <> 0 Then
		sSQL = "sp_GetDataWorkingType '" & dbOrderCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		Do While GetRSState(oRS) = True
			sWorkingcode = oRS.Collect("WorkingTypecode")

			sHTML = sHTML & "<img src=""/img/order_detail_icon/icon_w" & sWorkingcode & ".gif"" alt=""�h���Ј�"" width=""50"" height=""15""" & sSlash & ">&nbsp;"

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
				sHTML = sHTML & "<img src=""/img/order_detail_icon/icon_p" & dbWorkingPlacePrefectureCode & ".gif"" alt=""" & dbWorkingPlacePrefectureName & """ width=""50"" height=""15""" & sSlash & ">&nbsp;"
			End If
		End If
		Call RSClose(oRS)
		'</�Ζ��n>
	End If

	htmlOrderSpecialityImg = sHTML
End Function
%>
