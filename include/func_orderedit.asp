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
'��@���F�������@���l�[�ڍ׃y�[�W�o�͗p�@������
'�@�@�@�FChkAgencyEditOrder				�F�㗝�X�̋��l�[�o�^�����ۃ`�F�b�N
'�@�@�@�FChkEditOrder					�F���l�[�o�^�����ۃ`�F�b�N
'�@�@�@�FDspLisOrderCompanyInfo			�F���l�[�ҏW�y�[�W�̃��X�̏Љ��E�h�����Ə��HTML���擾
'�@�@�@�FGetHTMLEditTempOrderCompanyInfo�F���l�[�ҏW�y�[�W�̔h����Ƃ̔h�����Ə��HTML���擾
'�@�@�@�FGetHTMLEditOrderCompanyName	�F���l�[�ҏW�y�[�W�̊�Ɩ�HTML���擾
'�@�@�@�FGetHTMLEditOrderShowTypeSwitch	�F���l�[�ҏW�y�[�W�̉�Џ��E�E����E�C���^�r���[�؂�ւ��{�^���ƎQ�Ɖ�HTML���擾
'�@�@�@�FGetHTMLEditOrderCatchCopy		�F���l�[�ҏW�y�[�W�̃L���b�`�R�s�[�����i�傫���摜�ȂǁjHTML���擾
'�@�@�@�FGetHTMLEditOrderFreePR			�F���l�[�ҏW�y�[�W�̃t���[�o�qHTML���擾
'�@�@�@�FGetHTMLEditOrderPictureNow		�F���l�[�ҏW�y�[�W�̏������摜HTML���擾
'�@�@�@�FGetHTMLEditOrderBackGround		�F���l�[�ҏW�y�[�W�̗̍p�̔w�iHTML���擾
'�@�@�@�FGetHTMLEditBusiness			�F���l�[�ҏW�y�[�W�̋Ɩ����eHTML���擾
'�@�@�@�FGetHTMLEditCondition			�F���l�[�ҏW�y�[�W�̋Ζ�����HTML���擾
'�@�@�@�FGetHTMLEditNeedCondition		�F���l�[�ҏW�y�[�W�̕K�v����HTML���擾
'�@�@�@�FGetHTMLEditHowToEntry			�F���l�[�ҏW�y�[�W�̉�����HTML���擾
'�@�@�@�FGetHTMLEditContact				�F���l�[�ҏW�y�[�W�̒S���ҘA����HTML���擾
'�@�@�@�FGetHTMLElderInterview			�F���l�[�ҏW�y�[�W�̐�y�C���^�r���[HTML���擾
'�@�@�@�FGetWorkingType					�F���l�[�ڍ׃y�[�W�̋Ζ��`�ԕ���
'�@�@�@�FGetJobType						�F���l�[�ڍ׃y�[�W�̐E�핔��
'�@�@�@�FGetWorkingTime					�F���l�[�ڍ׃y�[�W�̋Ζ��`�ԕ���
'�@�@�@�FGetNearbyStation				�F���l�[�ڍ׃y�[�W�̍Ŋ�w����
'�@�@�@�FGetNearbyRailway				�F���l�[�ڍ׃y�[�W�̍Ŋ񉈐�����
'�@�@�@�FGetSkill						�F���l�[�ڍ׃y�[�W�̃X�L������
'�@�@�@�FGetLicense						�F���l�[�ڍ׃y�[�W�̎��i����
'�@�@�@�FGetOrderNote					�F���l�[�ڍ׃y�[�W�̎��i����
'�@�@�@�FGetOrderTitle					�F���l�[�ڍ׃y�[�W�̃^�C�g���ƃf�B�X�N���v�V�������擾
'�@�@�@�FGetSkillList					�F���l�[�ڍ׃y�[�W�̃X�L���̊e���ڕ\��
'�@�@�@�FGetHTMLOrderInputWorkingType	�F���l�[���͉�ʂ̌ٗp�`�ԕ���
'**********************************************************************************************************************

'******************************************************************************
'�T�@�v�F�㗝�X�̋��l�[�o�^�����ۃ`�F�b�N
'���@���FvAgencyCode		�F���O�C�����㗝�X�R�[�h
'�@�@�@�FvBranchSeq			�F���O�C�����㗝�X���_�ԍ�
'�@�@�@�FvApplicationCode	�F�\�����݃R�[�h
'�@�@�@�FvOrderCode			�F���R�[�h
'�߂�l�FBoolean	�F[True]���l�[�o�^�\ [False]���l�[�o�^�s��
'���@�l�F
'�g�p���F
'�X�@�V�F2010/03/30 LIS K.kokubo �쐬
'******************************************************************************
Function ChkAgencyEditOrder(ByVal vAgencyCode, ByVal vBranchSeq, ByVal vApplicationCode, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim sError
	Dim flgQE

	Dim dbUserCode
	Dim dbHakouDate
	Dim dbRiyoFromDate
	Dim dbRiyoToDate
	Dim dbPlanTypeName
	Dim dbCompanyKbn
	Dim dbInterviewFlag
	Dim dbTempPermitFlag
	Dim dbIntroPermitFlag
	Dim dbCheck
	Dim dbLicenseFlag

	ChkAgencyEditOrder = False

'	If vOrderCode = "" Then Exit Function

	'<���C�Z���X�؂�̓}�C���j���[�փ��_�C���N�g>
	sSQL = "EXEC up_ChkAGC_MyNaviLicense '" & vApplicationCode & "','" & vAgencyCode & "','" & vBranchSeq & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbUserCode = oRS.Collect("UserCode")
		dbHakouDate = oRS.Collect("HakouDate")
		dbRiyoFromDate = oRS.Collect("RiyoFromDate")
		dbRiyoToDate = oRS.Collect("RiyoToDate")
		dbPlanTypeName = oRS.Collect("PlanTypeName")
		dbCompanyKbn = oRS.Collect("CompanyKbn")
		dbInterviewFlag = oRS.Collect("InterviewFlag")
		dbTempPermitFlag = oRS.Collect("TempPermitFlag")
		dbIntroPermitFlag = oRS.Collect("IntroPermitFlag")

		If Not(dbHakouDate <= Date And dbRiyoToDate >= Date) Then Exit Function
	Else
		Exit Function
	End If
	Call RSClose(oRS)
	'</���C�Z���X�؂�̓}�C���j���[�փ��_�C���N�g>

	'<���O�C�����̊�Ƃ̏��R�[�h���ǂ������`�F�b�N>
	sSQL = "EXEC sp_ChkCompanyOrder '" & dbUserCode & "', '" & vOrderCode & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbCheck = oRS.Collect("CheckFlag")
		dbLicenseFlag = oRS.Collect("LicenseFlag")
	End If
	Call RSClose(oRS)
	If vOrderCode = "" Then dbCheck = "1"
	If dbCheck = "0" And dbLicenseFlag = "0" Then Exit Function
	'</���O�C�����̊�Ƃ̏��R�[�h���ǂ������`�F�b�N>

	G_USERID = dbUserCode
	G_USERTYPE = "company"
	G_APPLICATIONCODE = qsApplicationCode
	G_USEFLAG = "1"
	G_COMPANYKBN = dbCompanyKbn
	G_PLANTYPE = dbPlanTypeName
	G_INTERVIEWFLAG = dbInterviewFlag
	G_TEMPPERMITFLAG = dbTempPermitFlag
	G_INTROPERMITFLAG = dbIntroPermitFlag
	If dbRiyoFromDate <= Date And dbRiyoToDate >= Date Then
		G_PUBLICFLAG = "1"
	Else
		G_PUBLICFLAG = "0"
	End If

	ChkAgencyEditOrder = True
End Function

'******************************************************************************
'�T�@�v�F���l�[�o�^�����ۃ`�F�b�N
'���@���FvOrderCode	�F���R�[�h
'�@�@�@�FvUserID	�F���O�C�������[�U�R�[�h
'�@�@�@�FvUseFlag	�F���O�C������Ƃ̃��C�Z���X�̗L���t���O
'�߂�l�FBoolean	�F[True]���l�[�o�^�\ [False]���l�[�o�^�s��
'���@�l�F
'�g�p���F�����ƃi�r/company/order/edit01.asp
'�X�@�V�F2008/10/08 LIS K.kokubo �쐬
'******************************************************************************
Function ChkEditOrder(ByVal vOrderCode, ByVal vUserID, ByVal vUseFlag)
	Dim sSQL
	Dim oRS
	Dim sError
	Dim flgQE

	Dim dbCheck
	Dim dbLicenseFlag

'	If vOrderCode = "" Then Exit Function

	'<���C�Z���X�؂�̓}�C���j���[�փ��_�C���N�g>
	sSQL = "EXEC up_DtlNaviLicense_Now '" & vUserID & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("LicenseType1Flag") <> "1" Then Exit Function
	End If
	Call RSClose(oRS)
	'</���C�Z���X�؂�̓}�C���j���[�փ��_�C���N�g>

	'<���O�C�����̊�Ƃ̏��R�[�h���ǂ������`�F�b�N>
	sSQL = "sp_ChkCompanyOrder '" & vUserID & "', '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbCheck = oRS.Collect("CheckFlag")
		dbLicenseFlag = oRS.Collect("LicenseFlag")
	End If
	Call RSClose(oRS)
	If vOrderCode = "" Then dbCheck = "1"
	If dbCheck = "0" And dbLicenseFlag = "0" Then Exit Function
	'</���O�C�����̊�Ƃ̏��R�[�h���ǂ������`�F�b�N>

	ChkEditOrder = True
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
<%		End If
		'**TOP 08/08/21 Lis�� ADD
		If sCapitalAmount <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>���{��</h4></div>
<div class="value1"><p class="m0"><%= sCapitalAmount %></p></div>
<div style="clear:both;"></div>
<%		End If
		If sEmployeeNum <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
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
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
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
'�T�@�v�F���l�[�̉�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit22.asp
'�X�@�V�F2009/03/17 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHTMLEditTempOrderCompanyInfo(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
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
	Dim sCapital
	Dim sTempAllEmployeeNumber

	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function

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
	Else
		dbTempEstablishYear = "<span style=""color:#999999;"">[�ݗ��N�x]�������͂ł��B</span>"
	End If
	'</�ݗ��N�x>

	'<�Ǝ�>
	If dbTempIndustryTypeName <> "" Then
	Else
		dbTempIndustryTypeName = "<span style=""color:#999999;"">[�Ǝ�]�������͂ł��B</span>"
	End If
	'</�Ǝ�>

	'<���{>
	sCapital = ""
	If dbTempCapitalAmount & dbTempForeinCapital <> "" Then
		If dbTempCapitalAmount <> "" Then
			sCapital = sCapital & GetJapaneseYen(dbTempCapitalAmount)
		Else
			sCapital = sCapital & "<span style=""color:#999999;"">[���{��]�������͂ł��B</span>"
		End If

		If dbTempForeinCapital <> "" Then
			sCapital = sCapital & "&nbsp;�i�O���F" & dbTempForeinCapital & "�j"
		Else
			sCapital = sCapital & "<br><span style=""color:#999999;"">[�O��]�������͂ł��B</span><br>"
		End If
	End If
	'</���{>

	'<����>
	If dbTempListClass <> "" Then
	Else
		dbTempListClass = "<span style=""color:#999999;"">[����]�������͂ł��B</span>"
	End If
	'</����>

	'<�Ј���>
	If dbTempAllEmployeeNumber <> "" Then
		sTempAllEmployeeNumber = dbTempAllEmployeeNumber & "�l"
	Else
		dbTempAllEmployeeNumber = "<span style=""color:#999999;"">[�Ј���]�������͂ł��B</span>"
	End If
	'</�Ј���>

	flgLine = False

	sHTML = sHTML & "<a name=""edit22""></a>"
	sHTML = sHTML & "<h3>�h�����Ə��</h3>" & vbCrLf
	sHTML = sHTML & "<div style=""margin-left:15px;""><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit22.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""></div>"

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>�ݗ��N�x</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbTempEstablishYear & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>�Ǝ�</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbTempIndustryTypeName & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>���{��</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"
	sHTML = sHTML & "<p class=""m0"">" & sCapital & "</p>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>����</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbTempListClass & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>�Ј���</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sTempAllEmployeeNumber & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "<br>" & vbCrLf

	GetHTMLEditTempOrderCompanyInfo = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[�ҏW�y�[�W�̊�Ɩ��̂��o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit0.asp
'�X�@�V�F2008/10/10 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHTMLEditOrderCompanyName(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sHTML
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbSecretFlag		'�V�[�N���b�g�t���O
	Dim dbCompanyName		'��Ɩ���
	Dim dbCompanyNameF		'��Ɩ��̃J�i
	Dim dbCompanyKbn		'��Ƌ敪
	Dim dbCompanySpeciality	'��Ɠ���

	Dim sPublishLimitStr	'�f�ڊ����\���p������
	Dim sCautionStr			'�f�ڊ����\�����ӕ���������
	Dim flgNowPublic		'���݌f�ڒ��̋��l�[���� '[True]�f�ڒ� [False]��f��

	If GetRSState(rRS) = False Then Exit Function

	sHTML = ""
	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbSecretFlag = rRS.Collect("SecretFlag")
	dbCompanyName = rRS.Collect("CompanyName")
	dbCompanyNameF = rRS.Collect("CompanyName_F")
	dbCompanyKbn = rRS.Collect("CompanyKbn")
	dbCompanySpeciality = ChkStr(rRS.Collect("CompanySpeciality"))
	Call SetOrderCompanyName(dbCompanyName, dbCompanyNameF, dbOrderType, dbCompanyKbn, dbCompanySpeciality)

	'���X�Љ�Č��̏ꍇ�́u�]�E���k�Č��v�C���[�W��\��
	If dbOrderType = "2" Then sHTML = sHTML & "<img src=""/img/order/counselable_order.gif"" width=""150"" height=""25"" alt=""�]�E�x�����󂯂ĉ��傷�鋁�l�ł�"">"
	'�V�[�N���b�g���l�̏ꍇ�́u�V�[�N���b�g���l�v�C���[�W��\��
	'If dbSecretFlag = "1" Then sHTML = sHTML & "<img src=""/img/order/secret_order.gif"" width=""150"" height=""25"" alt=""���̋��l����X�J�E�g���󂯂��l�������{���ł��鋁�l�ł�"">"
	If dbSecretFlag = "1" Then sHTML = sHTML & "<p class=""m0"" style=""color:#ff9933; font-weight:bold;"">���X�J�E�g���󂯂��l�������{���ł��鋁�l���ł��B</p>"

	sHTML = sHTML & "<div style=""width:400px; margin-bottom:10px;"">"
	If G_COMPANYKBN = "2" Then
		sHTML = sHTML & "<a name=""edit21""></a>"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit21.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
		If dbCompanySpeciality = "" Then
			sHTML = sHTML & "&nbsp;<span style=""color:#999999;"">��[�\���p��Ж�]�������͂ł��B</span>"
		End If
	End If
	sHTML = sHTML & "<div style=""font-size:14px; font-weight:bold;"">" & dbCompanyName & "</div>"
	sHTML = sHTML & "<div style=""font-size:10px; color:#666666;"">" & dbCompanyNameF & "</div>"
	sHTML = sHTML & "</div>"

	GetHTMLEditOrderCompanyName = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̉�Џ��E�E����E�C���^�r���[�؂�ւ��{�^���ƎQ�Ɖ񐔂��o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvType			�F�\�������̎�� ["0"]�E���� ["1"]��Џ�� ["2"]�C���^�r���[
'�@�@�@�FvAccessCount	�F�\�������l�[�̃A�N�Z�X��
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit0.asp
'�X�@�V�F2008/10/10 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHTMLEditOrderShowTypeSwitch(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vType, ByVal vAccessCount)
	'<�ϐ��錾>
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode			'���R�[�h
	Dim dbOrderType			'�󒍎��
	Dim dbJobTypeDetail		'��̓I�E�햼
	Dim dbTopInterviewFlag	'�g�b�v�C���^�r���[�L���t���O
	Dim sUpdateDay

	Dim sHTML
	'</�ϐ��錾>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbJobTypeDetail = rRS.Collect("JobTypeDetail")
	dbTopInterviewFlag = rRS.Collect("TopInterviewFlag")
	'�X�V��
	sUpdateDay = GetDateStr(rRS.Collect("UpdateDay"), "/")

	If dbJobTypeDetail <> "" Then dbJobTypeDetail = dbJobTypeDetail & "�̂��d�����ڍ�"

	sHTML = sHTML & "<div style=""width:600px; margin-bottom:5px;"">"
	sHTML = sHTML & "<div style=""float:left; width:350px; margin:0px;"">"
	If vType = "0" Then
		'�d������\�����̏ꍇ
		sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_orderdetail_on.gif"" alt=""" & dbJobTypeDetail & """ border=""0"" width=""93"" height=""22""></div>"
		If dbOrderType = "0" Then
			'��ʂ̋��l�L���̏ꍇ�͉�Џ��ւ̃����N��\��
			sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><a href=""./company_order.asp?poc=" & dbOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_companyinfo_off.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></a></div>"
		End If

		If dbOrderType = "0" And dbTopInterviewFlag = "1" Then
			sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & dbOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_interview_off.gif"" alt=""�C���^�r���["" border=""0"" width=""93"" height=""22""></a></div>"
		End If
	ElseIf vType = "1" Then
		'��Џ���\�����̏ꍇ
		sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & dbOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & dbJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
		If dbOrderType = "0" Then
			'��ʂ̋��l�L���̏ꍇ�͉�Џ���\��
			sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_companyinfo_on.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></div>"
		End If

		If dbOrderType = "0" And dbTopInterviewFlag = "1" Then
			sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & dbOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_interview_off.gif"" alt=""�C���^�r���["" border=""0"" width=""93"" height=""22""></a></div>"
		End If

	ElseIf vType = "2" Then
		'�C���^�r���[��\�����̏ꍇ
		sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & dbOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & dbJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
		sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><a href=""./company_order.asp?poc=" & dbOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_companyinfo_off.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></a></div>"
		sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_interview_on.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></div>"
	End If
	sHTML = sHTML & "<div class=""clear:both; margin:0px;""></div>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div align=""right"" style=""float:right; width:250px;"">"
	sHTML = sHTML & "<p class=""m0"">���ԎQ�Ɖ񐔁F" & vAccessCount & "��@�X�V���F" & sUpdateDay & "</p>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""><img src=""/img/order/tab_border.gif"" alt="""" width=""600"" height=""5""></div>"
	sHTML = sHTML & "</div>" & vbCrLf

	GetHTMLEditOrderShowTypeSwitch = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[�̃L���b�`�R�s�[�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit0.asp
'�X�@�V�F2008/10/10 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHTMLEditOrderCatchCopy(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	'<�ϐ��錾>
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode
	Dim dbOptionNo		'�傫���ʐ^�̔ԍ�
	Dim dbJobTypeDetail	'��̓I�E�햼
	Dim dbCatchCopy		'�L���b�`�R�s�[

	Dim sHTML
	Dim sImg1
	Dim sClass
	Dim sImgOrderSpeciality
	'</�ϐ��錾>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")
	dbJobTypeDetail = ChkStr(rRS.Collect("JobTypeDetail"))
	dbCatchCopy = ChkStr(rRS.Collect("CatchCopy"))
	sImgOrderSpeciality = GetImgOrderSpeciality(rDB, rRS)

	If dbJobTypeDetail = "" Then dbJobTypeDetail = "<span style=""color:#999999;"">[��̓I�E�햼]�������͂ł��B</span>"
	If dbCatchCopy = "" Then dbCatchCopy = "<span style=""color:#999999;"">[�L���b�`�R�s�[]�������͂ł��B</span>"
	If sImgOrderSpeciality = "" Then sImgOrderSpeciality = "<span style=""color:#999999;"">[��W�̓���]�������͂ł��B</span>"

	'******************************************************************************
	'�傫���摜 start
	'------------------------------------------------------------------------------
	dbImageLimit = rRS.Collect("ImageLimit")
	dbOptionNo = ""
	sImg1 = ""
	If dbImageLimit > 0 Then
		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
				dbOptionNo = oRS.Collect("OptionNo1")
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & dbOptionNo
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

	If dbImageLimit > 0 Then
		sHTML = sHTML & "<div id=""catchcopy"">" & vbCrLf

		'<�E��>
		sHTML = sHTML & "<div class=""left"">"
		'
		If dbImageLimit = 1 Then
			sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" style=""margin-left:1px;"" onclick=""location.href='" & HTTP_CURRENTURL & "company/img_upload.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
		Else
			sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" style=""margin-left:1px;"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=001&amp;appcode=" & vApplicationCode & "';"">"
		End If
		sHTML = sHTML & "<br>"

		If sImg1 <> "" Then
			sHTML = sHTML & "<div class=""main_pics""><img class=""big"" src=""" & sImg1 & """ alt="""" border=""1"" id=""big_pics""></div>"
		Else
			sHTML = sHTML & "<img class=""big"" src=""" & sImg1 & """ alt=""[�ʐ^]�����o�^�ł��B"" border=""1"" width=""300"" height=""225"" style=""border:1px solid #999999;"">"
		End If
		sHTML = sHTML & "</div>"
		'</�E��>

		'<����>
		sHTML = sHTML & "<div style=""float:right; width:298px;"">"
		sHTML = sHTML & "<a name=""edit00""></a>"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit08.asp?ordercode=" & dbOrderCode & "&amp;place=1&amp;appcode=" & vApplicationCode & "';""><span style=""color:#ff0000; font-weight:normal; font-size:9px;"">���K�{</span><br>"
		sHTML = sHTML & "<h2 style=""margin-bottom:15px;"">" & dbJobTypeDetail & "</h2><br>"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit01.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
		sHTML = sHTML & "<a name=""edit01""></a>"
		sHTML = sHTML & "<p class=""m0"" style=""margin-bottom:15px;"">" & dbCatchCopy & "</p><br>"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit02.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
		sHTML = sHTML & "<a name=""edit02""></a>"
		sHTML = sHTML & "<div style=""margin:10px 0px;"">"
		If sImgOrderSpeciality <> "" Then
			sHTML = sHTML & "<div style=""border:solid 0px #cccccc;padding:5px;"">"
			sHTML = sHTML & "<div style=""font-size:12px;font-weight:normal;color:#008900;"">�y��W�̓����z</div>"
			sHTML = sHTML & sImgOrderSpeciality
			sHTML = sHTML & "</div>"
		End If
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "</div>"
		'<����>

		'float����
		sHTML = sHTML & "<br clear=""all"">"

		sHTML = sHTML & "</div>"
	Else
		sHTML = sHTML & "<div id=""catchcopy"" style=""width:600px;"">"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit08.asp?ordercode=" & dbOrderCode & "&amp;place=1&amp;appcode=" & vApplicationCode & "';""><span style=""color:#ff0000; font-weight:normal; font-size:9px;"">���K�{</span><br>"
		sHTML = sHTML & "<h2 style=""width:600px;"">" & dbJobTypeDetail & "</h2><br>"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit01.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>"
		sHTML = sHTML & "<a name=""edit01""></a>"
		sHTML = sHTML & "<p class=""m0"">" & dbCatchCopy & "</p><div style=""margin-top:10px;clear:both;""></div>"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit02.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
		sHTML = sHTML & "<a name=""edit02""></a>"
		sHTML = sHTML & "<div style=""margin-bottom:10px;"">"
		If sImgOrderSpeciality <> "" Then
			sHTML = sHTML & "<div style=""border:solid 0px #cccccc;padding:5px;"">"
			sHTML = sHTML & "<div style=""font-size:12px;font-weight:normal;color:#008900;"">�y��W�̓����z</div>"
			sHTML = sHTML & sImgOrderSpeciality
			sHTML = sHTML & "</div>"
		End If
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "</div>"
	End If

	GetHTMLEditOrderCatchCopy = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l��Ɖ摜�ꗗ�\���g�s�l�k�\��
'���@���FrDB			�F�ڑ����c�a�I�u�W�F�N�g
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvCategoryCode	�F�J�e�S���R�[�h
'�@�@�@�FvEditFlag		�F�ҏW�t���O
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit0.asp
'�X�@�V�F2008/10/10 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHTMLEditOrderPictureNow(ByRef rDB, ByRef rRS, ByVal vCategoryCode, ByVal vApplicationCode)
	'<�ϐ��錾>
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode
	Dim dbCompanyCode
	Dim dbImageLimit
	Dim dbOptionNo2
	Dim dbOptionNo3
	Dim dbOptionNo4
	Dim dbCaption2
	Dim dbCaption3
	Dim dbCaption4

	Dim sHTML
	Dim sURL
	Dim flgExistsPic
	'</�ϐ��錾>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbCompanyCode = rRS.Collect("CompanyCode")
	dbImageLimit = rRS.Collect("ImageLimit")
	flgExistsPic = False

	If dbImageLimit > 1 Then
		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', '" & vCategoryCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			dbOptionNo2 = oRS.Collect("OptionNo2")
			dbOptionNo3 = oRS.Collect("OptionNo3")
			dbOptionNo4 = oRS.Collect("OptionNo4")
			dbCaption2 = ChkStr(oRS.Collect("Caption2"))
			dbCaption3 = ChkStr(oRS.Collect("Caption3"))
			dbCaption4 = ChkStr(oRS.Collect("Caption4"))

			If Len(dbOptionNo2) > 0 Or Len(dbOptionNo3) > 0 Or Len(dbOptionNo4) > 0 Then
				flgExistsPic = True

				sHTML = sHTML & "<div id=""sub_pics"">"
				sHTML = sHTML & "<div style=""width:580px;margin:0 auto;"">"
				sURL = ""
				sHTML = sHTML & "<div align=""right"" style=""float:left; width:190px;"">"
				sHTML = sHTML & "<div align=""center"" style=""margin-bottom:3px;"">"
				sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=002&amp;appcode=" & vApplicationCode & "';"">"
				If dbOptionNo2 > 0 Then
					sHTML = sHTML & "&nbsp;<input class=""btn1"" type=""button"" value=""��@��"" onclick=""if(confirm('�ʐ^���O���܂����H') === true)location.href='" & HTTP_CURRENTURL & "company/order_img_relationdelete.asp?ordercode=" & dbOrderCode & "&amp;code=002&amp;appcode=" & vApplicationCode & "';"">"
				End If
				sHTML = sHTML & "</div>"
				If Len(oRS.Collect("OptionNo2")) > 0 Then
					sURL = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & dbOptionNo2
					sHTML = sHTML & "<div style=""width:182px; background-color:#ffffff;""><img src=""" & sURL & """ alt=""" & dbCaption2 & """ width=""180"" height=""135"" border=""1"" style=""border:1px solid #999999;""></div>"
				Else
					sHTML = sHTML & "<div style=""width:182px; background-color:#ffffff;""><table border=""1"" style=""width:180px; height:135px; border:1px solid #999999;""><tr><td></td></tr></table></div>"
				End If
				sHTML = sHTML & "<p class=""m0"" align=""left"" style=""width:182px; font-size:10px;"">" & dbCaption2 & "</p>"
				sHTML = sHTML & "</div>"

				sURL = ""
				sHTML = sHTML & "<div align=""right"" style=""float:left; width:190px;"">"
				sHTML = sHTML & "<div align=""center"" style=""margin-bottom:3px;"">"
				sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=003&amp;appcode=" & vApplicationCode & "';"">"
				If dbOptionNo3 > 0 Then
					sHTML = sHTML & "&nbsp;<input class=""btn1"" type=""button"" value=""��@��"" onclick=""if(confirm('�ʐ^���O���܂����H') === true)location.href='" & HTTP_CURRENTURL & "company/order_img_relationdelete.asp?ordercode=" & dbOrderCode & "&amp;code=003';"">"
				End If
				sHTML = sHTML & "</div>"
				If Len(oRS.Collect("OptionNo3")) > 0 Then
					sURL = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & dbOptionNo3
					sHTML = sHTML & "<div style=""width:182px; background-color:#ffffff;""><img src=""" & sURL & """ alt=""" & dbCaption3 & """ width=""180"" height=""135"" border=""1"" style=""border:1px solid #999999;""></div>"
				Else
					sHTML = sHTML & "<div style=""width:182px; background-color:#ffffff;""><table border=""1"" style=""width:180px; height:135px; border:1px solid #999999;""><tr><td></td></tr></table></div>"
				End If
				sHTML = sHTML & "<p class=""m0"" align=""left"" style=""width:182px; font-size:10px;"">" & dbCaption3 & "</p>"
				sHTML = sHTML & "</div>"

				sURL = ""
				sHTML = sHTML & "<div align=""right"" style=""float:left; width:190px;"">"
				sHTML = sHTML & "<div align=""center"" style=""margin-bottom:3px;"">"
				sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=004&amp;appcode=" & vApplicationCode & "';"">"
				If dbOptionNo4 > 0 Then
					sHTML = sHTML & "&nbsp;<input class=""btn1"" type=""button"" value=""��@��"" onclick=""if(confirm('�ʐ^���O���܂����H') === true)location.href='" & HTTP_CURRENTURL & "company/order_img_relationdelete.asp?ordercode=" & dbOrderCode & "&amp;code=004';"">"
				End If
				sHTML = sHTML & "</div>"
				If Len(oRS.Collect("OptionNo4")) > 0 Then
					sURL = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & dbOptionNo4
					sHTML = sHTML & "<div style=""width:182px; background-color:#ffffff;""><img src=""" & sURL & """ alt=""" & dbCaption4 & """ width=""180"" height=""135"" border=""1"" style=""border:1px solid #999999;""></div>"
				Else
					sHTML = sHTML & "<div style=""width:182px; background-color:#ffffff;""><table border=""1"" style=""width:180px; height:135px; border:1px solid #999999;""><tr><td></td></tr></table></div>"
				End If
				sHTML = sHTML & "<p class=""m0"" align=""left"" style=""width:182px; font-size:10px;"">" & dbCaption4 & "</p>"
				sHTML = sHTML & "</div>"

				sHTML = sHTML & "<br clear=""all"">"
				sHTML = sHTML & "</div>"
				sHTML = sHTML & "</div><br>"
			End If
		End If
	End If

	If flgExistsPic = False Then
		sHTML = sHTML & "<div align=""center"" id=""sub_pics"">"
		sHTML = sHTML & "<div style=""width:580px;margin:0 auto;"">"

		sHTML = sHTML & "<div align=""right"" style=""float:left; width:190px;"">"
		sHTML = sHTML & "<div align=""center"" style=""margin-bottom:3px;""><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=002&amp;appcode=" & vApplicationCode & "';""></div>"
		sHTML = sHTML & "<div style=""width:182px; background-color:#ffffff;""><table border=""1"" style=""width:180px; height:135px; border:1px solid #999999;""><tr><td></td></tr></table></div>"
		sHTML = sHTML & "<p class=""m0"" align=""left"" style=""width:182px; font-size:10px;""></p>"
		sHTML = sHTML & "</div>"

		sHTML = sHTML & "<div align=""right"" style=""float:left; width:190px;"">"
		sHTML = sHTML & "<div align=""center"" style=""margin-bottom:3px;""><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=003&amp;appcode=" & vApplicationCode & "';""></div>"
		sHTML = sHTML & "<div style=""width:182px; background-color:#ffffff;""><table border=""1"" style=""width:180px; height:135px; border:1px solid #999999;""><tr><td></td></tr></table></div>"
		sHTML = sHTML & "<p class=""m0"" align=""left"" style=""width:182px; font-size:10px;""></p>"
		sHTML = sHTML & "</div>"

		sHTML = sHTML & "<div align=""right"" style=""float:left; width:190px;"">"
		sHTML = sHTML & "<div align=""center"" style=""margin-bottom:3px;""><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=004&amp;appcode=" & vApplicationCode & "';""></div>"
		sHTML = sHTML & "<div style=""width:182px; background-color:#ffffff;""><table border=""1"" style=""width:180px; height:135px; border:1px solid #999999;""><tr><td></td></tr></table></div>"
		sHTML = sHTML & "<p class=""m0"" align=""left"" style=""width:182px; font-size:10px;""></p>"
		sHTML = sHTML & "</div>"

		sHTML = sHTML & "<br clear=""all"">"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "</div><br>"
	End If

	GetHTMLEditOrderPictureNow = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍ׃y�[�W�̃t���[�o�q���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvEditFlag		�F�ҏW�t���O
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit0.asp
'�X�@�V�F2008/10/10 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHTMLEditOrderFreePR(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	'<�ϐ��錾>
	Dim dbOrderCode		'���R�[�h
	Dim dbPRTitle1		'�o�q�^�C�g��1
	Dim dbPRTitle2		'�o�q�^�C�g��2
	Dim dbPRTitle3		'�o�q�^�C�g��3
	Dim dbPRContents1	'�o�q��1
	Dim dbPRContents2	'�o�q��2
	Dim dbPRContents3	'�o�q��3

	Dim sHTML
	'</�ϐ��錾>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbPRTitle1 = ChkStr(rRS.Collect("PRTitle1"))
	dbPRTitle2 = ChkStr(rRS.Collect("PRTitle2"))
	dbPRTitle3 = ChkStr(rRS.Collect("PRTitle3"))
	dbPRContents1 = Replace(ChkStr(rRS.Collect("PRContents1")), vbCrLf, "<br>")
	dbPRContents1 = Replace(dbPRContents1, vbCr, "<br>")
	dbPRContents1 = Replace(dbPRContents1, vbLf, "<br>")
	dbPRContents2 = Replace(ChkStr(rRS.Collect("PRContents2")), vbCrLf, "<br>")
	dbPRContents2 = Replace(dbPRContents2, vbCr, "<br>")
	dbPRContents2 = Replace(dbPRContents2, vbLf, "<br>")
	dbPRContents3 = Replace(ChkStr(rRS.Collect("PRContents3")), vbCrLf, "<br>")
	dbPRContents3 = Replace(dbPRContents3, vbCr, "<br>")
	dbPRContents3 = Replace(dbPRContents3, vbLf, "<br>")

	If dbPRTitle1 = "" Then dbPRTitle1 = "<span style=""color:#999999;"">[�o�q�P�^�C�g��]�������͂ł��B</span>"
	If dbPRTitle2 = "" Then dbPRTitle2 = "<span style=""color:#999999;"">[�o�q�Q�^�C�g��]�������͂ł��B</span>"
	If dbPRTitle3 = "" Then dbPRTitle3 = "<span style=""color:#999999;"">[�o�q�R�^�C�g��]�������͂ł��B</span>"
	If dbPRContents1 = "" Then dbPRContents1 = "<span style=""color:#999999;"">[�o�q�P���e]�������͂ł��B</span>"
	If dbPRContents2 = "" Then dbPRContents2 = "<span style=""color:#999999;"">[�o�q�Q���e]�������͂ł��B</span>"
	If dbPRContents3 = "" Then dbPRContents3 = "<span style=""color:#999999;"">[�o�q�R���e]�������͂ł��B</span>"

	sHTML = sHTML & "<a name=""edit03""></a>"
	sHTML = sHTML & "<h3>�o�q</h3>"
	sHTML = sHTML & "<div class=""freeprblock"">"

	sHTML = sHTML & "<div style=""margin-left:15px;""><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit03.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""></div>"

	If dbPRTitle1 <> "" Or dbPRContents1 <> "" Then
		sHTML = sHTML & "<h4>" & dbPRTitle1 & "</h4>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
		sHTML = sHTML & "<p class=""m0"">" & dbPRContents1 & "</p>"
	End If

	If dbPRTitle2 <> "" Or dbPRContents2 <> "" Then
		sHTML = sHTML & "<h4>" & dbPRTitle2 & "</h4>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
		sHTML = sHTML & "<p class=""m0"">" & dbPRContents2 & "</p>"
	End If

	If dbPRTitle3 <> "" Or dbPRContents3 <> "" Then
		sHTML = sHTML & "<h4>" & dbPRTitle3 & "</h4>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
		sHTML = sHTML & "<p class=""m0"">" & dbPRContents3 & "</p>"
	End If

	sHTML = sHTML & "</div>"

	GetHTMLEditOrderFreePR = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[�̗̍p�̔w�i���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvEditFlag		�F�ҏW�t���O
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit0.asp
'�X�@�V�F2008/10/10 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHTMLEditOrderBackGround(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	'<�ϐ��錾>
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode
	Dim dbOrderBackGround	'�̗p�̔w�i

	Dim sHTML
	'</�ϐ��錾>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	'�̗p�̔w�i�擾
	dbOrderBackGround = Replace(ChkStr(rRS.Collect("OrderBackGround")), vbCrLf, "<br>")

	If dbOrderBackGround = "" Then dbOrderBackGround = "<span style=""color:#999999;"">[�̗p�̔w�i]�������͂ł��B</span>"


	sHTML = sHTML & "<a name=""edit04""></a>"
	sHTML = sHTML & "<h3>�̗p�̔w�i</h3>" & vbCrLf
	sHTML = sHTML & "<div style=""margin-left:15px;""><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit04.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""></div>"

	'�̗p�̔w�i�o��
	sHTML = sHTML & "<p class=""m0"" style=""padding-left:15px;"">" & dbOrderBackGround & "</p>" & vbCrLf

	sHTML = sHTML & "<br>"

	GetHTMLEditOrderBackGround = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[�̋Ɩ����e���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvEditFlag		�F�ҏW�t���O
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit0.asp
'�X�@�V�F2008/10/10 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHTMLEditBusiness(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode		'���R�[�h
	Dim dbPlanType		'���l�[���C�Z���X�v�������
	Dim dbBizName1		'�d����������1
	Dim dbBizName2		'�d����������2
	Dim dbBizName3		'�d����������3
	Dim dbBizName4		'�d����������4
	Dim dbBizPercentage1'�d������1
	Dim dbBizPercentage2'�d������2
	Dim dbBizPercentage3'�d������3
	Dim dbBizPercentage4'�d������4
	Dim dbBusinessDetail'�S���Ɩ�
	Dim sClearSolid
	Dim flgLine			'�������t���O

	Dim sHTML
	Dim sBiz			'�d������HTML

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	dbOrderCode = rRS.Collect("OrderCode")
	dbPlanType = rRS.Collect("PlanTypeName")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'******************************************************************************
	'�d���̊��� start
	'------------------------------------------------------------------------------
	sBiz = ""
	dbBizName1 = ""
	dbBizName2 = ""
	dbBizName3 = ""
	dbBizName4 = ""
	dbBizPercentage1 = ""
	dbBizPercentage2 = ""
	dbBizPercentage3 = ""
	dbBizPercentage4 = ""

	dbBizName1 = ChkStr(rRS.Collect("BizName1"))
	dbBizName2 = ChkStr(rRS.Collect("BizName2"))
	dbBizName3 = ChkStr(rRS.Collect("BizName3"))
	dbBizName4 = ChkStr(rRS.Collect("BizName4"))
	dbBizPercentage1 = ChkStr(rRS.Collect("BizPercentage1"))
	dbBizPercentage2 = ChkStr(rRS.Collect("BizPercentage2"))
	dbBizPercentage3 = ChkStr(rRS.Collect("BizPercentage3"))
	dbBizPercentage4 = ChkStr(rRS.Collect("BizPercentage4"))

	If dbBizName1 = "" Then dbBizName1 = "<span style=""color:#999999;"">[�d���̊����P]�������͂ł��B</span>"
	If dbBizName2 = "" Then dbBizName2 = "<span style=""color:#999999;"">[�d���̊����Q]�������͂ł��B</span>"
	If dbBizName3 = "" Then dbBizName3 = "<span style=""color:#999999;"">[�d���̊����R]�������͂ł��B</span>"
	If dbBizName4 = "" Then dbBizName4 = "<span style=""color:#999999;"">[�d���̊����S]�������͂ł��B</span>"

	If dbBizName1 & dbBizName2 & dbBizName3 & dbBizName4 <> "" Then
		If dbBizName1 <> "" And dbBizPercentage1 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & dbBizName1 & "</td><td class=""biz2"">" & dbBizPercentage1 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(dbBizPercentage1) * 3 & """ height=""20""></td></tr>"
		If dbBizName2 <> "" And dbBizPercentage2 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & dbBizName2 & "</td><td class=""biz2"">" & dbBizPercentage2 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(dbBizPercentage2) * 3 & """ height=""20""></td></tr>"
		If dbBizName3 <> "" And dbBizPercentage3 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & dbBizName3 & "</td><td class=""biz2"">" & dbBizPercentage3 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(dbBizPercentage3) * 3 & """ height=""20""></td></tr>"
		If dbBizName4 <> "" And dbBizPercentage4 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & dbBizName4 & "</td><td class=""biz2"">" & dbBizPercentage4 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(dbBizPercentage4) * 3 & """ height=""20""></td></tr>"
		sBiz = "<table>" & sBiz & "</table>"
	End If
	'------------------------------------------------------------------------------
	'�d���̊��� end
	'******************************************************************************

	'******************************************************************************
	'�S���Ɩ� start
	'------------------------------------------------------------------------------
	dbBusinessDetail = Replace(ChkStr(rRS.Collect("BusinessDetail")), vbCrLf, "<br>")
	dbBusinessDetail = Replace(dbBusinessDetail, vbCr, "<br>")
	dbBusinessDetail = Replace(dbBusinessDetail, vbLf, "<br>")
	If dbBusinessDetail = "" Then dbBusinessDetail = "<span style=""color:#999999;"">[�S���Ɩ�]�������͂ł��B</span>"
	'------------------------------------------------------------------------------
	'�S���Ɩ� end
	'******************************************************************************

	sHTML = sHTML & "<h3>�Ɩ����e</h3>"

	flgLine = False

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<a name=""edit05""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit05.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>�S���Ɩ�</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbBusinessDetail & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	If (dbPlanType = "platinum" Or dbPlanType = "gold" Or dbPlanType = "old") Then
		If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
		sHTML = sHTML & "<a name=""edit06""></a>"
		sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit06.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>�d���̊���</h4></div>"
		'sHTML = sHTML & "<div class=""value1"">" & sBiz & "</div>"
		sHTML = sHTML & "<div class=""value1"">"
		sHTML = sHTML & "<table border=""0"">"
		sHTML = sHTML & "<tbody>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td>"
		sHTML = sHTML & "<script type=""text/javascript"" language=""javascript"">"
		sHTML = sHTML & "viewWorkAvg(" & dbBizPercentage1 & ", " & dbBizPercentage2 & ", " & dbBizPercentage3 & ", " & dbBizPercentage4 & ")"
		sHTML = sHTML & "</script>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td style=""padding-left:5px; vertical-align:middle;"">"
		sHTML = sHTML & "<table border=""0"">"
		sHTML = sHTML & "<tbody>"
		If dbBizName1 <> "" Then sHTML = sHTML & "<tr><td style=""width:16px; background-color:#ff9999; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & dbBizPercentage1 & "%</td><td>" & dbBizName1 & "</td></tr>"
		If dbBizName2 <> "" Then sHTML = sHTML & "<tr><td style=""width:16px; background-color:#9999ff; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & dbBizPercentage2 & "%</td><td>" & dbBizName2 & "</td></tr>"
		If dbBizName3 <> "" Then sHTML = sHTML & "<tr><td style=""width:16px; background-color:#99ff99; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & dbBizPercentage3 & "%</td><td>" & dbBizName3 & "</td></tr>"
		If dbBizName4 <> "" Then sHTML = sHTML & "<tr><td style=""width:16px; background-color:#ffff99; border-bottom:1px solid #ffffff;""></td><td style=""padding:0px 5px;"">" & dbBizPercentage4 & "%</td><td>" & dbBizName4 & "</td></tr>"
		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>"
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
	End If

	sHTML = sHTML & "<br>"
	sHTML = sHTML & "<br>" & vbCrLf

	GetHTMLEditBusiness = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[�̋Ζ��������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit0.asp
'���@���F2008/10/10 LIS K.Kokubo �쐬
'�@�@�@�F2009/04/16 LIS K.Kokubo ���[���ۋ����C�Z���X�̏ꍇ�͋Ζ��n�̕\������ʂ̋��l�L���ł��s��S�܂ł����\�������Ȃ�
'�@�@�@�F2009/04/22 LIS K.Kokubo �Љ��̋Ζ��`��(TTP�p)�Ή�
'�@�@�@�F2009/11/02 LIS K.Kokubo �r�n�g�n,�e�b�̋Ζ��n�\���Ή�
'******************************************************************************
Function GetHTMLEditCondition(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
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
	Dim dbPlanTypeName		'�i�r���C�Z���X���
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
	Dim sPlusMinus
	Dim flgFC				'FC�E�㗝�X�t���O
	Dim flgSOHOFC
	'</�ϐ��錾>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyKbn = rRS.Collect("CompanyKbn")
	dbPlanTypeName = rRS.Collect("PlanTypeName")
	dbTTPOrderFlag = rRS.Collect("TTPOrderFlag")

	'<�Ζ��`��>
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

	If sAfterWorkingType = "" Then sAfterWorkingType = "<span style=""color:#999999;"">[�Љ��̋Ζ��`��]�������͂ł��B</span>"
	If sWorkRange = "" Then sWorkRange = "<span style=""color:#999999;"">[�A�Ɗ���]�������͂ł��B</span>"
	If sWorkingType = "" Then sWorkingType = "<span style=""color:#999999;"">[�Ζ��`��]�������͂ł��B</span>"
	'</�Ζ��`��>

	'<�E��>
	sJobType = GetJobType(rDB, rRS)
	If sJobType = "" Then sJobType = "<span style=""color:#999999;"">[�E��]�������͂ł��B</span>"
	'</�E��>

	'<�E��ڍ�>
	dbJobTypeDetail = rRS.Collect("JobTypeDetail")
	'</�E��ڍ�>

	'<���^>
	flgFC = False
	'<�e�b�E�㗝�X�`�F�b�N>
	sSQL = "sp_GetDataWorkingType '" & qsOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		If oRS.Collect("WorkingTypeCode") = "006" Or oRS.Collect("WorkingTypeCode") = "007" Then flgFC = True
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
	'</�e�b�E�㗝�X�`�F�b�N>

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

	'���^
	If sYearlyIncome = "" Then sYearlyIncome = "<span style=""color:#999999;"">[�N��]�������͂ł��B</span>"
	If sMonthlyIncome = "" Then sMonthlyIncome = "<span style=""color:#999999;"">[����]�������͂ł��B</span>"
	If sDailyIncome = "" Then sDailyIncome = "<span style=""color:#999999;"">[����]�������͂ł��B</span>"
	If sHourlyIncome = "" Then sHourlyIncome = "<span style=""color:#999999;"">[����]�������͂ł��B</span>"

	'������
	If dbPercentagePay <> "" Then
		If dbPercentagePay = "1" Then
			dbPercentagePay = "����"
		ElseIf dbPercentagePay = "0" Then
			dbPercentagePay = "�Ȃ�"
		End If
	Else
		dbPercentagePay = "<span style=""color:#999999;"">[������]�������͂ł��B</span>"
	End If

	'��ʔ�
	If ChkStr(rRS.Collect("NaviTrafficPayFlag")) = "1" Then 
		sTrafficFee = "��ʔ�x������" & dbTrafficFeeType
		If IsNumber(dbTrafficFeeMonth, 0, False) = True Then
			sTrafficFee = sTrafficFee & "(" & FormatCanma(dbTrafficFeeMonth) & "�~�^��)"
		End If
	Else
		sTrafficFee = "<span style=""color:#999999;"">[��ʔ�]�������͂ł��B</span>"
	End If

	If dbSalaryRemark = "" Then dbSalaryRemark = "<span style=""color:#999999;"">[���^���l]�������͂ł��B</span>"
	'</���^>

	'<����>
	sWorkingTime = GetWorkingTime(rDB, rRS)
	dbWorkTimeRemark = ChkStr(rRS.Collect("WorkTimeRemark"))

	If sWorkingTime = "" Then sWorkingTime = "<span style=""color:#999999;"">[�A�Ǝ���]�������͂ł��B</span>"
	If dbWorkTimeRemark = "" Then dbWorkTimeRemark = "<span style=""color:#999999;"">[�A�Ǝ��Ԕ��l]�������͂ł��B</span>"
	'</����>

	'<�x��>
	dbWeeklyHolidayType = ChkStr(rRS.Collect("WeeklyHolidayTypeName"))
	dbHolidayRemark = ChkStr(rRS.Collect("HolidayRemark"))

	If dbWeeklyHolidayType = "" Then dbWeeklyHolidayType = "<span style=""color:#999999;"">[�T�x���]�������͂ł��B</span>"
	If dbHolidayRemark = "" Then dbHolidayRemark = "<span style=""color:#999999;"">[�x�����l]�������͂ł��B</span>"
	'</�x��>

	'<��W�l��>
	dbHumanNumber = ChkStr(rRS.Collect("HumanNumber"))

	If dbHumanNumber <> "" Then
		dbHumanNumber = dbHumanNumber & "�l"
	Else
		dbHumanNumber = "<span style=""color:#999999;"">[��W�l��]�������͂ł��B</span>"
	End If
	'</��W�l��>

	'<�Ζ��n>
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
			sWorkingPlace = sWorkingPlace & "<p class=""m0"""
			If flgSOHOFC = True Then
				sWorkingPlace = sWorkingPlace & " style=""display:inline;"""
			Else
				sWorkingPlace = sWorkingPlace & " style=""padding-left:15px;"""
			End If
			sWorkingPlace = sWorkingPlace & ">"
			sWorkingPlace = sWorkingPlace & "[�Ŋ�w]"
			sWorkingPlace = sWorkingPlace & sNearbyStation
			sWorkingPlace = sWorkingPlace & "<br>"
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

	If sWorkingPlace = "" Then sWorkingPlace = "<span style=""color:#999999;"">[�Ζ��n]�������͂ł��B</span>"
	If dbTransfer = "" Then dbTransfer = "<span style=""color:#999999;"">[�]�΂̗L��]�������͂ł��B</span>"
	If sMAP = "" Then sMAP = "<span style=""color:#999999;"">[�A�Ɛ�̒n�}�ʒu���]�����o�^�ł��B</span>"
	'</�Ζ��n>

	sHTML = sHTML & "<h3>�Ζ�����</h3>"

	sHTML = sHTML & "<div class=""category1"">"
	sHTML = sHTML & "<h4>"
	sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit07.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
	If G_COMPANYKBN = "1" Then sHTML = sHTML & "<span style=""color:#ff0000; font-weight:normal; font-size:9px;"">���K�{</span>"
	sHTML = sHTML & "<br>"
	sHTML = sHTML & "�Ζ��`��"
	sHTML = sHTML & "</h4>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div class=""value1"">"
	sHTML = sHTML & "<a name=""edit07""></a>"
	'<�Ζ��`��>
	sHTML = sHTML & "<p class=""m0"">" & sWorkingType & "</p>"
	'</�Ζ��`��>
	'<�Љ��̋Ζ��`��>
	If dbTTPOrderFlag = "1" Then sHTML = sHTML & "<p class=""m0"">" & sAfterWorkingType & "</p>"
	'</�Љ��̋Ζ��`��>
	'<�A�Ɗ���>
	sHTML = sHTML & "<p class=""m0"">" & sWorkRange & "</p>"
	'</�A�Ɗ���>
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<�E��>
	sHTML = sHTML & "<a name=""edit08""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit08.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><span style=""color:#ff0000; font-weight:normal; font-size:9px;"">���K�{</span><br>�E��</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"
	sHTML = sHTML & "<p class=""m0""><strong>" & dbJobTypeDetail & "</strong></p>"
	sHTML = sHTML & "<p class=""m0"">" & sJobType & "</p>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</�E��>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit09.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
	If flgFC = False Then sHTML = sHTML & "<span style=""color:#ff0000; font-weight:normal; font-size:9px;"">���K�{</span>"
	sHTML = sHTML & "<br>���^</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"
	sHTML = sHTML & "<a name=""edit09""></a>"
	If flgFC = True Then sHTML = sHTML & "<p class=""m0"" style=""color:#999999;"">���e�b�E�㗝�X�A�r�n�g�n��W�̂��߁A���^�̓��͕͂K�{�ł͂���܂���B</p>"
	'<�N��>
	sHTML = sHTML & "<h5>�N��</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sYearlyIncome & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</�N��>

	'<����>
	sHTML = sHTML & "<h5>����</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sMonthlyIncome & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</����>

	'<����>
	sHTML = sHTML & "<h5>����</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sDailyIncome & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</����>

	'<����>
	sHTML = sHTML & "<h5>����</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sHourlyIncome & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</����>

	'<������>
	sHTML = sHTML & "<h5>������</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbPercentagePay & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both; margin:0px;""></div>"
	'</������>

	'<��ʔ�>
	sHTML = sHTML & "<h5>��ʔ�</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sTrafficFee & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</��ʔ�>

	'<���^���l>
	sHTML = sHTML & "<h5>���^���l</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbSalaryRemark & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both; margin:0px;""></div>"
	'</���^���l>

	If flgFC = False Then
		sHTML = sHTML & "<p class=""m0"" style=""font-size:10px;"">"
		sHTML = sHTML & "���Œ�z�͏����Ɋ֌W�Ȃ�������z�ł��B(�N���̍Œ�z�͏����Ɋ֌W�Ȃ������錎���̍��v�ł��B)"
		sHTML = sHTML & "</p>"
	End If
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit10.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
	'FC,SOHO�Č��ȊO�̏ꍇ�͕K�{
	If rRS.Collect("FCSOHOOrderFlag") = "0" Then
		sHTML = sHTML & "<span style=""color:#ff0000; font-weight:normal; font-size:9px;"">���K�{</span>"
	End If
	sHTML = sHTML & "<br>����</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"
	'<�A�Ǝ���>
	sHTML = sHTML & "<a name=""edit10""></a>"
	If flgFC = True Then sHTML = sHTML & "<p class=""m0"" style=""color:#999999;"">���e�b�E�㗝�X�A�r�n�g�n��W�̂��߁A�A�Ǝ��Ԃ̓��͕͂K�{�ł͂���܂���B</p>"
	sHTML = sHTML & "<h5>�A�Ǝ���</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sWorkingTime & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</�A�Ǝ���>

	'<�A�Ǝ��Ԕ��l>
	sHTML = sHTML & "<h5>�A�Ǝ��Ԕ��l</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbWorkTimeRemark & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</�A�Ǝ��Ԕ��l>

	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit11.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>�x��</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"

	'<�x��>
	sHTML = sHTML & "<a name=""edit11""></a>"
	sHTML = sHTML & "<h5>�x��</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbWeeklyHolidayType & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</�x��>

	'<�x�����l>
	sHTML = sHTML & "<h5>�x�����l</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbHolidayRemark & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</�x�����l>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<��W�l��>
	sHTML = sHTML & "<a name=""edit12""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit12.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>��W�l��</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"
	sHTML = sHTML & "<p class=""m0"">" & dbHumanNumber & "</p>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</��W�l��>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<�Ζ��n>
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit13.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><span style=""color:#ff0000; font-weight:normal; font-size:9px;"">���K�{</span><br>�Ζ��n</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"
	sHTML = sHTML & "<a name=""edit13""></a>"
	sHTML = sHTML & "<h5>�Z��</h5>"
	sHTML = sHTML & "<div class=""value2"">"
	sHTML = sHTML & "<p class=""m0"">" & sWorkingPlace & "</p>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	sHTML = sHTML & "<h5>�]��</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbTransfer & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</�Ζ��n>

	sHTML = sHTML & "<br>"

	GetHTMLEditCondition = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[�̕K�v�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvEditFlag		�F�ҏW�t���O
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit0.asp
'���@���F2008/10/10 LIS K.Kokubo �쐬
'�@�@�@�F2012/03/12 LIS K.Kokubo ���ƔN�o��
'******************************************************************************
Function GetHTMLEditNeedCondition(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	'<�ϐ��錾>
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode			'���R�[�h
	Dim dbOrderType			'���l�[���
	Dim dbCompanyKbn		'��Ƌ敪
	Dim dbTempOrderFlag		'�h���Č��t���O
	Dim dbAgeMin			'�N���
	Dim dbAgeMax			'�N����
	Dim dbAgeReasonFlag		'�N��R�t���O
	Dim dbAgeReason			'�N��R
	Dim dbAgeReasonDetail	'�N������R�ڍ�
	Dim dbHopeSchoolHistory	'�w��
	Dim dbGraduateYearMin	'���ƔN����
	Dim dbGraduateYearMax	'���ƔN���

	Dim sHTML
	Dim sAge				'�N���
	Dim sSchoolHistory		'�w��
	Dim sSkillOS			'�n�r
	Dim sSkillApp			'�A�v���P�[�V����
	Dim sSkillDL			'�J������
	Dim sSkillDB			'�c�a
	Dim sSkillOther			'���̑��X�L��
	Dim sLicense			'���i
	Dim sLicenseOther		'���̑����i
	Dim sOtherNote			'���̑����L����
	'</�ϐ��錾>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyKbn = rRS.Collect("CompanyKbn")
	dbTempOrderFlag = rRS.Collect("TempOrderFlag")

	'<�N��>
	sAge = ""
	dbAgeMin = ChkStr(rRS.Collect("AgeMin"))
	dbAgeMax = ChkStr(rRS.Collect("AgeMax"))
	dbAgeReasonFlag = ChkStr(rRS.Collect("AgeReasonFlag"))
	dbAgeReason = ChkStr(rRS.Collect("AgeReason"))
	dbAgeReasonDetail = Replace(ChkStr(rRS.Collect("AgeReasonDetail")), vbCrLf, "<br>")

	If dbOrderType = "1" Or dbTempOrderFlag = "1" Then
		sAge = "�h���Č��̂��߁A�N��f�ڂ��Ă��܂���B<br>"
		sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[�H]�����ɂ���</a>"
	ElseIf dbAgeReasonFlag = "0" Or dbAgeReasonFlag = "" Or (dbAgeMin & dbAgeMax = "") Then
		sAge = "�N��s��<br>"
		'sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[�H]�����ɂ���</a>"
	Else
		If dbAgeMin <> "" Then dbAgeMin = dbAgeMin & "��"
		If dbAgeMax <> "" Then dbAgeMax = dbAgeMax & "��"
		sAge = dbAgeMin & "�`" & dbAgeMax
		If dbAgeReason <> "" Then sAge = sAge & "&nbsp;(" & dbAgeReason & ")<br>"
		If dbAgeReasonDetail <> "" Then sAge = sAge & dbAgeReasonDetail & "<br>"
		sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[�H]�����ɂ���</a><br>"
	End If

	If dbTempOrderFlag = "1" Then
		sAge = "<span style=""color:#999999;"">[�N��֘A]�h���Č��̂��߁A�N��̓��͂͂ł��܂���B</span>"
	ElseIf sAge = "" Then
		sAge = "<span style=""color:#999999;"">[�N��֘A]�������͂ł��B</span>"
	End If
	'</�N��>

	'<�w��>
	dbHopeSchoolHistory = ChkStr(rRS.Collect("HopeSchoolHistory"))
	dbGraduateYearMin = rRS.Collect("GraduateYearMin")
	dbGraduateYearMax = rRS.Collect("GraduateYearMax")
	If dbHopeSchoolHistory <> "" Then
		sSchoolHistory = dbHopeSchoolHistory & "���ȏ�<br>"

		If dbGraduateYearMin + dbGraduateYearMax > 0 Then
			sSchoolHistory = sSchoolHistory & "[���ƔN] "
			If dbGraduateYearMin > 0 Then
				sSchoolHistory = sSchoolHistory & dbGraduateYearMin & "�N��"
			End If
			sSchoolHistory = sSchoolHistory & " �` "
			If dbGraduateYearMax > 0 Then
				sSchoolHistory = sSchoolHistory & dbGraduateYearMax & "�N��"
			End If
		Else
			sSchoolHistory = sSchoolHistory & "<span style=""color:#999999;"">[���ƔN]�������͂ł��B</span>"
		End If
	Else
		sSchoolHistory = "<span style=""color:#999999;"">[�w��]�������͂ł��B</span><br>"
		sSchoolHistory = sSchoolHistory & "<span style=""color:#999999;"">[���ƔN]�������͂ł��B</span>"
	End If
	'</�w��>

	'<���i>
	sLicense = GetLicense(rDB, rRS)
	sLicenseOther = GetOrderNote(rDB, rRS, "OtherLicense")

	If sLicense = "" Then sLicense = "<span style=""color:#999999;"">[���i]�������͂ł��B</span>"
	If sLicenseOther = "" Then sLicenseOther = "<span style=""color:#999999;"">[���̑����i]�������͂ł��B</span>"
	'</���i>

	'<�X�L��>
	sSkillOS = GetSkill(rDB, rRS, "OS")
	sSkillApp = GetSkill(rDB, rRS, "Application")
	sSkillDL = GetSkill(rDB, rRS, "DevelopmentLanguage")
	sSkillDB = GetSkill(rDB, rRS, "Database")
	sSkillOther = GetOrderNote(rDB, rRS, "OtherSkill")

	If sSkillOS = "" Then sSkillOS = "<span style=""color:#999999;"">[�n�r]�������͂ł��B</span>"
	If sSkillApp = "" Then sSkillApp = "<span style=""color:#999999;"">[�A�v���P�[�V����]�������͂ł��B</span>"
	If sSkillDL = "" Then sSkillDL = "<span style=""color:#999999;"">[�J������]�������͂ł��B</span>"
	If sSkillDB = "" Then sSkillDB = "<span style=""color:#999999;"">[�f�[�^�x�[�X]�������͂ł��B</span>"
	If sSkillOther = "" Then sSkillOther = "<span style=""color:#999999;"">[���̑��X�L��]�������͂ł��B</span>"
	'</�X�L��>

	'<���̑����L����>
	sOtherNote = ""
	If dbOrderType = "0" Then
		sOtherNote = GetOrderNote(rDB, rRS, "OtherNote")
	End If

	If sOtherNote = "" Then sOtherNote = "<span style=""color:#999999;"">[���̑����L����]�������͂ł��B</span>"
	'</���̑����L����>

	sHTML = sHTML & "<h3>�K�v����</h3>" & vbCrLf

	'<�N��>
	sHTML = sHTML & "<a name=""edit14""></a>"
	sHTML = sHTML & "<div class=""category1"">"
	sHTML = sHTML & "<h4>"
	If dbTempOrderFlag = "0" Then sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit14.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>"
	sHTML = sHTML & "�N��"
	sHTML = sHTML & "</h4>"
	sHTML = sHTML & "</div>" & vbCrLf
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sAge & "</p></div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	'</�N��>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<��]�w��>
	sHTML = sHTML & "<a name=""edit15""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit15.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>��]�w��</h4></div>" & vbCrLf
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sSchoolHistory & "</p></div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	'</��]�w��>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<���i�o��>
	sHTML = sHTML & "<a name=""edit16""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit16.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>���i</h4></div>" & vbCrLf
	sHTML = sHTML & "<div class=""value1"">" & vbCrLf

	sHTML = sHTML & "<h5>���i</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2"">" & sLicense & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "<h5>���̑����i</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2"">" & sLicenseOther & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	'</���i�o��>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<�X�L���o��>
	sHTML = sHTML & "<a name=""edit17""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit17.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>�X�L��</h4></div>" & vbCrLf
	sHTML = sHTML & "<div class=""value1"">" & vbCrLf

	sHTML = sHTML & "<h5>�n�r</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2"">" & sSkillOS & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "<h5>���ع����</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2"">" & sSkillApp & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "<h5>�J������</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2"">" & sSkillDL & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "<h5>�f�[�^�x�[�X</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2"">" & sSkillDB & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "<h5>���̑��X�L��</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sSkillOther & "</p></div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	'</�X�L���o��>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<���̑����L����>
	sHTML = sHTML & "<a name=""edit18""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit18.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>���L����</h4></div>" & vbCrLf
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sOtherNote & "</p></div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	'</���̑����L����>

	sHTML = sHTML & "<br>"

	GetHTMLEditNeedCondition = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[�̉�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit0.asp
'�X�@�V�F2008/10/10 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHTMLEditHowToEntry(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	Dim dbOrderCode			'���R�[�h
	Dim dbEntryInfo			'������@
	Dim dbProcess1			'STEP1
	Dim dbProcess2			'STEP2
	Dim dbProcess3			'STEP3
	Dim dbProcess4			'STEP4
	Dim sCSectionName		'���X�S������
	Dim sCPersonName		'���X�S���Җ�
	Dim sCTel				'���X�A����
	Dim sLis				'���X�S����
	Dim dbWValueURL			'�v�o�����[�̎��Ѝ̗p�y�[�W�t�q�k
	Dim sClearSolid
	Dim flgLine				'�������t���O

	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sOrderType = ChkStr(rRS.Collect("OrderType"))
	dbOrderCode = ChkStr(rRS.Collect("OrderCode"))
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'******************************************************************************
	'������@ start
	'------------------------------------------------------------------------------
	dbEntryInfo = Replace(ChkStr(rRS.Collect("EntryInfo")), vbCrLf, "<br>")
	dbEntryInfo = Replace(dbEntryInfo, vbCr, "<br>")
	dbEntryInfo = Replace(dbEntryInfo, vbLf, "<br>")

	If dbEntryInfo = "" Then dbEntryInfo = "<span style=""color:#999999;"">[������@]�������͂ł��B</span>"
	'------------------------------------------------------------------------------
	'������@ end
	'******************************************************************************

	'******************************************************************************
	'�I�l�菇 start
	'------------------------------------------------------------------------------
	dbProcess1 = ChkStr(rRS.Collect("Process1"))
	dbProcess2 = ChkStr(rRS.Collect("Process2"))
	dbProcess3 = ChkStr(rRS.Collect("Process3"))
	dbProcess4 = ChkStr(rRS.Collect("Process4"))

	If dbProcess1 = "" Then dbProcess1 = "<span style=""color:#999999;"">[�I�l�菇�P]�������͂ł��B</span>"
	If dbProcess2 = "" Then dbProcess2 = "<span style=""color:#999999;"">[�I�l�菇�Q]�������͂ł��B</span>"
	If dbProcess3 = "" Then dbProcess3 = "<span style=""color:#999999;"">[�I�l�菇�R]�������͂ł��B</span>"
	If dbProcess4 = "" Then dbProcess4 = "<span style=""color:#999999;"">[�I�l�菇�S]�������͂ł��B</span>"
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
	'------------------------------------------------------------------------------
	'�A���� end
	'******************************************************************************

	'******************************************************************************
	'�v�o�����[�̎��Ѝ̗p�y�[�W�t�q�k start
	'------------------------------------------------------------------------------
	dbWValueURL = ChkStr(rRS.Collect("WValueURL"))
	'------------------------------------------------------------------------------
	'�v�o�����[�̎��Ѝ̗p�y�[�W�t�q�k end
	'******************************************************************************

	flgLine = False

	sHTML = sHTML & "<a name=""edit19""></a>"
	sHTML = sHTML & "<h3>������</h3>" & vbCrLf
	sHTML = sHTML & "<div style=""margin-left:15px;""><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit19.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""></div>"

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>���R�[�h</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbOrderCode & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>������@</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbEntryInfo & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>�I�l�菇</h4></div>" & vbCrLf
	sHTML = sHTML & "<div class=""value1"">" & vbCrLf

	If dbProcess1 <> "" Then
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:60px; color:#666666; font-weight:bold;"">�X�e�b�v�P</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:300px;"">" & dbProcess1 & "</p>"
		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	End If

	If dbProcess2 <> "" Then
		sHTML = sHTML & "<p style=""width:60px; color:#666666; text-align:center;"">��</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:60px; color:#666666; font-weight:bold;"">�X�e�b�v�Q</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:300px;"">" & dbProcess2 & "</p>"
		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	End If

	If dbProcess3 <> "" Then
		sHTML = sHTML & "<p style=""width:60px; color:#666666; text-align:center;"">��</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:60px; color:#666666; font-weight:bold;"">�X�e�b�v�R</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:300px;"">" & dbProcess3 & "</p>"
		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	End If

	If dbProcess4 <> "" Then
		sHTML = sHTML & "<p style=""width:60px; color:#666666; text-align:center;"">��</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:60px; color:#666666; font-weight:bold;"">�X�e�b�v�S</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:300px;"">" & dbProcess4 & "</p>"
		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	End If

	sHTML = sHTML & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If dbWValueURL <> "" Then
		If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True

		sHTML = sHTML & "<div class=""category1""><h4>���Ѝ̗p<br>�y�[�W</h4></div>"
		sHTML = sHTML & "<div class=""value1""><a href=""" & dbWValueURL & """ target=""_blank""><img src=""/img/order/btn_wvalue.gif"" border=""0"" alt=""���Ѝ̗p�y�[�W""></a></div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	End If

	sHTML = sHTML & "<br>" & vbCrLf

	GetHTMLEditHowToEntry = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[�̒S���ҘA������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvEditFlag		�F�ҏW�t���O
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit0.asp
'���@���F2008/10/10 LIS K.Kokubo �쐬
'�@�@�@�F2009/04/02 LIS K.Kokubo ���[���ۋ��v�����̏ꍇ�͘A�����\��
'******************************************************************************
Function GetHTMLEditContact(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	'<�ϐ��錾>
	Dim dbOrderCode			'���R�[�h
	Dim sCompanyCode		'��ƃR�[�h
	Dim dbCompanyName		'��Ɩ���
	Dim dbCompanyNameF		'��Ɩ��̃J�i
	Dim dbCompanyKbn		'��Ƌ敪
	Dim dbCompanySpeciality	'��Ɠ���
	Dim dbCSectionName		'�d���̘A����S������
	Dim dbCPersonName		'�d���̘A����S���Җ�
	Dim dbCPersonNameF		'�d���̘A����S���҃J�i
	Dim dbCTel				'�d���̘A����d�b�ԍ�
	Dim dbCMail				'�d���̘A���惁�[���A�h���X
	Dim sPerson
	Dim sContact
	Dim sOrderType
	Dim dbPlanTypeName
	Dim flgLine				'�������t���O

	Dim sHTML
	'</�ϐ��錾>

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
	dbCompanyName = rRS.Collect("CompanyName")
	dbCompanyNameF = rRS.Collect("CompanyName_F")
	dbCompanyKbn = rRS.Collect("CompanyKbn")
	dbCompanySpeciality = rRS.Collect("CompanySpeciality")

	'Call SetOrderCompanyName(dbCompanyName, dbCompanyNameF, sOrderType, dbCompanyKbn, dbCompanySpeciality)
	'------------------------------------------------------------------------------
	'��Ж� end
	'******************************************************************************

	'******************************************************************************
	'�d���̘A���� start
	'------------------------------------------------------------------------------
	If sOrderType = "0" Then
		dbCSectionName = ChkStr(rRS.Collect("ContactSectionName"))
		dbCPersonName = ChkStr(rRS.Collect("ContactPersonName"))
		dbCPersonNameF = ChkStr(rRS.Collect("ContactPersonName_F"))
		dbCTel = ChkStr(rRS.Collect("ContactTelNumber"))
		dbCMail = ChkStr(rRS.Collect("ContactMailAddress"))

		If dbCompanyKbn = "2" Then
			'�l�މ�Ђ̋��l�[�̏ꍇ�́u���O�v�{�u�l�މ�Ж��v
			sPerson = dbCPersonName & "&nbsp;(�l�މ�ЁF" & dbCompanyName & ")"
		Else
			'��ʊ�Ƃ̋��l�[�̏ꍇ�́u���O�v�{�u�J�i�v
			sPerson = dbCPersonName
			If dbCPersonNameF <> "" Then sPerson = sPerson & "(" & dbCPersonNameF & ")"
		End If
	End If

	If dbCSectionName = "" Then dbCSectionName = "<span style=""color:#999999;"">[�A����̕�����]�������͂ł��B</span>"
	If sPerson = "" Then sPerson = "<span style=""color:#999999;"">[�A����̒S���Җ�]�������͂ł��B</span>"
	If dbCTel = "" Then
		dbCTel = "<span style=""color:#999999;"">[�A����̓d�b�ԍ�]�������͂ł��B</span>"
	Else
		dbCTel = dbCTel & "<span style=""font-size:10px;"">�@���d�b���ł̂��₢���킹�̍ہA�u�����ƃi�r�������v�ƌ����ƃX���[�Y�ł��B</span>"
	End If
	If dbCMail = "" Then dbCMail = "<span style=""color:#999999;"">[�A����̃��[���A�h���X]�������͂ł��B</span>"

	sContact = ""
	If dbCTel <> "" Then sContact = sContact & dbCTel
	If sContact <> "" Then sContact = sContact & "<br>"
	If dbCMail <> "" Then sContact = sContact & dbCMail
	'------------------------------------------------------------------------------
	'�d���̘A����
	'******************************************************************************

	flgLine = False
	sHTML = sHTML & "<a name=""edit20""></a>"
	sHTML = sHTML & "<h3 class=""sp"">�S���ҏ��</h3>"
	sHTML = sHTML & "<div style=""margin-left:15px;""><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit20.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><span style=""color:#ff0000; font-weight:normal; font-size:9px;"">���K�{</span></div>"

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
	sHTML = sHTML & "<div class=""category1""><h4>�S����</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sPerson & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
	sHTML = sHTML & "<div class=""category1""><h4>�S������</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbCSectionName & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	If dbPlanTypeName <> "mail" Then
		'���[���ۋ��v�����̏ꍇ�͘A�����\��
		If flgLine = True Then sHTML = sHTML & "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
		sHTML = sHTML & "<div class=""category1""><h4>�A����</h4></div>"

		sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sContact & "</p></div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
	End If

	sHTML = sHTML & "<br>"

	GetHTMLEditContact = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[�ڍׂ̐�y�C���^�r���[���o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvEditFlag		�F�ҏW�t���O
'���@�l�F
'�g�p���F�����ƃi�r/company/orderedit/edit0.asp
'�X�@�V�F2008/10/10 LIS K.Kokubo �쐬
'******************************************************************************
Function GetHTMLElderInterview(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
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

	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")

	sSQL = "EXEC up_LstC_ElderInterview '" & dbOrderCode & "', '1'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	sHTML = ""
	sHTML = sHTML & "<a name=""elderinterview""></a>"
	sHTML = sHTML & "<h3>��y�C���^�r���[</h3>"
	sHTML = sHTML & "<div style=""margin-left:15px;""><input class=""btn1"" type=""button"" value=""�ҁ@�W"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/elderinterview/list.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""></div>"

	If GetRSState(oRS) = True Then
		sHTML = sHTML & "<div class=""freeprblock"">"

		Do While GetRSState(oRS) = True
			dbSeq = oRS.Collect("Seq")
			dbProfile = oRS.Collect("Profile")
			dbQuestion = oRS.Collect("Question")
			dbAnswer = oRS.Collect("Answer")
			dbPublicFlag = oRS.Collect("PublicFlag")
			dbPictureFlag = oRS.Collect("PictureFlag")

			sHTML = sHTML & "<h4>" & dbProfile & "</h4>"
			sHTML = sHTML & "<div style=""clear:both;""></div>"

			If dbPictureFlag = "1" Then
				'��y�ʐ^�L��
				sHTML = sHTML & "<div style=""width:580px; margin-left:20px;"">"
				sHTML = sHTML & "<div style=""float:left; width:182px; padding-top:5px;"">"
				sHTML = sHTML & "<img src=""/company/elderinterview/picture.asp?ordercode=" & dbOrderCode & "&amp;seq=" & dbSeq & """ alt="""" border=""1"" width=""180"" height=""135"" style=""border:1px solid:#999999;"">"
				sHTML = sHTML & "</div>"
				sHTML = sHTML & "<div style=""float:left; width:398px;"">"
				sHTML = sHTML & "<p style=""margin:0px; padding-left:5px;"">��" & dbQuestion & "</p>"
				sHTML = sHTML & "<p style=""margin:0px; padding-left:5px;"">" & dbAnswer & "</p>"
				sHTML = sHTML & "</div>"
				sHTML = sHTML & "<div style=""clear:both;""></div>"
				sHTML = sHTML & "</div>"
			Else
				'��y�ʐ^����
				sHTML = sHTML & "<p class=""m0"">��" & dbQuestion & "</p>"
				sHTML = sHTML & "<p class=""m0"">" & dbAnswer & "</p>"
			End If
			oRS.MoveNext
		Loop

		sHTML = sHTML & "</div>"
	Else
		sHTML = sHTML & "<div style=""margin-left:15px;""><span style=""color:#999999;"">[��y�C���^�r���[]�������͂ł��B</span></div>"
	End If

	sHTML = sHTML & "<br>"

	GetHTMLElderInterview = sHTML
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
		sWorkingType = sWorkingType & oRS.Collect("WorkingTypeName")

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
'�X�@�V�F2006/05/08 LIS K.Kokubo �쐬
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

		If sStation <> "" Then sStation = sStation & ","
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
'�X�@�V�F2006/05/08 LIS K.Kokubo �쐬
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
'���@���F2008/10/08 LIS K.Kokubo �쐬
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
	Dim sOrderType
	Dim sCompanyKbn

	If GetRSState(rRS) = False Then Exit Function

	sOrderType = rRS.Collect("OrderType")
	sCompanyKbn = rRS.Collect("CompanyKbn")

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
	'2009/12/01 LIS K.Kokubo ADD 
	If rRS.Collect("AppointmentFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_appointmentflag.gif"" alt=""���Ј��o�p���x����"" width=""50"" height=""15"">&nbsp;"
	'2009/12/01 LIS K.Kokubo ADD 
	If rRS.Collect("SocietyInsuranceFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_societyinsuranceflag.gif"" alt=""�Еۊ���"" width=""50"" height=""15"">&nbsp;"
	'2008/05/08 LIS K.Kokubo ADD �V�[�N���b�g���l
	If rRS.Collect("SecretFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order/secret.gif"" alt=""�X�J�E�g���󂯂��l�������{���ł��鋁�l���"" width=""50"" height=""15"">&nbsp;"

	GetImgOrderSpeciality = sHTML
End Function

'******************************************************************************
'�T�@�v�F���l�[���͉�ʂ̌ٗp�`�ԕ���
'���@���FrDB		�F
'�@�@�@�FvUserType	�F
'�@�@�@�FvOrderCode	�F
'�߂�l�F
'�쐬�ҁFLis K.Kokubo
'�쐬���F2007/03/27
'���@�l�F
'�g�p���F�����ƃi�r/company/company_reg2.asp
'******************************************************************************
Function GetHTMLOrderInputWorkingType(ByRef rDB, ByVal vUserType, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sHTML
	Dim idx
	Dim idxMax
	Dim dbWorkingTypeCode
	Dim dbWorkingTypeName

	sHTML = ""

	'���X�g�{�b�N�X�o�͌��w��
	If vUserType = "company" Then
		idxMax = 3
	ElseIf vUserType = "dispatch" Then
		idxMax = 1
	End If

	sSQL = "sp_GetDataWorkingType '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	idx = 1
	Do While GetRSState(oRS) = True Or idx <= idxMax
		dbWorkingTypeCode = ""
		dbWorkingTypeName = ""

		If GetRSState(oRS) = True Then
			dbWorkingTypeCode = oRS.Collect("WorkingTypeCode")
			dbWorkingTypeName = oRS.Collect("WorkingTypeName")
		End If

		If vUserType = "company" Then
			'��ʊ�� or �Љ��� �̏ꍇ�́A�u�h���E�Љ�\��h���v���\��
			If sHTML <> "" Then sHTML = sHTML & "<div style=""float:left; width:3px;""></div>" & vbCrLf
			sHTML = sHTML & "<div style=""float:left; width:198px;"">"
			sHTML = sHTML & "<p class=""m0"" style=""text-align:center;"">�Ζ��`��" & idx & "</p>"
			sHTML = sHTML & "<select name=""frmworkingtypecode" & idx & """ size=""6"" style=""width:98%;"" onchange=""ChkDuplication('frmworkingtypecode', '��]�Ζ��`��');"">"
			sHTML = sHTML & "<option value="""">�I�����ĉ�����</option>"
			sHTML = sHTML & "<option value=""002"""
			If dbWorkingTypeCode = "002" Then sHTML = sHTML & " selected"
			sHTML = sHTML & ">���Ј�</option>"
			sHTML = sHTML & "<option value=""003"""
			If dbWorkingTypeCode = "003" Then sHTML = sHTML & " selected"
			sHTML = sHTML & ">�_��Ј�</option>"
			sHTML = sHTML & "<option value=""005"""
			If dbWorkingTypeCode = "005" Then sHTML = sHTML & " selected"
			sHTML = sHTML & ">�p�[�g�E�A���o�C�g</option>"
			sHTML = sHTML & "<option value=""006"""
			If dbWorkingTypeCode = "006" Then sHTML = sHTML & " selected"
			sHTML = sHTML & ">SOHO�i�ݑ�A���o�C�g�E���Ɓj</option>"
			sHTML = sHTML & "<option value=""007"""
			If dbWorkingTypeCode = "007" Then sHTML = sHTML & " selected"
			sHTML = sHTML & ">FC�E�㗝�X</option>"
			sHTML = sHTML & "</select>"
			sHTML = sHTML & "</div>" & vbCrLf
		ElseIf vUserType = "dispatch" Then
			'�h����� �̏ꍇ�́u�h���E�Љ�\��h���v�̂ݕ\��
			sHTML = sHTML & "<select name=""frmworkingtypecode" & idx & """ onchange=""ChkDuplication('frmworkingtypecode', '��]�Ζ��`��');"">"
			sHTML = sHTML & "<option value="">�I�����ĉ�����</option>"
			sHTML = sHTML & "<option value=""001"""
			If dbWorkingTypeCode = "001" Then Response.Write " selected"
			sHTML = sHTML & ">�h��</option>"
			sHTML = sHTML & "<option value=""004"""
			If dbWorkingTypeCode = "004" Then Response.Write " selected"
			sHTML = sHTML & ">�Љ�\��h��</option>"
			sHTML = sHTML & "</select>"
		End If

		If GetRSState(oRS) = True Then oRS.MoveNext
		idx = idx + 1
	Loop
	Call RSClose(oRS)

	If vUserType = "company" Then sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	GetHTMLOrderInputWorkingType = sHTML
End Function
%>
