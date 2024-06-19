<%
'**********************************************************************************************************************
'概　要：求人票一覧ページ /order/order_list_entity.asp
'　　　：求人票詳細ページ /order/order_detail_entity.asp
'　　　：企業情報ページ /order/company_order.asp
'　　　：上記ページで出力用の関数群をこのファイルに用意する。
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　求人票詳細ページ出力用　■■■
'　　　：ChkAgencyEditOrder				：代理店の求人票登録権限可否チェック
'　　　：ChkEditOrder					：求人票登録権限可否チェック
'　　　：DspLisOrderCompanyInfo			：求人票編集ページのリスの紹介先・派遣先企業情報HTMLを取得
'　　　：GetHTMLEditTempOrderCompanyInfo：求人票編集ページの派遣企業の派遣先企業情報HTMLを取得
'　　　：GetHTMLEditOrderCompanyName	：求人票編集ページの企業名HTMLを取得
'　　　：GetHTMLEditOrderShowTypeSwitch	：求人票編集ページの会社情報・職種情報・インタビュー切り替えボタンと参照回数HTMLを取得
'　　　：GetHTMLEditOrderCatchCopy		：求人票編集ページのキャッチコピー部分（大きい画像など）HTMLを取得
'　　　：GetHTMLEditOrderFreePR			：求人票編集ページのフリーＰＲHTMLを取得
'　　　：GetHTMLEditOrderPictureNow		：求人票編集ページの小さい画像HTMLを取得
'　　　：GetHTMLEditOrderBackGround		：求人票編集ページの採用の背景HTMLを取得
'　　　：GetHTMLEditBusiness			：求人票編集ページの業務内容HTMLを取得
'　　　：GetHTMLEditCondition			：求人票編集ページの勤務条件HTMLを取得
'　　　：GetHTMLEditNeedCondition		：求人票編集ページの必要条件HTMLを取得
'　　　：GetHTMLEditHowToEntry			：求人票編集ページの応募情報HTMLを取得
'　　　：GetHTMLEditContact				：求人票編集ページの担当者連絡先HTMLを取得
'　　　：GetHTMLElderInterview			：求人票編集ページの先輩インタビューHTMLを取得
'　　　：GetWorkingType					：求人票詳細ページの勤務形態部分
'　　　：GetJobType						：求人票詳細ページの職種部分
'　　　：GetWorkingTime					：求人票詳細ページの勤務形態部分
'　　　：GetNearbyStation				：求人票詳細ページの最寄駅部分
'　　　：GetNearbyRailway				：求人票詳細ページの最寄沿線部分
'　　　：GetSkill						：求人票詳細ページのスキル部分
'　　　：GetLicense						：求人票詳細ページの資格部分
'　　　：GetOrderNote					：求人票詳細ページの資格部分
'　　　：GetOrderTitle					：求人票詳細ページのタイトルとディスクリプションを取得
'　　　：GetSkillList					：求人票詳細ページのスキルの各項目表示
'　　　：GetHTMLOrderInputWorkingType	：求人票入力画面の雇用形態部分
'**********************************************************************************************************************

'******************************************************************************
'概　要：代理店の求人票登録権限可否チェック
'引　数：vAgencyCode		：ログイン中代理店コード
'　　　：vBranchSeq			：ログイン中代理店拠点番号
'　　　：vApplicationCode	：申し込みコード
'　　　：vOrderCode			：情報コード
'戻り値：Boolean	：[True]求人票登録可能 [False]求人票登録不可
'備　考：
'使用元：
'更　新：2010/03/30 LIS K.kokubo 作成
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

	'<ライセンス切れはマイメニューへリダイレクト>
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
	'</ライセンス切れはマイメニューへリダイレクト>

	'<ログイン中の企業の情報コードかどうかをチェック>
	sSQL = "EXEC sp_ChkCompanyOrder '" & dbUserCode & "', '" & vOrderCode & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbCheck = oRS.Collect("CheckFlag")
		dbLicenseFlag = oRS.Collect("LicenseFlag")
	End If
	Call RSClose(oRS)
	If vOrderCode = "" Then dbCheck = "1"
	If dbCheck = "0" And dbLicenseFlag = "0" Then Exit Function
	'</ログイン中の企業の情報コードかどうかをチェック>

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
'概　要：求人票登録権限可否チェック
'引　数：vOrderCode	：情報コード
'　　　：vUserID	：ログイン中ユーザコード
'　　　：vUseFlag	：ログイン中企業のライセンスの有効フラグ
'戻り値：Boolean	：[True]求人票登録可能 [False]求人票登録不可
'備　考：
'使用元：しごとナビ/company/order/edit01.asp
'更　新：2008/10/08 LIS K.kokubo 作成
'******************************************************************************
Function ChkEditOrder(ByVal vOrderCode, ByVal vUserID, ByVal vUseFlag)
	Dim sSQL
	Dim oRS
	Dim sError
	Dim flgQE

	Dim dbCheck
	Dim dbLicenseFlag

'	If vOrderCode = "" Then Exit Function

	'<ライセンス切れはマイメニューへリダイレクト>
	sSQL = "EXEC up_DtlNaviLicense_Now '" & vUserID & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("LicenseType1Flag") <> "1" Then Exit Function
	End If
	Call RSClose(oRS)
	'</ライセンス切れはマイメニューへリダイレクト>

	'<ログイン中の企業の情報コードかどうかをチェック>
	sSQL = "sp_ChkCompanyOrder '" & vUserID & "', '" & vOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbCheck = oRS.Collect("CheckFlag")
		dbLicenseFlag = oRS.Collect("LicenseFlag")
	End If
	Call RSClose(oRS)
	If vOrderCode = "" Then dbCheck = "1"
	If dbCheck = "0" And dbLicenseFlag = "0" Then Exit Function
	'</ログイン中の企業の情報コードかどうかをチェック>

	ChkEditOrder = True
End Function

'******************************************************************************
'概　要：求人票詳細ページのリスの紹介先・派遣先企業情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'作成者：Lis Kokubo
'作成日：2007/02/11
'備　考：
'使用元：しごとナビ/order/order_detail_entity.asp
'******************************************************************************
Function DspLisOrderCompanyInfo(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sCompanyCode		'企業コード
	Dim sOrderType			'受注区分
	Dim sListClass			'株式公開
	Dim sIndustryType		'業種
	Dim sCapitalAmount		'資本額		'**TOP 08/08/21 Lis林 ADD
	Dim sEmployeeNum		'社員数
	Dim sAccountingPeriod1	'決算期1
	Dim sSalesAmount1		'売上高1
	Dim sOrdinaryProfit1	'経常利益1
	Dim sAccountingPeriod2	'決算期2
	Dim sSalesAmount2		'売上高2
	Dim sOrdinaryProfit2	'経常利益2
	Dim sAccountingPeriod3	'決算期3
	Dim sSalesAmount3		'売上高3
	Dim sOrdinaryProfit3	'経常利益3
	Dim sImportantNotice	'特記事項
	Dim sflgAct							'**BTM 08/08/21 Lis林 ADD
	Dim sPR					'事業内容
	Dim sImgTitle			'タイトルイメージ
	Dim sIntrDisp			'派遣 or 紹介文言
	Dim flgDsp
	Dim flgLine				'線引きフラグ

	DspLisOrderCompanyInfo = False

	If GetRSState(rRS) = False Then Exit Function

	If GetRSState(rRS) = True Then
		'******************************************************************************
		'企業コード start
		'------------------------------------------------------------------------------
		sCompanyCode = rRS.Collect("CompanyCode")
		sOrderType = rRS.Collect("OrderType")
		If sOrderType = "2" Then
			sImgTitle = "/img/order/lisorderpr2.gif"
			sIntrDisp = "紹介先"
		Else
			sImgTitle = "/img/order/lisorderpr1.gif"
			sIntrDisp = "派遣先"
		End If
		'------------------------------------------------------------------------------
		'企業コード end
		'******************************************************************************

		'******************************************************************************
		'株式公開 start
		'------------------------------------------------------------------------------
		sListClass = ""
		sListClass = rRS.Collect("ListClass")
		'------------------------------------------------------------------------------
		'株式公開 end
		'******************************************************************************

		'******************************************************************************
		'業種 start
		'------------------------------------------------------------------------------
		sIndustryType = ""
		sIndustryType = ChkStr(rRS.Collect("IndustryTypeName"))
		'------------------------------------------------------------------------------
		'業種 end
		'******************************************************************************

		'******************************************************************************
		'会社紹介 start
		'------------------------------------------------------------------------------
		sPR = ""
		sPR = Replace(ChkStr(rRS.Collect("BusinessContents")), vbCrLf, "<br>")
		sPR = Replace(sPR, vbCr, "<br>")
		sPR = Replace(sPR, vbLf, "<br>")
		'------------------------------------------------------------------------------
		'会社紹介 end
		'******************************************************************************
		'**TOP 08/08/21 Lis林 ADD
		'******************************************************************************
		'資本額 start
		'------------------------------------------------------------------------------
		sCapitalAmount = ""
		sCapitalAmount = ChkStr(rRS.Collect("CapitalAmount"))
		if IsNumeric(sCapitalAmount) = True then
			sCapitalAmount = GetJapaneseYen(sCapitalAmount)
		elseif sCapitalAmount <> "" then
			if InStr(sCapitalAmount,"円") > 0 then		'"円"が入っていたらそのまま
			else
				sCapitalAmount = sCapitalAmount & "円"
			end if
		end if
		'------------------------------------------------------------------------------
		'資本額 end
		'******************************************************************************

		'******************************************************************************
		'社員数 start
		'------------------------------------------------------------------------------
		sEmployeeNum = ""
		sEmployeeNum = ChkStr(rRS.Collect("AllEmployeeNum"))
		If sEmployeeNum <> "" Then sEmployeeNum = sEmployeeNum & "人"
		'------------------------------------------------------------------------------
		'社員数 end
		'******************************************************************************
		
		'******************************************************************************
		'決算期・売上高・経常利益 start
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
		'if sSalesAmount1 <> "" and InStr(sSalesAmount1,"円") <= 0 then sSalesAmount1 = sSalesAmount1 & "円"
		sOrdinaryProfit1 = ChkStr(rRS.Collect("OrdinaryProfit1"))
		'if sOrdinaryProfit1 <> "" and InStr(sOrdinaryProfit1,"円") <= 0 then sOrdinaryProfit1 = sOrdinaryProfit1 & "円"
		sAccountingPeriod2 = ChkStr(rRS.Collect("AccountingPeriod2"))
		sSalesAmount2 = ChkStr(rRS.Collect("SalesAmount2"))
		'if sSalesAmount2 <> "" and InStr(sSalesAmount2,"円") <= 0 then sSalesAmount2 = sSalesAmount2 & "円"
		sOrdinaryProfit2 = ChkStr(rRS.Collect("OrdinaryProfit2"))
		'if sOrdinaryProfit2 <> "" and InStr(sOrdinaryProfit2,"円") <= 0 then sOrdinaryProfit2 = sOrdinaryProfit2 & "円"
		sAccountingPeriod3 = ChkStr(rRS.Collect("AccountingPeriod3"))
		sSalesAmount3 = ChkStr(rRS.Collect("SalesAmount3"))
		'if sSalesAmount3 <> "" and InStr(sSalesAmount3,"円") <= 0 then sSalesAmount3 = sSalesAmount3 & "円"
		sOrdinaryProfit3 = ChkStr(rRS.Collect("OrdinaryProfit3"))
		'if sOrdinaryProfit3 <> "" and InStr(sOrdinaryProfit3,"円") <= 0 then sOrdinaryProfit3 = sOrdinaryProfit3 & "円"
		sImportantNotice = ChkStr(rRS.Collect("ImportantNotice"))
		'------------------------------------------------------------------------------
		'決算期・売上高・経常利益 end
		'******************************************************************************
		'**BTM 08/08/21 Lis林 ADD
	End If

	flgLine = False

	'**TOP 08/08/21 Lis林 REP
	'If sListClass & sIndustryType & sPR <> "" Then
	If sListClass & sIndustryType & sPR & sCapitalAmount & sEmployeeNum <> "" or _
		(InStr(sImportantNotice,"非公開") <= 0 and _
		((sAccountingPeriod1 <> "" and sSalesAmount1 <> "" and InStr(sAccountingPeriod1 & sSalesAmount1,"非公開") <= 0) or _
		 (sAccountingPeriod2 <> "" and sSalesAmount2 <> "" and InStr(sAccountingPeriod2 & sSalesAmount2,"非公開") <= 0) or _
		 (sAccountingPeriod3 <> "" and sSalesAmount3 <> "" and InStr(sAccountingPeriod3 & sSalesAmount3,"非公開") <= 0))) Then
	'**BTM 08/08/21 Lis林 REP
		DspLisOrderCompanyInfo = True
%>
<h3><%= sIntrDisp %>企業情報</h3>
<%
		If sListClass <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>株式公開</h4></div>
<div class="value1"><p class="m0"><%= sListClass %></p></div>
<div style="clear:both;"></div>
<%
		End If

		If sIndustryType <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>業種</h4></div>
<div class="value1"><p class="m0"><%= sIndustryType %></p></div>
<div style="clear:both;"></div>
<%
		End If


		If sPR <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
			

%>
<div class="category1"><h4>事業内容</h4></div>
<div class="value1"><p class="m0"><%= sPR %></p></div>
<div style="clear:both;"></div>
<%		End If
		'**TOP 08/08/21 Lis林 ADD
		If sCapitalAmount <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>資本金</h4></div>
<div class="value1"><p class="m0"><%= sCapitalAmount %></p></div>
<div style="clear:both;"></div>
<%		End If
		If sEmployeeNum <> "" Then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>社員数</h4></div>
<div class="value1"><p class="m0"><%= sEmployeeNum %></p></div>
<div style="clear:both;"></div>
<%		End If
		sflgAct = ""
		If InStr(sImportantNotice,"非公開") <= 0 and _
		((sAccountingPeriod1 <> "" and sSalesAmount1 <> "" and InStr(sAccountingPeriod1 & sSalesAmount1,"非公開") <= 0) or _
		 (sAccountingPeriod2 <> "" and sSalesAmount2 <> "" and InStr(sAccountingPeriod2 & sSalesAmount2,"非公開") <= 0) or _
		 (sAccountingPeriod3 <> "" and sSalesAmount3 <> "" and InStr(sAccountingPeriod3 & sSalesAmount3,"非公開") <= 0)) then
			If flgLine = True Then Response.Write "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
			flgLine = True
%>
<div class="category1"><h4>売上実績</h4></div>
<div class="value1"><p class="m0">
<%			'売上高１・経常利益１・決算期１
			if sAccountingPeriod1 <> "" and sSalesAmount1 <> "" and InStr(sAccountingPeriod1 & sSalesAmount1,"非公開") <= 0 then
				if sSalesAmount1 <> "" and InStr(sSalesAmount1,"非公開") <= 0 then
					response.write "売上高：" & sSalesAmount1 & "　"
				end if
				if sOrdinaryProfit1 <> "" and InStr(sOrdinaryProfit1,"非公開") <= 0 then
					response.write "経常利益：" & sOrdinaryProfit1
				end if
				if sAccountingPeriod1 <> "" and InStr(sAccountingPeriod1,"非公開") <= 0 then
					response.write "（決算期：" & sAccountingPeriod1 & "）<br>"
				end if
				sflgAct = "1"
			end if
			'売上高２・経常利益２・決算期２
			if sAccountingPeriod2 <> "" and sSalesAmount2 <> "" and InStr(sAccountingPeriod2 & sSalesAmount2,"非公開") <= 0 then
				if sSalesAmount2 <> "" and InStr(sSalesAmount2,"非公開") <= 0 then
					response.write "売上高：" & sSalesAmount2 & "　"
				end if
				if sOrdinaryProfit2 <> "" and InStr(sOrdinaryProfit2,"非公開") <= 0 then
					response.write "経常利益：" & sOrdinaryProfit2
				end if
				if sAccountingPeriod2 <> "" and InStr(sAccountingPeriod2,"非公開") <= 0 then
					response.write "（決算期：" & sAccountingPeriod2 & "）<br>"
				end if
				sflgAct = "1"
			end if
			'売上高３・経常利益３・決算期３
			if sAccountingPeriod3 <> "" and sSalesAmount3 <> "" and InStr(sAccountingPeriod3 & sSalesAmount3,"非公開") <= 0 then
				if sSalesAmount3 <> "" and InStr(sSalesAmount3,"非公開") <= 0 then
					response.write "売上高：" & sSalesAmount3 & "　"
				end if
				if sOrdinaryProfit3 <> "" and InStr(sOrdinaryProfit3,"非公開") <= 0 then
					response.write "経常利益：" & sOrdinaryProfit3
				end if
				if sAccountingPeriod3 <> "" and InStr(sAccountingPeriod3,"非公開") <= 0 then
					response.write "（決算期：" & sAccountingPeriod3 & "）<br>"
				end if
				sflgAct = "1"
			end if
			'特記事項
			If sflgAct = "1" and sImportantNotice <> "" and InStr(sImportantNotice,"非公開") <= 0 then
				response.write "（"
				if InStr(sImportantNotice,"※") <= 0 then response.write "※"
				response.write  sImportantNotice & "）<br>"
			End If
%>
</p></div>
<div style="clear:both;"></div>
<%		End If
%><p class="m0" style="font-size:10px;margin:0px 20px;color:red;">
※人材<%= left(sIntrDisp,2) %>でご案内するお仕事のため、詳しい会社情報は下のボタンやお電話などで直接お問合せください。</p>
<%		response.write "<p>　</p>"
		'**BTM 08/08/21 Lis林 ADD
	End If
End Function

'******************************************************************************
'概　要：求人票の応募情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'備　考：
'使用元：しごとナビ/company/orderedit/edit22.asp
'更　新：2009/03/17 LIS K.Kokubo 作成
'******************************************************************************
Function GetHTMLEditTempOrderCompanyInfo(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	Dim dbOrderCode			'情報コード
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
	Dim flgLine				'線引きフラグ
	Dim sCapital
	Dim sTempAllEmployeeNumber

	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function

	'<派遣先企業情報取得>
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
	'</派遣先企業情報取得>

	'<設立年度>
	If dbTempEstablishYear <> "" Then
		dbTempEstablishYear = dbTempEstablishYear & "年"
	Else
		dbTempEstablishYear = "<span style=""color:#999999;"">[設立年度]が未入力です。</span>"
	End If
	'</設立年度>

	'<業種>
	If dbTempIndustryTypeName <> "" Then
	Else
		dbTempIndustryTypeName = "<span style=""color:#999999;"">[業種]が未入力です。</span>"
	End If
	'</業種>

	'<資本>
	sCapital = ""
	If dbTempCapitalAmount & dbTempForeinCapital <> "" Then
		If dbTempCapitalAmount <> "" Then
			sCapital = sCapital & GetJapaneseYen(dbTempCapitalAmount)
		Else
			sCapital = sCapital & "<span style=""color:#999999;"">[資本金]が未入力です。</span>"
		End If

		If dbTempForeinCapital <> "" Then
			sCapital = sCapital & "&nbsp;（外資：" & dbTempForeinCapital & "）"
		Else
			sCapital = sCapital & "<br><span style=""color:#999999;"">[外資]が未入力です。</span><br>"
		End If
	End If
	'</資本>

	'<株式>
	If dbTempListClass <> "" Then
	Else
		dbTempListClass = "<span style=""color:#999999;"">[株式]が未入力です。</span>"
	End If
	'</株式>

	'<社員数>
	If dbTempAllEmployeeNumber <> "" Then
		sTempAllEmployeeNumber = dbTempAllEmployeeNumber & "人"
	Else
		dbTempAllEmployeeNumber = "<span style=""color:#999999;"">[社員数]が未入力です。</span>"
	End If
	'</社員数>

	flgLine = False

	sHTML = sHTML & "<a name=""edit22""></a>"
	sHTML = sHTML & "<h3>派遣先企業情報</h3>" & vbCrLf
	sHTML = sHTML & "<div style=""margin-left:15px;""><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit22.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""></div>"

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>設立年度</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbTempEstablishYear & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>業種</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbTempIndustryTypeName & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>資本金</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"
	sHTML = sHTML & "<p class=""m0"">" & sCapital & "</p>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>株式</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbTempListClass & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>社員数</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sTempAllEmployeeNumber & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "<br>" & vbCrLf

	GetHTMLEditTempOrderCompanyInfo = sHTML
End Function

'******************************************************************************
'概　要：求人票編集ページの企業名称を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'備　考：
'使用元：しごとナビ/company/orderedit/edit0.asp
'更　新：2008/10/10 LIS K.Kokubo 作成
'******************************************************************************
Function GetHTMLEditOrderCompanyName(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sHTML
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbSecretFlag		'シークレットフラグ
	Dim dbCompanyName		'企業名称
	Dim dbCompanyNameF		'企業名称カナ
	Dim dbCompanyKbn		'企業区分
	Dim dbCompanySpeciality	'企業特徴

	Dim sPublishLimitStr	'掲載期限表示用文字列
	Dim sCautionStr			'掲載期限表示注意文言文字列
	Dim flgNowPublic		'現在掲載中の求人票判定 '[True]掲載中 [False]非掲載

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

	'リス紹介案件の場合は「転職相談案件」イメージを表示
	If dbOrderType = "2" Then sHTML = sHTML & "<img src=""/img/order/counselable_order.gif"" width=""150"" height=""25"" alt=""転職支援を受けて応募する求人です"">"
	'シークレット求人の場合は「シークレット求人」イメージを表示
	'If dbSecretFlag = "1" Then sHTML = sHTML & "<img src=""/img/order/secret_order.gif"" width=""150"" height=""25"" alt=""この求人からスカウトを受けた人だけが閲覧できる求人です"">"
	If dbSecretFlag = "1" Then sHTML = sHTML & "<p class=""m0"" style=""color:#ff9933; font-weight:bold;"">■スカウトを受けた人だけが閲覧できる求人情報です。</p>"

	sHTML = sHTML & "<div style=""width:400px; margin-bottom:10px;"">"
	If G_COMPANYKBN = "2" Then
		sHTML = sHTML & "<a name=""edit21""></a>"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit21.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
		If dbCompanySpeciality = "" Then
			sHTML = sHTML & "&nbsp;<span style=""color:#999999;"">※[表示用会社名]が未入力です。</span>"
		End If
	End If
	sHTML = sHTML & "<div style=""font-size:14px; font-weight:bold;"">" & dbCompanyName & "</div>"
	sHTML = sHTML & "<div style=""font-size:10px; color:#666666;"">" & dbCompanyNameF & "</div>"
	sHTML = sHTML & "</div>"

	GetHTMLEditOrderCompanyName = sHTML
End Function

'******************************************************************************
'概　要：求人票詳細ページの会社情報・職種情報・インタビュー切り替えボタンと参照回数を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vType			：表示中情報の種類 ["0"]職種情報 ["1"]会社情報 ["2"]インタビュー
'　　　：vAccessCount	：表示中求人票のアクセス回数
'備　考：
'使用元：しごとナビ/company/orderedit/edit0.asp
'更　新：2008/10/10 LIS K.Kokubo 作成
'******************************************************************************
Function GetHTMLEditOrderShowTypeSwitch(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vType, ByVal vAccessCount)
	'<変数宣言>
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode			'情報コード
	Dim dbOrderType			'受注種類
	Dim dbJobTypeDetail		'具体的職種名
	Dim dbTopInterviewFlag	'トップインタビュー有無フラグ
	Dim sUpdateDay

	Dim sHTML
	'</変数宣言>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbJobTypeDetail = rRS.Collect("JobTypeDetail")
	dbTopInterviewFlag = rRS.Collect("TopInterviewFlag")
	'更新日
	sUpdateDay = GetDateStr(rRS.Collect("UpdateDay"), "/")

	If dbJobTypeDetail <> "" Then dbJobTypeDetail = dbJobTypeDetail & "のお仕事情報詳細"

	sHTML = sHTML & "<div style=""width:600px; margin-bottom:5px;"">"
	sHTML = sHTML & "<div style=""float:left; width:350px; margin:0px;"">"
	If vType = "0" Then
		'仕事情報を表示中の場合
		sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_orderdetail_on.gif"" alt=""" & dbJobTypeDetail & """ border=""0"" width=""93"" height=""22""></div>"
		If dbOrderType = "0" Then
			'一般の求人広告の場合は会社情報へのリンクを表示
			sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><a href=""./company_order.asp?poc=" & dbOrderCode & """ title=""会社情報""><img src=""/img/order/tab_companyinfo_off.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></a></div>"
		End If

		If dbOrderType = "0" And dbTopInterviewFlag = "1" Then
			sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & dbOrderCode & """ title=""会社情報""><img src=""/img/order/tab_interview_off.gif"" alt=""インタビュー"" border=""0"" width=""93"" height=""22""></a></div>"
		End If
	ElseIf vType = "1" Then
		'会社情報を表示中の場合
		sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & dbOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & dbJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
		If dbOrderType = "0" Then
			'一般の求人広告の場合は会社情報を表示
			sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_companyinfo_on.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></div>"
		End If

		If dbOrderType = "0" And dbTopInterviewFlag = "1" Then
			sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & dbOrderCode & """ title=""会社情報""><img src=""/img/order/tab_interview_off.gif"" alt=""インタビュー"" border=""0"" width=""93"" height=""22""></a></div>"
		End If

	ElseIf vType = "2" Then
		'インタビューを表示中の場合
		sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & dbOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & dbJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
		sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><a href=""./company_order.asp?poc=" & dbOrderCode & """ title=""会社情報""><img src=""/img/order/tab_companyinfo_off.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></a></div>"
		sHTML = sHTML & "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_interview_on.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></div>"
	End If
	sHTML = sHTML & "<div class=""clear:both; margin:0px;""></div>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div align=""right"" style=""float:right; width:250px;"">"
	sHTML = sHTML & "<p class=""m0"">月間参照回数：" & vAccessCount & "回　更新日：" & sUpdateDay & "</p>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""><img src=""/img/order/tab_border.gif"" alt="""" width=""600"" height=""5""></div>"
	sHTML = sHTML & "</div>" & vbCrLf

	GetHTMLEditOrderShowTypeSwitch = sHTML
End Function

'******************************************************************************
'概　要：求人票のキャッチコピー部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'備　考：
'使用元：しごとナビ/company/orderedit/edit0.asp
'更　新：2008/10/10 LIS K.Kokubo 作成
'******************************************************************************
Function GetHTMLEditOrderCatchCopy(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	'<変数宣言>
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode
	Dim dbOptionNo		'大きい写真の番号
	Dim dbJobTypeDetail	'具体的職種名
	Dim dbCatchCopy		'キャッチコピー

	Dim sHTML
	Dim sImg1
	Dim sClass
	Dim sImgOrderSpeciality
	'</変数宣言>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")
	dbJobTypeDetail = ChkStr(rRS.Collect("JobTypeDetail"))
	dbCatchCopy = ChkStr(rRS.Collect("CatchCopy"))
	sImgOrderSpeciality = GetImgOrderSpeciality(rDB, rRS)

	If dbJobTypeDetail = "" Then dbJobTypeDetail = "<span style=""color:#999999;"">[具体的職種名]が未入力です。</span>"
	If dbCatchCopy = "" Then dbCatchCopy = "<span style=""color:#999999;"">[キャッチコピー]が未入力です。</span>"
	If sImgOrderSpeciality = "" Then sImgOrderSpeciality = "<span style=""color:#999999;"">[募集の特徴]が未入力です。</span>"

	'******************************************************************************
	'大きい画像 start
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
	'大きい画像 end
	'******************************************************************************

	If dbImageLimit > 0 Then
		sHTML = sHTML & "<div id=""catchcopy"">" & vbCrLf

		'<右側>
		sHTML = sHTML & "<div class=""left"">"
		'
		If dbImageLimit = 1 Then
			sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" style=""margin-left:1px;"" onclick=""location.href='" & HTTP_CURRENTURL & "company/img_upload.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
		Else
			sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" style=""margin-left:1px;"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=001&amp;appcode=" & vApplicationCode & "';"">"
		End If
		sHTML = sHTML & "<br>"

		If sImg1 <> "" Then
			sHTML = sHTML & "<div class=""main_pics""><img class=""big"" src=""" & sImg1 & """ alt="""" border=""1"" id=""big_pics""></div>"
		Else
			sHTML = sHTML & "<img class=""big"" src=""" & sImg1 & """ alt=""[写真]が未登録です。"" border=""1"" width=""300"" height=""225"" style=""border:1px solid #999999;"">"
		End If
		sHTML = sHTML & "</div>"
		'</右側>

		'<左側>
		sHTML = sHTML & "<div style=""float:right; width:298px;"">"
		sHTML = sHTML & "<a name=""edit00""></a>"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit08.asp?ordercode=" & dbOrderCode & "&amp;place=1&amp;appcode=" & vApplicationCode & "';""><span style=""color:#ff0000; font-weight:normal; font-size:9px;"">※必須</span><br>"
		sHTML = sHTML & "<h2 style=""margin-bottom:15px;"">" & dbJobTypeDetail & "</h2><br>"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit01.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
		sHTML = sHTML & "<a name=""edit01""></a>"
		sHTML = sHTML & "<p class=""m0"" style=""margin-bottom:15px;"">" & dbCatchCopy & "</p><br>"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit02.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
		sHTML = sHTML & "<a name=""edit02""></a>"
		sHTML = sHTML & "<div style=""margin:10px 0px;"">"
		If sImgOrderSpeciality <> "" Then
			sHTML = sHTML & "<div style=""border:solid 0px #cccccc;padding:5px;"">"
			sHTML = sHTML & "<div style=""font-size:12px;font-weight:normal;color:#008900;"">【募集の特徴】</div>"
			sHTML = sHTML & sImgOrderSpeciality
			sHTML = sHTML & "</div>"
		End If
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "</div>"
		'<左側>

		'float解除
		sHTML = sHTML & "<br clear=""all"">"

		sHTML = sHTML & "</div>"
	Else
		sHTML = sHTML & "<div id=""catchcopy"" style=""width:600px;"">"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit08.asp?ordercode=" & dbOrderCode & "&amp;place=1&amp;appcode=" & vApplicationCode & "';""><span style=""color:#ff0000; font-weight:normal; font-size:9px;"">※必須</span><br>"
		sHTML = sHTML & "<h2 style=""width:600px;"">" & dbJobTypeDetail & "</h2><br>"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit01.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>"
		sHTML = sHTML & "<a name=""edit01""></a>"
		sHTML = sHTML & "<p class=""m0"">" & dbCatchCopy & "</p><div style=""margin-top:10px;clear:both;""></div>"
		sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit02.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
		sHTML = sHTML & "<a name=""edit02""></a>"
		sHTML = sHTML & "<div style=""margin-bottom:10px;"">"
		If sImgOrderSpeciality <> "" Then
			sHTML = sHTML & "<div style=""border:solid 0px #cccccc;padding:5px;"">"
			sHTML = sHTML & "<div style=""font-size:12px;font-weight:normal;color:#008900;"">【募集の特徴】</div>"
			sHTML = sHTML & sImgOrderSpeciality
			sHTML = sHTML & "</div>"
		End If
		sHTML = sHTML & "</div>"
		sHTML = sHTML & "</div>"
	End If

	GetHTMLEditOrderCatchCopy = sHTML
End Function

'******************************************************************************
'概　要：求人企業画像一覧表示ＨＴＭＬ表示
'引　数：rDB			：接続中ＤＢオブジェクト
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vCategoryCode	：カテゴリコード
'　　　：vEditFlag		：編集フラグ
'備　考：
'使用元：しごとナビ/company/orderedit/edit0.asp
'更　新：2008/10/10 LIS K.Kokubo 作成
'******************************************************************************
Function GetHTMLEditOrderPictureNow(ByRef rDB, ByRef rRS, ByVal vCategoryCode, ByVal vApplicationCode)
	'<変数宣言>
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
	'</変数宣言>

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
				sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=002&amp;appcode=" & vApplicationCode & "';"">"
				If dbOptionNo2 > 0 Then
					sHTML = sHTML & "&nbsp;<input class=""btn1"" type=""button"" value=""削　除"" onclick=""if(confirm('写真を外しますか？') === true)location.href='" & HTTP_CURRENTURL & "company/order_img_relationdelete.asp?ordercode=" & dbOrderCode & "&amp;code=002&amp;appcode=" & vApplicationCode & "';"">"
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
				sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=003&amp;appcode=" & vApplicationCode & "';"">"
				If dbOptionNo3 > 0 Then
					sHTML = sHTML & "&nbsp;<input class=""btn1"" type=""button"" value=""削　除"" onclick=""if(confirm('写真を外しますか？') === true)location.href='" & HTTP_CURRENTURL & "company/order_img_relationdelete.asp?ordercode=" & dbOrderCode & "&amp;code=003';"">"
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
				sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=004&amp;appcode=" & vApplicationCode & "';"">"
				If dbOptionNo4 > 0 Then
					sHTML = sHTML & "&nbsp;<input class=""btn1"" type=""button"" value=""削　除"" onclick=""if(confirm('写真を外しますか？') === true)location.href='" & HTTP_CURRENTURL & "company/order_img_relationdelete.asp?ordercode=" & dbOrderCode & "&amp;code=004';"">"
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
		sHTML = sHTML & "<div align=""center"" style=""margin-bottom:3px;""><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=002&amp;appcode=" & vApplicationCode & "';""></div>"
		sHTML = sHTML & "<div style=""width:182px; background-color:#ffffff;""><table border=""1"" style=""width:180px; height:135px; border:1px solid #999999;""><tr><td></td></tr></table></div>"
		sHTML = sHTML & "<p class=""m0"" align=""left"" style=""width:182px; font-size:10px;""></p>"
		sHTML = sHTML & "</div>"

		sHTML = sHTML & "<div align=""right"" style=""float:left; width:190px;"">"
		sHTML = sHTML & "<div align=""center"" style=""margin-bottom:3px;""><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=003&amp;appcode=" & vApplicationCode & "';""></div>"
		sHTML = sHTML & "<div style=""width:182px; background-color:#ffffff;""><table border=""1"" style=""width:180px; height:135px; border:1px solid #999999;""><tr><td></td></tr></table></div>"
		sHTML = sHTML & "<p class=""m0"" align=""left"" style=""width:182px; font-size:10px;""></p>"
		sHTML = sHTML & "</div>"

		sHTML = sHTML & "<div align=""right"" style=""float:left; width:190px;"">"
		sHTML = sHTML & "<div align=""center"" style=""margin-bottom:3px;""><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTP_CURRENTURL & "company/order_img_list.asp?ordercode=" & dbOrderCode & "&amp;code=004&amp;appcode=" & vApplicationCode & "';""></div>"
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
'概　要：求人票詳細ページのフリーＰＲを出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vEditFlag		：編集フラグ
'備　考：
'使用元：しごとナビ/company/orderedit/edit0.asp
'更　新：2008/10/10 LIS K.Kokubo 作成
'******************************************************************************
Function GetHTMLEditOrderFreePR(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	'<変数宣言>
	Dim dbOrderCode		'情報コード
	Dim dbPRTitle1		'ＰＲタイトル1
	Dim dbPRTitle2		'ＰＲタイトル2
	Dim dbPRTitle3		'ＰＲタイトル3
	Dim dbPRContents1	'ＰＲ文1
	Dim dbPRContents2	'ＰＲ文2
	Dim dbPRContents3	'ＰＲ文3

	Dim sHTML
	'</変数宣言>

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

	If dbPRTitle1 = "" Then dbPRTitle1 = "<span style=""color:#999999;"">[ＰＲ１タイトル]が未入力です。</span>"
	If dbPRTitle2 = "" Then dbPRTitle2 = "<span style=""color:#999999;"">[ＰＲ２タイトル]が未入力です。</span>"
	If dbPRTitle3 = "" Then dbPRTitle3 = "<span style=""color:#999999;"">[ＰＲ３タイトル]が未入力です。</span>"
	If dbPRContents1 = "" Then dbPRContents1 = "<span style=""color:#999999;"">[ＰＲ１内容]が未入力です。</span>"
	If dbPRContents2 = "" Then dbPRContents2 = "<span style=""color:#999999;"">[ＰＲ２内容]が未入力です。</span>"
	If dbPRContents3 = "" Then dbPRContents3 = "<span style=""color:#999999;"">[ＰＲ３内容]が未入力です。</span>"

	sHTML = sHTML & "<a name=""edit03""></a>"
	sHTML = sHTML & "<h3>ＰＲ</h3>"
	sHTML = sHTML & "<div class=""freeprblock"">"

	sHTML = sHTML & "<div style=""margin-left:15px;""><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit03.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""></div>"

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
'概　要：求人票の採用の背景を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vEditFlag		：編集フラグ
'備　考：
'使用元：しごとナビ/company/orderedit/edit0.asp
'更　新：2008/10/10 LIS K.Kokubo 作成
'******************************************************************************
Function GetHTMLEditOrderBackGround(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	'<変数宣言>
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode
	Dim dbOrderBackGround	'採用の背景

	Dim sHTML
	'</変数宣言>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	'採用の背景取得
	dbOrderBackGround = Replace(ChkStr(rRS.Collect("OrderBackGround")), vbCrLf, "<br>")

	If dbOrderBackGround = "" Then dbOrderBackGround = "<span style=""color:#999999;"">[採用の背景]が未入力です。</span>"


	sHTML = sHTML & "<a name=""edit04""></a>"
	sHTML = sHTML & "<h3>採用の背景</h3>" & vbCrLf
	sHTML = sHTML & "<div style=""margin-left:15px;""><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit04.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""></div>"

	'採用の背景出力
	sHTML = sHTML & "<p class=""m0"" style=""padding-left:15px;"">" & dbOrderBackGround & "</p>" & vbCrLf

	sHTML = sHTML & "<br>"

	GetHTMLEditOrderBackGround = sHTML
End Function

'******************************************************************************
'概　要：求人票の業務内容を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vEditFlag		：編集フラグ
'備　考：
'使用元：しごとナビ/company/orderedit/edit0.asp
'更　新：2008/10/10 LIS K.Kokubo 作成
'******************************************************************************
Function GetHTMLEditBusiness(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode		'情報コード
	Dim dbPlanType		'求人票ライセンスプラン種類
	Dim dbBizName1		'仕事割合文言1
	Dim dbBizName2		'仕事割合文言2
	Dim dbBizName3		'仕事割合文言3
	Dim dbBizName4		'仕事割合文言4
	Dim dbBizPercentage1'仕事割合1
	Dim dbBizPercentage2'仕事割合2
	Dim dbBizPercentage3'仕事割合3
	Dim dbBizPercentage4'仕事割合4
	Dim dbBusinessDetail'担当業務
	Dim sClearSolid
	Dim flgLine			'線引きフラグ

	Dim sHTML
	Dim sBiz			'仕事割合HTML

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	dbOrderCode = rRS.Collect("OrderCode")
	dbPlanType = rRS.Collect("PlanTypeName")
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'******************************************************************************
	'仕事の割合 start
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

	If dbBizName1 = "" Then dbBizName1 = "<span style=""color:#999999;"">[仕事の割合１]が未入力です。</span>"
	If dbBizName2 = "" Then dbBizName2 = "<span style=""color:#999999;"">[仕事の割合２]が未入力です。</span>"
	If dbBizName3 = "" Then dbBizName3 = "<span style=""color:#999999;"">[仕事の割合３]が未入力です。</span>"
	If dbBizName4 = "" Then dbBizName4 = "<span style=""color:#999999;"">[仕事の割合４]が未入力です。</span>"

	If dbBizName1 & dbBizName2 & dbBizName3 & dbBizName4 <> "" Then
		If dbBizName1 <> "" And dbBizPercentage1 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & dbBizName1 & "</td><td class=""biz2"">" & dbBizPercentage1 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(dbBizPercentage1) * 3 & """ height=""20""></td></tr>"
		If dbBizName2 <> "" And dbBizPercentage2 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & dbBizName2 & "</td><td class=""biz2"">" & dbBizPercentage2 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(dbBizPercentage2) * 3 & """ height=""20""></td></tr>"
		If dbBizName3 <> "" And dbBizPercentage3 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & dbBizName3 & "</td><td class=""biz2"">" & dbBizPercentage3 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(dbBizPercentage3) * 3 & """ height=""20""></td></tr>"
		If dbBizName4 <> "" And dbBizPercentage4 <> "0" Then sBiz = sBiz & "<tr><td class=""biz1"">" & dbBizName4 & "</td><td class=""biz2"">" & dbBizPercentage4 & "%</td><td class=""biz3"" valign=""middle""><img src=""/img/bar.gif"" alt="""" width=""" & CInt(dbBizPercentage4) * 3 & """ height=""20""></td></tr>"
		sBiz = "<table>" & sBiz & "</table>"
	End If
	'------------------------------------------------------------------------------
	'仕事の割合 end
	'******************************************************************************

	'******************************************************************************
	'担当業務 start
	'------------------------------------------------------------------------------
	dbBusinessDetail = Replace(ChkStr(rRS.Collect("BusinessDetail")), vbCrLf, "<br>")
	dbBusinessDetail = Replace(dbBusinessDetail, vbCr, "<br>")
	dbBusinessDetail = Replace(dbBusinessDetail, vbLf, "<br>")
	If dbBusinessDetail = "" Then dbBusinessDetail = "<span style=""color:#999999;"">[担当業務]が未入力です。</span>"
	'------------------------------------------------------------------------------
	'担当業務 end
	'******************************************************************************

	sHTML = sHTML & "<h3>業務内容</h3>"

	flgLine = False

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<a name=""edit05""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit05.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>担当業務</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbBusinessDetail & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	If (dbPlanType = "platinum" Or dbPlanType = "gold" Or dbPlanType = "old") Then
		If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
		sHTML = sHTML & "<a name=""edit06""></a>"
		sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit06.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>仕事の割合</h4></div>"
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
'概　要：求人票の勤務条件を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'備　考：
'使用元：しごとナビ/company/orderedit/edit0.asp
'履　歴：2008/10/10 LIS K.Kokubo 作成
'　　　：2009/04/16 LIS K.Kokubo メール課金ライセンスの場合は勤務地の表示を一般の求人広告でも市区郡までしか表示させない
'　　　：2009/04/22 LIS K.Kokubo 紹介後の勤務形態(TTP用)対応
'　　　：2009/11/02 LIS K.Kokubo ＳＯＨＯ,ＦＣの勤務地表示対応
'******************************************************************************
Function GetHTMLEditCondition(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	'<変数宣言>
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim oRS3
	Dim flgQE
	Dim sError

	Dim dbOrderCode			'情報コード
	Dim dbOrderType			'求人票種類
	Dim dbCompanyKbn		'企業区分
	Dim dbJobTypeDetail		'職種詳細
	Dim dbYearlyIncomeMin	'年収下限
	Dim dbYearlyIncomeMax	'年収上限
	Dim dbMonthlyIncomeMin	'月給下限
	Dim dbMonthlyIncomeMax	'月給上限
	Dim dbDailyIncomeMin	'日給下限
	Dim dbDailyIncomeMax	'日給上限
	Dim dbHourlyIncomeMin	'時給下限
	Dim dbHourlyIncomeMax	'時給上限
	Dim dbPercentagePay		'歩合制
	Dim dbSalaryRemark		'給与備考
	Dim dbTrafficFeeType	'
	Dim dbTrafficFeeMonth	'交通費／１ヶ月
	Dim dbAfterWorkingTypeCode'紹介後の勤務形態
	Dim dbWorkStartDay		'就業開始日
	Dim dbWorkEndDay		'就業終了日
	Dim dbWorkTimeRemark	'就業時間備考
	Dim dbWeeklyHolidayType	'週休
	Dim dbHolidayRemark		'休日備考
	Dim dbHumanNumber		'募集人数
	Dim dbWorkingPlaceSeq	'勤務地番号
	Dim dbWorkingPlacePrefectureName'勤務地都道府県名
	Dim dbWorkingPlaceCity	'勤務地市区郡
	Dim dbWorkingPlaceAddressAll'勤務地住所全体
	Dim dbWorkingPlaceSection'勤務地部署
	Dim dbWorkingPlaceTelephoneNumber'勤務地TEL
	Dim dbMapFlag			'地図有無フラグ
	Dim dbTransfer			'転勤
	Dim dbPlanTypeName		'ナビライセンス種類
	Dim dbTTPOrderFlag		'紹介予定派遣案件フラグ

	Dim sHTML
	Dim sWorkingType		'勤務形態
	Dim sJobType			'職種
	Dim sSalary				'給与
	Dim sYearlyIncome		'年収
	Dim sMonthlyIncome		'月給
	Dim sDailyIncome		'日給
	Dim sHourlyIncome		'時給
	Dim sTrafficFee			'交通費
	Dim sAfterWorkingType	'紹介後の勤務形態
	Dim sWorkRange			'就業期間
	Dim sWorkUpdate			'就業期間の更新有無
	Dim sWorkingTime		'就業時間
	Dim sMAP				'地図情報
	Dim sWorkingPlace		'就業場所
	Dim sNearbyStation		'最寄駅
	Dim sNearbyRailway		'沿線
	Dim sNearbyStationBlock	'最寄駅,沿線ブロック
	Dim iMaxRow
	Dim sDisplay
	Dim sPlusMinus
	Dim flgFC				'FC・代理店フラグ
	Dim flgSOHOFC
	'</変数宣言>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyKbn = rRS.Collect("CompanyKbn")
	dbPlanTypeName = rRS.Collect("PlanTypeName")
	dbTTPOrderFlag = rRS.Collect("TTPOrderFlag")

	'<勤務形態>
	dbAfterWorkingTypeCode = ChkStr(rRS.Collect("AfterWorkingTypeCode"))
	dbWorkStartDay = ChkStr(rRS.Collect("WorkStartDay"))
	dbWorkEndDay = ChkStr(rRS.Collect("WorkEndDay"))

	'勤務形態
	sWorkingType = GetWorkingType(rDB, rRS)
	flgSOHOFC = False
	If IsRE(sWorkingType,"((SOHO)|(FC))",True) = True Then flgSOHOFC = True

	'紹介後の勤務形態
	sAfterWorkingType = ""
	If dbAfterWorkingTypeCode <> "" Then
		sAfterWorkingType = "※紹介後の勤務形態&nbsp;･･･&nbsp;" & GetDetail("WorkingType", dbAfterWorkingTypeCode)
	End If

	'就業期間
	sWorkRange = ""
	If dbWorkStartDay & dbWorkEndDay <> "" Then
		If dbWorkStartDay <> "" Then sWorkRange = sWorkRange & GetDateStr(dbWorkStartDay, "/")
		If sWorkRange <> "" Then sWorkRange = sWorkRange & "～"
		If dbWorkEndDay <> "" Then sWorkRange = sWorkRange & GetDateStr(dbWorkEndDay, "/")
	End If

	If dbOrderType = "1" Then
		If rRS.Collect("WorkUpdateFlag") = "1" Then
			sWorkUpdate = "有"
		Else
			sWorkUpdate = "無"
		End If
		sWorkRange = sWorkRange & "(更新" & sWorkUpdate & ")"
	End If

	If sAfterWorkingType = "" Then sAfterWorkingType = "<span style=""color:#999999;"">[紹介後の勤務形態]が未入力です。</span>"
	If sWorkRange = "" Then sWorkRange = "<span style=""color:#999999;"">[就業期間]が未入力です。</span>"
	If sWorkingType = "" Then sWorkingType = "<span style=""color:#999999;"">[勤務形態]が未入力です。</span>"
	'</勤務形態>

	'<職種>
	sJobType = GetJobType(rDB, rRS)
	If sJobType = "" Then sJobType = "<span style=""color:#999999;"">[職種]が未入力です。</span>"
	'</職種>

	'<職種詳細>
	dbJobTypeDetail = rRS.Collect("JobTypeDetail")
	'</職種詳細>

	'<給与>
	flgFC = False
	'<ＦＣ・代理店チェック>
	sSQL = "sp_GetDataWorkingType '" & qsOrderCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		If oRS.Collect("WorkingTypeCode") = "006" Or oRS.Collect("WorkingTypeCode") = "007" Then flgFC = True
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
	'</ＦＣ・代理店チェック>

	dbYearlyIncomeMin = ChkStr(rRS.Collect("YearlyIncomeMin"))
	dbYearlyIncomeMax = ChkStr(rRS.Collect("YearlyIncomeMax"))
	If dbYearlyIncomeMin = "0" Then dbYearlyIncomeMin = ""
	If dbYearlyIncomeMax = "0" Then dbYearlyIncomeMax = ""
	If dbYearlyIncomeMin <> "" Then dbYearlyIncomeMin = GetJapaneseYen(dbYearlyIncomeMin)
	If dbYearlyIncomeMax <> "" Then dbYearlyIncomeMax = GetJapaneseYen(dbYearlyIncomeMax)
	If dbYearlyIncomeMin & dbYearlyIncomeMax <> "" Then
		If dbYearlyIncomeMin <> "" Then sYearlyIncome = sYearlyIncome & dbYearlyIncomeMin
		sYearlyIncome = sYearlyIncome & "&nbsp;～&nbsp;"
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
		sMonthlyIncome = sMonthlyIncome & "&nbsp;～&nbsp;"
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
		sDailyIncome = sDailyIncome & "&nbsp;～&nbsp;"
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
		sHourlyIncome = sHourlyIncome & "&nbsp;～&nbsp;"
		If dbHourlyIncomeMax <> "" Then sHourlyIncome = sHourlyIncome & dbHourlyIncomeMax
	End If

	dbPercentagePay = ChkStr(rRS.Collect("PercentagePayFlag"))
	dbSalaryRemark = Replace(ChkStr(rRS.Collect("IncomeRemark")), vbCrLf, "<br>")
	dbSalaryRemark = Replace(dbSalaryRemark, vbCr, "<br>")
	dbSalaryRemark = Replace(dbSalaryRemark, vbLf, "<br>")
	sTrafficFee = ""
	dbTrafficFeeType = ChkStr(rRS.Collect("TrafficFeeType"))
	dbTrafficFeeMonth = ChkStr(rRS.Collect("MonthTrafficFee"))

	'給与
	If sYearlyIncome = "" Then sYearlyIncome = "<span style=""color:#999999;"">[年収]が未入力です。</span>"
	If sMonthlyIncome = "" Then sMonthlyIncome = "<span style=""color:#999999;"">[月給]が未入力です。</span>"
	If sDailyIncome = "" Then sDailyIncome = "<span style=""color:#999999;"">[日給]が未入力です。</span>"
	If sHourlyIncome = "" Then sHourlyIncome = "<span style=""color:#999999;"">[時給]が未入力です。</span>"

	'歩合制
	If dbPercentagePay <> "" Then
		If dbPercentagePay = "1" Then
			dbPercentagePay = "あり"
		ElseIf dbPercentagePay = "0" Then
			dbPercentagePay = "なし"
		End If
	Else
		dbPercentagePay = "<span style=""color:#999999;"">[歩合制]が未入力です。</span>"
	End If

	'交通費
	If ChkStr(rRS.Collect("NaviTrafficPayFlag")) = "1" Then 
		sTrafficFee = "交通費支給あり" & dbTrafficFeeType
		If IsNumber(dbTrafficFeeMonth, 0, False) = True Then
			sTrafficFee = sTrafficFee & "(" & FormatCanma(dbTrafficFeeMonth) & "円／月)"
		End If
	Else
		sTrafficFee = "<span style=""color:#999999;"">[交通費]が未入力です。</span>"
	End If

	If dbSalaryRemark = "" Then dbSalaryRemark = "<span style=""color:#999999;"">[給与備考]が未入力です。</span>"
	'</給与>

	'<時間>
	sWorkingTime = GetWorkingTime(rDB, rRS)
	dbWorkTimeRemark = ChkStr(rRS.Collect("WorkTimeRemark"))

	If sWorkingTime = "" Then sWorkingTime = "<span style=""color:#999999;"">[就業時間]が未入力です。</span>"
	If dbWorkTimeRemark = "" Then dbWorkTimeRemark = "<span style=""color:#999999;"">[就業時間備考]が未入力です。</span>"
	'</時間>

	'<休日>
	dbWeeklyHolidayType = ChkStr(rRS.Collect("WeeklyHolidayTypeName"))
	dbHolidayRemark = ChkStr(rRS.Collect("HolidayRemark"))

	If dbWeeklyHolidayType = "" Then dbWeeklyHolidayType = "<span style=""color:#999999;"">[週休種類]が未入力です。</span>"
	If dbHolidayRemark = "" Then dbHolidayRemark = "<span style=""color:#999999;"">[休日備考]が未入力です。</span>"
	'</休日>

	'<募集人数>
	dbHumanNumber = ChkStr(rRS.Collect("HumanNumber"))

	If dbHumanNumber <> "" Then
		dbHumanNumber = dbHumanNumber & "人"
	Else
		dbHumanNumber = "<span style=""color:#999999;"">[募集人数]が未入力です。</span>"
	End If
	'</募集人数>

	'<勤務地>
	iMaxRow = 0
	sWorkingPlace = ""
	sNearbyStationBlock = ""
	sSQL = "EXEC up_LstC_WorkingPlace '" & dbOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing
		iMaxRow = oRS.RecordCount
		'<最寄駅>
		sSQL = "EXEC up_LstC_NearbyStation '" & dbOrderCode & "', '';"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then Set oRS2.ActiveConnection = Nothing
		'</最寄駅>
		'<最寄沿線>
		sSQL = "EXEC up_LstC_NearbyRailwayLine '" & rRS.Collect("OrderCode") & "','','';"
		flgQE = QUERYEXE(rDB, oRS3, sSQL, sError)
		If GetRSState(oRS3) = True Then Set oRS3.ActiveConnection = Nothing
		'</最寄沿線>
	End If
	Do While GetRSState(oRS) = True
		dbWorkingPlaceSeq = ChkStr(oRS.Collect("WorkingPlaceSeq"))
		dbWorkingPlacePrefectureName = ChkStr(oRS.Collect("WorkingPlacePrefectureName"))
		dbWorkingPlaceCity = ChkStr(oRS.Collect("WorkingPlaceCity"))
		dbWorkingPlaceAddressAll = ChkStr(oRS.Collect("WorkingPlaceAddressAll"))
		dbWorkingPlaceSection = ChkStr(oRS.Collect("WorkingPlaceSection"))
		dbWorkingPlaceTelephoneNumber = ChkStr(oRS.Collect("WorkingPlaceTelephoneNumber"))
		dbMapFlag = ChkStr(oRS.Collect("MapFlag"))

		If sWorkingPlace <> "" And flgSOHOFC = True Then sWorkingPlace = sWorkingPlace & "、"

		'<勤務地>
		sWorkingPlace = sWorkingPlace & "<div"
		If flgSOHOFC = True Then sWorkingPlace = sWorkingPlace & " style=""display:inline;"""
		sWorkingPlace = sWorkingPlace & ">"
		If iMaxRow > 1 And flgSOHOFC = False Then sWorkingPlace = sWorkingPlace & "【勤務地" & dbWorkingPlaceSeq & "】"

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
			If dbMapFlag = "1" Then sWorkingPlace = sWorkingPlace & "&nbsp;[<span style=""color:#0045f9;cursor:pointer;"" onclick=""open('" & HTTP_CURRENTURL & "map/showmap.asp?ordercode=" & dbOrderCode & "&wpseq=" & dbWorkingPlaceSeq & "', 'map', 'width=700,height=650');"">地図</span>]"
		End If

		'<最寄駅>
		sNearbyStation = ""
		oRS2.Filter = "WorkingPlaceSeq = " & dbWorkingPlaceSeq
		If GetRSState(oRS2) = True Then
			sNearbyStation = GetNearbyStation(rDB, oRS2)
		End If
		oRS2.Filter = 0
		'</最寄駅>
		'<最寄沿線>
		sNearbyRailway = ""
		oRS3.Filter = "WorkingPlaceSeq = " & dbWorkingPlaceSeq
		If GetRSState(oRS3) = True Then
			sNearbyRailway = GetNearbyRailway(rDB, oRS3)
		End If
		oRS3.Filter = 0
		'</最寄沿線>

		If sNearbyStation <> "" Then
			sWorkingPlace = sWorkingPlace & "<p class=""m0"""
			If flgSOHOFC = True Then
				sWorkingPlace = sWorkingPlace & " style=""display:inline;"""
			Else
				sWorkingPlace = sWorkingPlace & " style=""padding-left:15px;"""
			End If
			sWorkingPlace = sWorkingPlace & ">"
			sWorkingPlace = sWorkingPlace & "[最寄駅]"
			sWorkingPlace = sWorkingPlace & sNearbyStation
			sWorkingPlace = sWorkingPlace & "<br>"
			sWorkingPlace = sWorkingPlace & "[沿線]"
			sWorkingPlace = sWorkingPlace & sNearbyRailway
			sWorkingPlace = sWorkingPlace & "</p>"
		End If
		'</勤務地>

		sWorkingPlace = sWorkingPlace & "</div>"
		oRS.MoveNext
	Loop

	'転勤
	If (dbOrderType = "0" Or dbOrderType = "2") And dbCompanyKbn <> "4" Then
		'ﾘｽの派遣求人票 または 派遣会社の求人票の場合は表示しない

		dbTransfer = ChkStr(rRS.Collect("Transfer"))
		If dbTransfer <> "" Then
			If dbTransfer = "有" Then
				dbTransfer = "転勤あり"
			ElseIf dbTransfer = "無" Then
				dbTransfer = "転勤なし"
			End If
		End If
	End If

	If sWorkingPlace = "" Then sWorkingPlace = "<span style=""color:#999999;"">[勤務地]が未入力です。</span>"
	If dbTransfer = "" Then dbTransfer = "<span style=""color:#999999;"">[転勤の有無]が未入力です。</span>"
	If sMAP = "" Then sMAP = "<span style=""color:#999999;"">[就業先の地図位置情報]が未登録です。</span>"
	'</勤務地>

	sHTML = sHTML & "<h3>勤務条件</h3>"

	sHTML = sHTML & "<div class=""category1"">"
	sHTML = sHTML & "<h4>"
	sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit07.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
	If G_COMPANYKBN = "1" Then sHTML = sHTML & "<span style=""color:#ff0000; font-weight:normal; font-size:9px;"">※必須</span>"
	sHTML = sHTML & "<br>"
	sHTML = sHTML & "勤務形態"
	sHTML = sHTML & "</h4>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div class=""value1"">"
	sHTML = sHTML & "<a name=""edit07""></a>"
	'<勤務形態>
	sHTML = sHTML & "<p class=""m0"">" & sWorkingType & "</p>"
	'</勤務形態>
	'<紹介後の勤務形態>
	If dbTTPOrderFlag = "1" Then sHTML = sHTML & "<p class=""m0"">" & sAfterWorkingType & "</p>"
	'</紹介後の勤務形態>
	'<就業期間>
	sHTML = sHTML & "<p class=""m0"">" & sWorkRange & "</p>"
	'</就業期間>
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<職種>
	sHTML = sHTML & "<a name=""edit08""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit08.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><span style=""color:#ff0000; font-weight:normal; font-size:9px;"">※必須</span><br>職種</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"
	sHTML = sHTML & "<p class=""m0""><strong>" & dbJobTypeDetail & "</strong></p>"
	sHTML = sHTML & "<p class=""m0"">" & sJobType & "</p>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</職種>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit09.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
	If flgFC = False Then sHTML = sHTML & "<span style=""color:#ff0000; font-weight:normal; font-size:9px;"">※必須</span>"
	sHTML = sHTML & "<br>給与</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"
	sHTML = sHTML & "<a name=""edit09""></a>"
	If flgFC = True Then sHTML = sHTML & "<p class=""m0"" style=""color:#999999;"">※ＦＣ・代理店、ＳＯＨＯ募集のため、給与の入力は必須ではありません。</p>"
	'<年収>
	sHTML = sHTML & "<h5>年収</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sYearlyIncome & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</年収>

	'<月給>
	sHTML = sHTML & "<h5>月給</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sMonthlyIncome & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</月給>

	'<日給>
	sHTML = sHTML & "<h5>日給</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sDailyIncome & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</日給>

	'<時給>
	sHTML = sHTML & "<h5>時給</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sHourlyIncome & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</時給>

	'<歩合制>
	sHTML = sHTML & "<h5>歩合制</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbPercentagePay & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both; margin:0px;""></div>"
	'</歩合制>

	'<交通費>
	sHTML = sHTML & "<h5>交通費</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sTrafficFee & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</交通費>

	'<給与備考>
	sHTML = sHTML & "<h5>給与備考</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbSalaryRemark & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both; margin:0px;""></div>"
	'</給与備考>

	If flgFC = False Then
		sHTML = sHTML & "<p class=""m0"" style=""font-size:10px;"">"
		sHTML = sHTML & "※最低額は条件に関係なく得られる額です。(年収の最低額は条件に関係なく得られる月給の合計です。)"
		sHTML = sHTML & "</p>"
	End If
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit10.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';"">"
	'FC,SOHO案件以外の場合は必須
	If rRS.Collect("FCSOHOOrderFlag") = "0" Then
		sHTML = sHTML & "<span style=""color:#ff0000; font-weight:normal; font-size:9px;"">※必須</span>"
	End If
	sHTML = sHTML & "<br>時間</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"
	'<就業時間>
	sHTML = sHTML & "<a name=""edit10""></a>"
	If flgFC = True Then sHTML = sHTML & "<p class=""m0"" style=""color:#999999;"">※ＦＣ・代理店、ＳＯＨＯ募集のため、就業時間の入力は必須ではありません。</p>"
	sHTML = sHTML & "<h5>就業時間</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sWorkingTime & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</就業時間>

	'<就業時間備考>
	sHTML = sHTML & "<h5>就業時間備考</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbWorkTimeRemark & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</就業時間備考>

	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit11.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>休日</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"

	'<休日>
	sHTML = sHTML & "<a name=""edit11""></a>"
	sHTML = sHTML & "<h5>休日</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbWeeklyHolidayType & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</休日>

	'<休日備考>
	sHTML = sHTML & "<h5>休日備考</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbHolidayRemark & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</休日備考>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<募集人数>
	sHTML = sHTML & "<a name=""edit12""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit12.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>募集人数</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"
	sHTML = sHTML & "<p class=""m0"">" & dbHumanNumber & "</p>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</募集人数>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<勤務地>
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit13.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><span style=""color:#ff0000; font-weight:normal; font-size:9px;"">※必須</span><br>勤務地</h4></div>"
	sHTML = sHTML & "<div class=""value1"">"
	sHTML = sHTML & "<a name=""edit13""></a>"
	sHTML = sHTML & "<h5>住所</h5>"
	sHTML = sHTML & "<div class=""value2"">"
	sHTML = sHTML & "<p class=""m0"">" & sWorkingPlace & "</p>"
	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	sHTML = sHTML & "<h5>転勤</h5>"
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & dbTransfer & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	sHTML = sHTML & "</div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"
	'</勤務地>

	sHTML = sHTML & "<br>"

	GetHTMLEditCondition = sHTML
End Function

'******************************************************************************
'概　要：求人票の必要条件を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vEditFlag		：編集フラグ
'備　考：
'使用元：しごとナビ/company/orderedit/edit0.asp
'履　歴：2008/10/10 LIS K.Kokubo 作成
'　　　：2012/03/12 LIS K.Kokubo 卒業年出力
'******************************************************************************
Function GetHTMLEditNeedCondition(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	'<変数宣言>
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbOrderCode			'情報コード
	Dim dbOrderType			'求人票種類
	Dim dbCompanyKbn		'企業区分
	Dim dbTempOrderFlag		'派遣案件フラグ
	Dim dbAgeMin			'年齢下限
	Dim dbAgeMax			'年齢上限
	Dim dbAgeReasonFlag		'年齢理由フラグ
	Dim dbAgeReason			'年齢理由
	Dim dbAgeReasonDetail	'年齢制限理由詳細
	Dim dbHopeSchoolHistory	'学歴
	Dim dbGraduateYearMin	'卒業年下限
	Dim dbGraduateYearMax	'卒業年上限

	Dim sHTML
	Dim sAge				'年齢制限
	Dim sSchoolHistory		'学歴
	Dim sSkillOS			'ＯＳ
	Dim sSkillApp			'アプリケーション
	Dim sSkillDL			'開発言語
	Dim sSkillDB			'ＤＢ
	Dim sSkillOther			'その他スキル
	Dim sLicense			'資格
	Dim sLicenseOther		'その他資格
	Dim sOtherNote			'その他特記事項
	'</変数宣言>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyKbn = rRS.Collect("CompanyKbn")
	dbTempOrderFlag = rRS.Collect("TempOrderFlag")

	'<年齢>
	sAge = ""
	dbAgeMin = ChkStr(rRS.Collect("AgeMin"))
	dbAgeMax = ChkStr(rRS.Collect("AgeMax"))
	dbAgeReasonFlag = ChkStr(rRS.Collect("AgeReasonFlag"))
	dbAgeReason = ChkStr(rRS.Collect("AgeReason"))
	dbAgeReasonDetail = Replace(ChkStr(rRS.Collect("AgeReasonDetail")), vbCrLf, "<br>")

	If dbOrderType = "1" Or dbTempOrderFlag = "1" Then
		sAge = "派遣案件のため、年齢掲載していません。<br>"
		sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[？]制限について</a>"
	ElseIf dbAgeReasonFlag = "0" Or dbAgeReasonFlag = "" Or (dbAgeMin & dbAgeMax = "") Then
		sAge = "年齢不問<br>"
		'sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[？]制限について</a>"
	Else
		If dbAgeMin <> "" Then dbAgeMin = dbAgeMin & "歳"
		If dbAgeMax <> "" Then dbAgeMax = dbAgeMax & "歳"
		sAge = dbAgeMin & "～" & dbAgeMax
		If dbAgeReason <> "" Then sAge = sAge & "&nbsp;(" & dbAgeReason & ")<br>"
		If dbAgeReasonDetail <> "" Then sAge = sAge & dbAgeReasonDetail & "<br>"
		sAge = sAge & "<a href=""javascript:void(0);"" onclick=""window.open('/infomation/age_limitation_exception_reason.asp','age_limit','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=620,height=400')"">[？]制限について</a><br>"
	End If

	If dbTempOrderFlag = "1" Then
		sAge = "<span style=""color:#999999;"">[年齢関連]派遣案件のため、年齢の入力はできません。</span>"
	ElseIf sAge = "" Then
		sAge = "<span style=""color:#999999;"">[年齢関連]が未入力です。</span>"
	End If
	'</年齢>

	'<学歴>
	dbHopeSchoolHistory = ChkStr(rRS.Collect("HopeSchoolHistory"))
	dbGraduateYearMin = rRS.Collect("GraduateYearMin")
	dbGraduateYearMax = rRS.Collect("GraduateYearMax")
	If dbHopeSchoolHistory <> "" Then
		sSchoolHistory = dbHopeSchoolHistory & "卒以上<br>"

		If dbGraduateYearMin + dbGraduateYearMax > 0 Then
			sSchoolHistory = sSchoolHistory & "[卒業年] "
			If dbGraduateYearMin > 0 Then
				sSchoolHistory = sSchoolHistory & dbGraduateYearMin & "年卒"
			End If
			sSchoolHistory = sSchoolHistory & " ～ "
			If dbGraduateYearMax > 0 Then
				sSchoolHistory = sSchoolHistory & dbGraduateYearMax & "年卒"
			End If
		Else
			sSchoolHistory = sSchoolHistory & "<span style=""color:#999999;"">[卒業年]が未入力です。</span>"
		End If
	Else
		sSchoolHistory = "<span style=""color:#999999;"">[学歴]が未入力です。</span><br>"
		sSchoolHistory = sSchoolHistory & "<span style=""color:#999999;"">[卒業年]が未入力です。</span>"
	End If
	'</学歴>

	'<資格>
	sLicense = GetLicense(rDB, rRS)
	sLicenseOther = GetOrderNote(rDB, rRS, "OtherLicense")

	If sLicense = "" Then sLicense = "<span style=""color:#999999;"">[資格]が未入力です。</span>"
	If sLicenseOther = "" Then sLicenseOther = "<span style=""color:#999999;"">[その他資格]が未入力です。</span>"
	'</資格>

	'<スキル>
	sSkillOS = GetSkill(rDB, rRS, "OS")
	sSkillApp = GetSkill(rDB, rRS, "Application")
	sSkillDL = GetSkill(rDB, rRS, "DevelopmentLanguage")
	sSkillDB = GetSkill(rDB, rRS, "Database")
	sSkillOther = GetOrderNote(rDB, rRS, "OtherSkill")

	If sSkillOS = "" Then sSkillOS = "<span style=""color:#999999;"">[ＯＳ]が未入力です。</span>"
	If sSkillApp = "" Then sSkillApp = "<span style=""color:#999999;"">[アプリケーション]が未入力です。</span>"
	If sSkillDL = "" Then sSkillDL = "<span style=""color:#999999;"">[開発言語]が未入力です。</span>"
	If sSkillDB = "" Then sSkillDB = "<span style=""color:#999999;"">[データベース]が未入力です。</span>"
	If sSkillOther = "" Then sSkillOther = "<span style=""color:#999999;"">[その他スキル]が未入力です。</span>"
	'</スキル>

	'<その他特記事項>
	sOtherNote = ""
	If dbOrderType = "0" Then
		sOtherNote = GetOrderNote(rDB, rRS, "OtherNote")
	End If

	If sOtherNote = "" Then sOtherNote = "<span style=""color:#999999;"">[その他特記事項]が未入力です。</span>"
	'</その他特記事項>

	sHTML = sHTML & "<h3>必要条件</h3>" & vbCrLf

	'<年齢>
	sHTML = sHTML & "<a name=""edit14""></a>"
	sHTML = sHTML & "<div class=""category1"">"
	sHTML = sHTML & "<h4>"
	If dbTempOrderFlag = "0" Then sHTML = sHTML & "<input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit14.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>"
	sHTML = sHTML & "年齢"
	sHTML = sHTML & "</h4>"
	sHTML = sHTML & "</div>" & vbCrLf
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sAge & "</p></div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	'</年齢>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<希望学歴>
	sHTML = sHTML & "<a name=""edit15""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit15.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>希望学歴</h4></div>" & vbCrLf
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sSchoolHistory & "</p></div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	'</希望学歴>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<資格出力>
	sHTML = sHTML & "<a name=""edit16""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit16.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>資格</h4></div>" & vbCrLf
	sHTML = sHTML & "<div class=""value1"">" & vbCrLf

	sHTML = sHTML & "<h5>資格</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2"">" & sLicense & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "<h5>その他資格</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2"">" & sLicenseOther & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	'</資格出力>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<スキル出力>
	sHTML = sHTML & "<a name=""edit17""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit17.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>スキル</h4></div>" & vbCrLf
	sHTML = sHTML & "<div class=""value1"">" & vbCrLf

	sHTML = sHTML & "<h5>ＯＳ</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2"">" & sSkillOS & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "<h5>ｱﾌﾟﾘｹｰｼｮﾝ</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2"">" & sSkillApp & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "<h5>開発言語</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2"">" & sSkillDL & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "<h5>データベース</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2"">" & sSkillDB & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "<h5>その他スキル</h5>" & vbCrLf
	sHTML = sHTML & "<div class=""value2""><p class=""m0"">" & sSkillOther & "</p></div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	sHTML = sHTML & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	'</スキル出力>

	sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"

	'<その他特記事項>
	sHTML = sHTML & "<a name=""edit18""></a>"
	sHTML = sHTML & "<div class=""category1""><h4><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit18.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><br>特記事項</h4></div>" & vbCrLf
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sOtherNote & "</p></div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	'</その他特記事項>

	sHTML = sHTML & "<br>"

	GetHTMLEditNeedCondition = sHTML
End Function

'******************************************************************************
'概　要：求人票の応募情報を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'備　考：
'使用元：しごとナビ/company/orderedit/edit0.asp
'更　新：2008/10/10 LIS K.Kokubo 作成
'******************************************************************************
Function GetHTMLEditHowToEntry(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	Dim dbOrderCode			'情報コード
	Dim dbEntryInfo			'応募方法
	Dim dbProcess1			'STEP1
	Dim dbProcess2			'STEP2
	Dim dbProcess3			'STEP3
	Dim dbProcess4			'STEP4
	Dim sCSectionName		'リス担当部署
	Dim sCPersonName		'リス担当者名
	Dim sCTel				'リス連絡先
	Dim sLis				'リス担当者
	Dim dbWValueURL			'Ｗバリューの自社採用ページＵＲＬ
	Dim sClearSolid
	Dim flgLine				'線引きフラグ

	Dim sHTML

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sOrderType = ChkStr(rRS.Collect("OrderType"))
	dbOrderCode = ChkStr(rRS.Collect("OrderCode"))
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'******************************************************************************
	'応募方法 start
	'------------------------------------------------------------------------------
	dbEntryInfo = Replace(ChkStr(rRS.Collect("EntryInfo")), vbCrLf, "<br>")
	dbEntryInfo = Replace(dbEntryInfo, vbCr, "<br>")
	dbEntryInfo = Replace(dbEntryInfo, vbLf, "<br>")

	If dbEntryInfo = "" Then dbEntryInfo = "<span style=""color:#999999;"">[応募方法]が未入力です。</span>"
	'------------------------------------------------------------------------------
	'応募方法 end
	'******************************************************************************

	'******************************************************************************
	'選考手順 start
	'------------------------------------------------------------------------------
	dbProcess1 = ChkStr(rRS.Collect("Process1"))
	dbProcess2 = ChkStr(rRS.Collect("Process2"))
	dbProcess3 = ChkStr(rRS.Collect("Process3"))
	dbProcess4 = ChkStr(rRS.Collect("Process4"))

	If dbProcess1 = "" Then dbProcess1 = "<span style=""color:#999999;"">[選考手順１]が未入力です。</span>"
	If dbProcess2 = "" Then dbProcess2 = "<span style=""color:#999999;"">[選考手順２]が未入力です。</span>"
	If dbProcess3 = "" Then dbProcess3 = "<span style=""color:#999999;"">[選考手順３]が未入力です。</span>"
	If dbProcess4 = "" Then dbProcess4 = "<span style=""color:#999999;"">[選考手順４]が未入力です。</span>"
	'------------------------------------------------------------------------------
	'選考手順 end
	'******************************************************************************

	'******************************************************************************
	'連絡先 start
	'------------------------------------------------------------------------------
	sCSectionName = ChkStr(rRS.Collect("LisDepartment"))
	sCPersonName = ChkStr(rRS.Collect("EmployeeName"))
	sCTel = ChkStr(rRS.Collect("LisTelephoneNumber"))
	sLis = sCPersonName & "［リス株式会社" & sCSectionName & "］　" & sCTel & "<br>(この案件はリス株式会社が取りまとめています。)"
	'------------------------------------------------------------------------------
	'連絡先 end
	'******************************************************************************

	'******************************************************************************
	'Ｗバリューの自社採用ページＵＲＬ start
	'------------------------------------------------------------------------------
	dbWValueURL = ChkStr(rRS.Collect("WValueURL"))
	'------------------------------------------------------------------------------
	'Ｗバリューの自社採用ページＵＲＬ end
	'******************************************************************************

	flgLine = False

	sHTML = sHTML & "<a name=""edit19""></a>"
	sHTML = sHTML & "<h3>応募情報</h3>" & vbCrLf
	sHTML = sHTML & "<div style=""margin-left:15px;""><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit19.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""></div>"

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>情報コード</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbOrderCode & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>応募方法</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbEntryInfo & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
	flgLine = True

	sHTML = sHTML & "<div class=""category1""><h4>選考手順</h4></div>" & vbCrLf
	sHTML = sHTML & "<div class=""value1"">" & vbCrLf

	If dbProcess1 <> "" Then
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:60px; color:#666666; font-weight:bold;"">ステップ１</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:300px;"">" & dbProcess1 & "</p>"
		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	End If

	If dbProcess2 <> "" Then
		sHTML = sHTML & "<p style=""width:60px; color:#666666; text-align:center;"">▼</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:60px; color:#666666; font-weight:bold;"">ステップ２</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:300px;"">" & dbProcess2 & "</p>"
		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	End If

	If dbProcess3 <> "" Then
		sHTML = sHTML & "<p style=""width:60px; color:#666666; text-align:center;"">▼</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:60px; color:#666666; font-weight:bold;"">ステップ３</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:300px;"">" & dbProcess3 & "</p>"
		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	End If

	If dbProcess4 <> "" Then
		sHTML = sHTML & "<p style=""width:60px; color:#666666; text-align:center;"">▼</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:60px; color:#666666; font-weight:bold;"">ステップ４</p>"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:300px;"">" & dbProcess4 & "</p>"
		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	End If

	sHTML = sHTML & "</div>" & vbCrLf
	sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf

	If dbWValueURL <> "" Then
		If flgLine = True Then sHTML = sHTML & "<table class=""odline1"" border=""0""><tr><td></td></tr></table>"
		flgLine = True

		sHTML = sHTML & "<div class=""category1""><h4>自社採用<br>ページ</h4></div>"
		sHTML = sHTML & "<div class=""value1""><a href=""" & dbWValueURL & """ target=""_blank""><img src=""/img/order/btn_wvalue.gif"" border=""0"" alt=""自社採用ページ""></a></div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
	End If

	sHTML = sHTML & "<br>" & vbCrLf

	GetHTMLEditHowToEntry = sHTML
End Function

'******************************************************************************
'概　要：求人票の担当者連絡先を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vEditFlag		：編集フラグ
'備　考：
'使用元：しごとナビ/company/orderedit/edit0.asp
'履　歴：2008/10/10 LIS K.Kokubo 作成
'　　　：2009/04/02 LIS K.Kokubo メール課金プランの場合は連絡先非表示
'******************************************************************************
Function GetHTMLEditContact(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vApplicationCode)
	'<変数宣言>
	Dim dbOrderCode			'情報コード
	Dim sCompanyCode		'企業コード
	Dim dbCompanyName		'企業名称
	Dim dbCompanyNameF		'企業名称カナ
	Dim dbCompanyKbn		'企業区分
	Dim dbCompanySpeciality	'企業特徴
	Dim dbCSectionName		'仕事の連絡先担当部署
	Dim dbCPersonName		'仕事の連絡先担当者名
	Dim dbCPersonNameF		'仕事の連絡先担当者カナ
	Dim dbCTel				'仕事の連絡先電話番号
	Dim dbCMail				'仕事の連絡先メールアドレス
	Dim sPerson
	Dim sContact
	Dim sOrderType
	Dim dbPlanTypeName
	Dim flgLine				'線引きフラグ

	Dim sHTML
	'</変数宣言>

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sCompanyCode = rRS.Collect("CompanyCode")
	sOrderType = rRS.Collect("OrderType")
	If sOrderType <> "0" Then Exit Function
	dbPlanTypeName = rRS.Collect("PlanTypeName")
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'******************************************************************************
	'会社名 start
	'------------------------------------------------------------------------------
	dbCompanyName = rRS.Collect("CompanyName")
	dbCompanyNameF = rRS.Collect("CompanyName_F")
	dbCompanyKbn = rRS.Collect("CompanyKbn")
	dbCompanySpeciality = rRS.Collect("CompanySpeciality")

	'Call SetOrderCompanyName(dbCompanyName, dbCompanyNameF, sOrderType, dbCompanyKbn, dbCompanySpeciality)
	'------------------------------------------------------------------------------
	'会社名 end
	'******************************************************************************

	'******************************************************************************
	'仕事の連絡先 start
	'------------------------------------------------------------------------------
	If sOrderType = "0" Then
		dbCSectionName = ChkStr(rRS.Collect("ContactSectionName"))
		dbCPersonName = ChkStr(rRS.Collect("ContactPersonName"))
		dbCPersonNameF = ChkStr(rRS.Collect("ContactPersonName_F"))
		dbCTel = ChkStr(rRS.Collect("ContactTelNumber"))
		dbCMail = ChkStr(rRS.Collect("ContactMailAddress"))

		If dbCompanyKbn = "2" Then
			'人材会社の求人票の場合は「名前」＋「人材会社名」
			sPerson = dbCPersonName & "&nbsp;(人材会社：" & dbCompanyName & ")"
		Else
			'一般企業の求人票の場合は「名前」＋「カナ」
			sPerson = dbCPersonName
			If dbCPersonNameF <> "" Then sPerson = sPerson & "(" & dbCPersonNameF & ")"
		End If
	End If

	If dbCSectionName = "" Then dbCSectionName = "<span style=""color:#999999;"">[連絡先の部署名]が未入力です。</span>"
	If sPerson = "" Then sPerson = "<span style=""color:#999999;"">[連絡先の担当者名]が未入力です。</span>"
	If dbCTel = "" Then
		dbCTel = "<span style=""color:#999999;"">[連絡先の電話番号]が未入力です。</span>"
	Else
		dbCTel = dbCTel & "<span style=""font-size:10px;"">　※電話等でのお問い合わせの際、「しごとナビを見た」と言うとスムーズです。</span>"
	End If
	If dbCMail = "" Then dbCMail = "<span style=""color:#999999;"">[連絡先のメールアドレス]が未入力です。</span>"

	sContact = ""
	If dbCTel <> "" Then sContact = sContact & dbCTel
	If sContact <> "" Then sContact = sContact & "<br>"
	If dbCMail <> "" Then sContact = sContact & dbCMail
	'------------------------------------------------------------------------------
	'仕事の連絡先
	'******************************************************************************

	flgLine = False
	sHTML = sHTML & "<a name=""edit20""></a>"
	sHTML = sHTML & "<h3 class=""sp"">担当者情報</h3>"
	sHTML = sHTML & "<div style=""margin-left:15px;""><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/orderedit/edit20.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""><span style=""color:#ff0000; font-weight:normal; font-size:9px;"">※必須</span></div>"

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
	sHTML = sHTML & "<div class=""category1""><h4>担当者</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sPerson & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	If flgLine = True Then sHTML = sHTML & "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
	flgLine = True
	sHTML = sHTML & "<div class=""category1""><h4>担当部署</h4></div>"
	sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & dbCSectionName & "</p></div>"
	sHTML = sHTML & "<div style=""clear:both;""></div>"

	If dbPlanTypeName <> "mail" Then
		'メール課金プランの場合は連絡先非表示
		If flgLine = True Then sHTML = sHTML & "<table class=""odline1sp"" border=""0""><tr><td></td></tr></table>"
		flgLine = True
		sHTML = sHTML & "<div class=""category1""><h4>連絡先</h4></div>"

		sHTML = sHTML & "<div class=""value1""><p class=""m0"">" & sContact & "</p></div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
	End If

	sHTML = sHTML & "<br>"

	GetHTMLEditContact = sHTML
End Function

'******************************************************************************
'概　要：求人票詳細の先輩インタビューを出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vEditFlag		：編集フラグ
'備　考：
'使用元：しごとナビ/company/orderedit/edit0.asp
'更　新：2008/10/10 LIS K.Kokubo 作成
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
	sHTML = sHTML & "<h3>先輩インタビュー</h3>"
	sHTML = sHTML & "<div style=""margin-left:15px;""><input class=""btn1"" type=""button"" value=""編　集"" onclick=""location.href='" & HTTPS_CURRENTURL & "company/elderinterview/list.asp?ordercode=" & dbOrderCode & "&amp;appcode=" & vApplicationCode & "';""></div>"

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
				'先輩写真有り
				sHTML = sHTML & "<div style=""width:580px; margin-left:20px;"">"
				sHTML = sHTML & "<div style=""float:left; width:182px; padding-top:5px;"">"
				sHTML = sHTML & "<img src=""/company/elderinterview/picture.asp?ordercode=" & dbOrderCode & "&amp;seq=" & dbSeq & """ alt="""" border=""1"" width=""180"" height=""135"" style=""border:1px solid:#999999;"">"
				sHTML = sHTML & "</div>"
				sHTML = sHTML & "<div style=""float:left; width:398px;"">"
				sHTML = sHTML & "<p style=""margin:0px; padding-left:5px;"">■" & dbQuestion & "</p>"
				sHTML = sHTML & "<p style=""margin:0px; padding-left:5px;"">" & dbAnswer & "</p>"
				sHTML = sHTML & "</div>"
				sHTML = sHTML & "<div style=""clear:both;""></div>"
				sHTML = sHTML & "</div>"
			Else
				'先輩写真無し
				sHTML = sHTML & "<p class=""m0"">■" & dbQuestion & "</p>"
				sHTML = sHTML & "<p class=""m0"">" & dbAnswer & "</p>"
			End If
			oRS.MoveNext
		Loop

		sHTML = sHTML & "</div>"
	Else
		sHTML = sHTML & "<div style=""margin-left:15px;""><span style=""color:#999999;"">[先輩インタビュー]が未入力です。</span></div>"
	End If

	sHTML = sHTML & "<br>"

	GetHTMLElderInterview = sHTML
End Function

'******************************************************************************
'概　要：求人票詳細ページの勤務形態部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
'使用元：staff/company_detail.asp
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

		'リス紹介or紹介会社'従来版If (rRS.Fields("OrderType") ="" and rRS.Fields("Companykbn") = "2") or (rRS.Fields("OrderType") ="2") Then
		If (rRS.Collect("OrderType") ="0" And rRS.Collect("Companykbn") = "2") Or (rRS.Collect("OrderType") ="2") Then
			Select Case oRS.Collect("WorkingTypeCode")
				Case "001": sWorkingType = sWorkingType & "【<a href=""javascript:void(0)"" onclick='window.open(""/staff/koyoukeitai_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>派遣とは</a>】" 
				Case "002","003": sWorkingType = sWorkingType & "【<a href=""javascript:void(0)"" onclick='window.open(""/staff/s_shokai.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>人材紹介とは</a>】" 
				Case "004": sWorkingType = sWorkingType & "【<a href=""javascript:void(0)"" onclick='window.open(""/staff/syoukaiyotei_memo.htm"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=650,height=510"")'>紹介予定派遣とは</a>】" 
			End Select
		End If

		oRS.MoveNext
		If GetRSState(oRS) = True Then sWorkingType = sWorkingType & "<br>"
	Loop
	Call RSClose(oRS)

	GetWorkingType = sWorkingType
End Function

'******************************************************************************
'概　要：求人票詳細ページの職種部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
'使用元：staff/company_detail.asp
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
'概　要：求人票詳細ページの勤務形態部分
'引　数：rDB	：接続中のDBConnection
'　　　：rRS	：up_DtlOrderで生成されたレコードセットオブジェクト
'備　考：
'更　新：2006/05/08 LIS K.Kokub 作成
'　　　：2009/11/17 LIS K.Kokubo FC,SOHO案件の場合は勤務時間を返さない
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
			sWorkingTime = sWorkingTime & sWST & "～" & sWET
		End If
		oRS.MoveNext
		If GetRSState(oRS) = True And sWST & sWET <> "" Then sWorkingTime = sWorkingTime & "<br>"
	Loop
	Call RSClose(oRS)

	GetWorkingTime = sWorkingTime
End Function

'******************************************************************************
'概　要：求人票詳細ページの最寄駅部分
'引　数：rDB	：接続中のDBConnection
'　　　：rRS	：up_LstC_NearbyStationで生成されたレコードセットオブジェクト
'　　　：vWPSeq	：勤務地番号
'使　用：ナビ/include/func_order.asp
'備　考：
'更　新：2006/05/08 LIS K.Kokubo 作成
'　　　：2008/10/22 LIS K.Kokubo 求人票勤務地複数化対応
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
		If dbToStationTime <> "" Then sToStation = dbToStationTime & "分"
		If dbToStationRemark <> "" Then sToStation = dbToStationRemark & sToStation
		If sToStation <> "" Then sToStation = "(" & sToStation & ")"

		If sStation <> "" Then sStation = sStation & ","
		sStation = sStation & dbStationName & "駅" & sToStation

		rRS.MoveNext
	Loop

	GetNearbyStation = sStation
End Function

'******************************************************************************
'概　要：求人票詳細ページの最寄沿線部分
'引　数：rDB	：接続中のDBConnection
'　　　：rRS	：up_LstC_NearbyRailwayLineで生成されたレコードセットオブジェクト
'使　用：ナビ/include/func_order.asp
'備　考：
'更　新：2006/05/08 LIS K.Kokubo 作成
'　　　：2008/10/22 LIS K.Kokubo 求人票勤務地複数化対応
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
	If iRowCnt > 3 Then sRailway = sRailway & "&nbsp;他"

	GetNearbyRailway = sRailway
End Function

'******************************************************************************
'概　要：求人票詳細ページのスキル部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
'使用元：
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
			sSkill = sSkill & "<br>　<span style=""color:#339933;"">■</span>" & oRS.Collect("Period") & "年以上は尚可"
		End If
		sSkill = sSkill & "</p>"
		If iSkill Mod SKILLCOL = 0 Then sSkill = sSkill & "<br clear=""all"">"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'中途半端で終わった場合の調整
	If sSkill <> "" And iSkill Mod SKILLCOL <> 0 Then sSkill = sSkill & "<br clear=""all"">"

	GetSkill = sSkill
End Function

'******************************************************************************
'概　要：求人票詳細ページの資格部分
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
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

	'中途半端で終わった場合の調整
	If sLicense <> "" And iLicense Mod LICENSECOL <> 0 Then sLicense = sLicense & "<br clear=""all"">"

	GetLicense = sLicense
End Function

'******************************************************************************
'概　要：求人票詳細ページのその他情報取得
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vCode			：C_Noteテーブルの Code フィールド値
'作成者：Lis Kokubo
'作成日：2006/05/08
'備　考：
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
'概　要：求人票詳細のタイトルとディスクリプションを取得
'作成者：Lis Kokubo
'作成日：2007/02/12
'戻り値：rTitle			：タイトル（具体的職種名）
'　　　：rDescription	：説明文（担当業務）
'使用元：しごとナビ/order/order_detail.asp
'備　考：
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
		rKeywords = "求人情報,転職," & ChkStr(oRS.Collect("PrefectureName"))
		If ChkStr(oRS.Collect("JobTypeName")) <> "" Then rKeywords = rKeywords & "," & ChkStr(oRS.Collect("JobTypeName"))
		If ChkStr(oRS.Collect("WorkingTypeName")) <> "" Then rKeywords = rKeywords & "," & ChkStr(oRS.Collect("WorkingTypeName"))
		rDescription = "転職・求人情報：" & ChkStr(oRS.Collect("BusinessDetail"))
		If rDescription = "" Then rDescription = "転職・求人情報：" & ChkStr(oRS.Collect("JobTypeDetail"))
	End If
	Call RSClose(oRS)

	If rTitle <> "" Then rTitle = rTitle & "&nbsp;"
	rTitle = rTitle & sWorkingType

	GetOrderTitle = flgQE
End Function

'******************************************************************************
'概　要：スキルの各項目表示
'作成者：Lis Kokubo
'作成日：2007/02/14
'戻り値：
'　　　：
'使用元：しごとナビ/order/order_detail.asp
'備　考：
'******************************************************************************
Function GetSkillList(ByVal vTitleImg, ByVal vTitleAlt, ByVal vSkill)
	GetSkillList = ""
	If Len(vSkill) = 0 Then Exit Function
	GetSkillList = "<tr><td valign=""top""><img src=""" & vTitleImg & """ alt=""" & vTitleAlt & """ width=""50"" height=""12""></td><td style=""padding-left:5px;"">" & vSkill & "</td></tr>"
End Function

'******************************************************************************
'概　要：レコメンドお仕事情報一覧出力
'引　数：rDB		：DB接続オブジェクト
'　　　：vUserType	：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID	：利用中ユーザのユーザID [Session("userid")]
'　　　：vOrderCode	：閲覧中求人票の情報コード
'　　　：vRCMD		：レコメンド種類 ["1"]こんなお仕事情報も見てます ["2"]近い条件のお仕事情報
'　　　：vMyOrder	：自社求人票か否か ["1"]自社求人票
'戻り値：
'作成日：2007/05/31
'作成者：Lis Kokubo
'備　考：
'更　新：
'******************************************************************************
Function DspRecommendOrderList(ByRef rDB, ByVal vUserType, ByVal vUserID, ByVal vOrderCode, ByVal vRCMD, ByVal vMyOrder)
	Const MAXCOLS = 3

	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sTitle
	Dim idx			'ループカウントアップ変数
	Dim iCols		'列数
	Dim aPadding(2)	'各列のパディング
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
			sTitle = "この求人情報を見た人はこんな求人情報も見ています"
		Case "2"
			sSQL = "up_SearchHighRelationOrder '" & CONF_OrderCode & "'"
			sTitle = "この求人情報の条件に近い求人情報"
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
'概　要：自社求人票の掲載状態を変更する
'引　数：rDB			：接続中のDBConnection
'　　　：vOrderCodes	：更新対象の情報コード群（カンマ区切り）
'　　　：vPublicFlags	：更新対象の公開フラグ群（カンマ区切り）
'作成者：Lis Kokubo
'作成日：2007/04/02
'備　考：
'使用元：しごとナビ/order/order_list_entity.asp
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
'概　要：自社求人票を削除する
'引　数：rDB			：接続中のDBConnection
'　　　：vOrderCodes	：更新対象の情報コード群（カンマ区切り）
'作成者：Lis Kokubo
'作成日：2007/04/02
'備　考：
'使用元：しごとナビ/order/order_list_entity.asp
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
'概　要：求人票の特徴
'引　数：rDB
'　　　：rRS
'戻り値：
'備　考：
'履　歴：2008/10/08 LIS K.Kokubo 作成
'　　　：2009/03/18 LIS K.Kokubo 特徴追加(ナビ無料化対応)
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
	'アクセス数が100を超えていれば「HOT」表示（リス安藤）
	If rRS.Collect("AccessCount") > 100 Then sHTML = sHTML & "<img src=""/img/c_HOT_green.gif"" alt=""人気"" width=""50"" height=""15"">&nbsp;"
	'UPDATEと今日から10日引いた日で「新着」表示(リス安藤)
	If rRS.Collect("Updateday") > NOW()-10 Then sHTML = sHTML & "<img src=""/img/c_NEW_green.gif"" alt=""新着"" width=""50"" height=""15"">&nbsp;"
	'未経験者ＯＫの場合、わかばマーク表示(リス安藤)
	If rRS.Collect("InexperiencedPersonFlag") = "1" Then sHTML = sHTML & "<img src=""/img/no_experience.gif"" alt=""未経験者／第二新卒歓迎"" width=""50"" height=""15"">&nbsp;"
	'Ｕターン・Ｉターン
	If rRS.Collect("UITurnFlag") = "1" Then sHTML = sHTML & "<img src=""/img/ui_turn.gif"" alt=""Ｕターン・Ｉターン"" width=""50"" height=""15"">&nbsp;"
	'語学を活かす仕事
	If rRS.Collect("UtilizeLanguageFlag") = "1" Then sHTML = sHTML & "<img src=""/img/linguistic_job.gif"" alt=""語学を活かす仕事"" width=""50"" height=""15"">&nbsp;"
	'年間休日120日以上
	If rRS.Collect("ManyHolidayFlag") = "1" Then sHTML = sHTML & "<img src=""/img/year_holidaycnt.gif"" alt=""年間休日120日以上"" width=""50"" height=""15"">&nbsp;"
	'2006/01/10 M.Hayashi ADD フレックスタイム制度あり
	If rRS.Collect("FlexTimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_flextime.gif"" alt=""フレックスタイム制度あり"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NearStationFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_nearstation.gif"" alt=""駅近(徒歩5分以内)"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NoSmokingFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_nosmoking.gif"" alt=""禁煙・分煙"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("NewlyBuiltFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_newlybuilt.gif"" alt=""新築ビル・オフィス(5年以内)"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("LandmarkFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_landmark.gif"" alt=""高層(15階以上)ビル"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("RenovationFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_renovation.gif"" alt=""リノベーションビル・オフィス(5年以内)"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("DesignersFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_designers.gif"" alt=""デザイナーズビル・オフィス"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("CompanyCafeteriaFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_companycafeteria.gif"" alt=""社員食堂"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("ShortOvertimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_shortovertime.gif"" alt=""残業10h/月以内"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("MaternityFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_maternity.gif"" alt=""産休・育休実績あり"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("DressFreeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_dressfree.gif"" alt=""服装自由"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("MammyFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_mammy.gif"" alt=""子育てママ歓迎"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("FixedTimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_fixedtime.gif"" alt=""18時までに退社"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("ShortTimeFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_shorttime.gif"" alt=""1日6時間以内労働"" width=""50"" height=""15"">&nbsp;"
	'2008/08/19 LIS M.Hayashi ADD 
	If rRS.Collect("HandicappedFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_handicapped.gif"" alt=""障害者歓迎"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("RentAllFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_rentallflag.gif"" alt=""住宅費用全額補助あり"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("RentPartFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_rentpartflag.gif"" alt=""住宅費用一部補助あり"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MealsFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_mealsflag.gif"" alt=""食事・賄い付き案件"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MealsAssistanceFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_mealsassistanceflag.gif"" alt=""食事補助制度あり"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("TrainingCostFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_trainingcostflag.gif"" alt=""研修費助成制度あり"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("EntrepreneurCostFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_entrepreneurcostflag.gif"" alt=""起業機材補助制度あり"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("MoneyFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_moneyflag.gif"" alt=""無利子・低利子補助制度あり"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("LandShopFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_landshopflag.gif"" alt=""土地・店舗等提供制度あり"" width=""50"" height=""15"">&nbsp;"
	'2009/03/18 LIS K.Kokubo ADD 
	If rRS.Collect("FindJobFestiveFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_findjobfestiveflag.gif"" alt=""就職お祝い金制度あり"" width=""50"" height=""15"">&nbsp;"
	'2009/12/01 LIS K.Kokubo ADD 
	If rRS.Collect("AppointmentFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_appointmentflag.gif"" alt=""正社員登用制度あり"" width=""50"" height=""15"">&nbsp;"
	'2009/12/01 LIS K.Kokubo ADD 
	If rRS.Collect("SocietyInsuranceFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order_detail_icon/oc_societyinsuranceflag.gif"" alt=""社保完備"" width=""50"" height=""15"">&nbsp;"
	'2008/05/08 LIS K.Kokubo ADD シークレット求人
	If rRS.Collect("SecretFlag") = "1" Then sHTML = sHTML & "<img src=""/img/order/secret.gif"" alt=""スカウトを受けた人だけが閲覧できる求人情報"" width=""50"" height=""15"">&nbsp;"

	GetImgOrderSpeciality = sHTML
End Function

'******************************************************************************
'概　要：求人票入力画面の雇用形態部分
'引　数：rDB		：
'　　　：vUserType	：
'　　　：vOrderCode	：
'戻り値：
'作成者：Lis K.Kokubo
'作成日：2007/03/27
'備　考：
'使用元：しごとナビ/company/company_reg2.asp
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

	'リストボックス出力個数指定
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
			'一般企業 or 紹介企業 の場合は、「派遣・紹介予定派遣」を非表示
			If sHTML <> "" Then sHTML = sHTML & "<div style=""float:left; width:3px;""></div>" & vbCrLf
			sHTML = sHTML & "<div style=""float:left; width:198px;"">"
			sHTML = sHTML & "<p class=""m0"" style=""text-align:center;"">勤務形態" & idx & "</p>"
			sHTML = sHTML & "<select name=""frmworkingtypecode" & idx & """ size=""6"" style=""width:98%;"" onchange=""ChkDuplication('frmworkingtypecode', '希望勤務形態');"">"
			sHTML = sHTML & "<option value="""">選択して下さい</option>"
			sHTML = sHTML & "<option value=""002"""
			If dbWorkingTypeCode = "002" Then sHTML = sHTML & " selected"
			sHTML = sHTML & ">正社員</option>"
			sHTML = sHTML & "<option value=""003"""
			If dbWorkingTypeCode = "003" Then sHTML = sHTML & " selected"
			sHTML = sHTML & ">契約社員</option>"
			sHTML = sHTML & "<option value=""005"""
			If dbWorkingTypeCode = "005" Then sHTML = sHTML & " selected"
			sHTML = sHTML & ">パート・アルバイト</option>"
			sHTML = sHTML & "<option value=""006"""
			If dbWorkingTypeCode = "006" Then sHTML = sHTML & " selected"
			sHTML = sHTML & ">SOHO（在宅アルバイト・副業）</option>"
			sHTML = sHTML & "<option value=""007"""
			If dbWorkingTypeCode = "007" Then sHTML = sHTML & " selected"
			sHTML = sHTML & ">FC・代理店</option>"
			sHTML = sHTML & "</select>"
			sHTML = sHTML & "</div>" & vbCrLf
		ElseIf vUserType = "dispatch" Then
			'派遣企業 の場合は「派遣・紹介予定派遣」のみ表示
			sHTML = sHTML & "<select name=""frmworkingtypecode" & idx & """ onchange=""ChkDuplication('frmworkingtypecode', '希望勤務形態');"">"
			sHTML = sHTML & "<option value="">選択して下さい</option>"
			sHTML = sHTML & "<option value=""001"""
			If dbWorkingTypeCode = "001" Then Response.Write " selected"
			sHTML = sHTML & ">派遣</option>"
			sHTML = sHTML & "<option value=""004"""
			If dbWorkingTypeCode = "004" Then Response.Write " selected"
			sHTML = sHTML & ">紹介予定派遣</option>"
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
