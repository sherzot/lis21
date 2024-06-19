<%
'**********************************************************************************************************************
'概　要：人材一覧 /staff/person_result.asp
'　　　：プロフィール /staff/person_detail.asp
'　　　：上記ページで出力用の関数群をこのファイルに用意する。
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　出力用　■■■
'　　　：GetHtmlNearbyStation	：基本情報ページの最寄駅一覧ＨＴＭＬを取得
'　　　：GetHtmlResumeListBase	：学歴・職歴・資格一覧の１行分のＨＴＭＬを取得
'　　　：GetHtmlSkillList		：スキルページでスキル大分類１つ分のＨＴＭＬを取得
'**********************************************************************************************************************

'******************************************************************************
'概　要：各編集項目へのリンクボタンＨＴＭＬ
'引　数：
'使用元：しごとナビ/staff/edit/edit1.asp
'　　　：しごとナビ/staff/edit/edit2.asp
'　　　：しごとナビ/staff/edit/edit3.asp
'　　　：しごとナビ/staff/edit/edit4.asp
'　　　：しごとナビ/staff/edit/edit5.asp
'　　　：しごとナビ/staff/edit/edit6.asp
'　　　：しごとナビ/staff/edit/edit7.asp
'　　　：しごとナビ/staff/edit/edit8.asp
'　　　：しごとナビ/staff/edit/edit9.asp
'備　考：
'更　新：2008/03/11 LIS K.Kokubo 作成
'******************************************************************************
Function GetHtmlStaffEditList(ByVal vStaffCode, ByVal vCurrentURL)
	Dim sHTML
	Dim sDisabled

	sHTML = ""

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit1.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit1.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""基本情報""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit2.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit2.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""学歴・職歴""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit3.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit3.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""資格""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit4.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit4.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""スキル""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit5.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit5.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""IT系職歴""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit6.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit6.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""希望条件""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit7.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit7.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""志望動機""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit8.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit8.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""得意分野等""></form>"

	sDisabled = ""
	If InStr(vCurrentURL, "/staff/edit/edit9.asp") > 0 Then sDisabled = " disabled=""disabled"""
	sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit9.asp?staffcode=" & vStaffCode & """" & sDisabled & " style=""display:inline;""><input type=""submit"" value=""自己ＰＲ""></form>"

	sHTML = sHTML & "<br>"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/resume.asp?papersize=a3"" target=""_blank"">ＪＩＳ履歴書Ａ３</a>&nbsp;"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/resume.asp?papersize=a4"" target=""_blank"">ＪＩＳ履歴書Ａ４</a>&nbsp;"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/resume.asp?papersize=b4"" target=""_blank"">履歴書Ｂ４</a>&nbsp;"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/resume.asp?papersize=b5"" target=""_blank"">履歴書Ｂ５</a>&nbsp;"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/resume.asp?papersize=b42"" target=""_blank"">バイト履歴書Ｂ４</a>&nbsp;"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/experience.asp"" target=""_blank"">職務経歴書</a>&nbsp;"
	sHTML = sHTML & "<a href=""" & HTTPS_CURRENTURL & "staff/pdf/experienceit.asp"" target=""_blank"">ＩＴ系職務経歴書</a>&nbsp;"

	GetHtmlStaffEditList = sHTML
End Function

'******************************************************************************
'概　要：基本情報の最寄駅部分のＨＴＭＬを取得
'引　数：rDB		：接続中のDBコネクション
'　　　：vStaffCode	：求職者コード
'使用元：しごとナビ/staff/edit/edit1.asp
'備　考：
'更　新：2008/03/21 LIS K.Kokubo 作成
'******************************************************************************
Function GetHtmlNearbyStation(ByRef rDB, ByVal vStaffCode)
	Dim sSQl
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbSeq
	Dim dbStationCode			'最寄駅：駅コード
	Dim dbStationName			'最寄駅：駅名称
	Dim dbToStationBusFlag		'最寄駅までの交通手段：バス
	Dim dbToStationCarFlag		'最寄駅までの交通手段：車
	Dim dbToStationBicycleFlag	'最寄駅までの交通手段：自転車
	Dim dbToStationWalkFlag		'最寄駅までの交通手段：徒歩
	Dim dbOtherTransportation	'最寄駅までの交通手段：その他
	Dim dbToStationTime			'最寄駅までの時間

	Dim sHTML
	Dim sToStation
	Dim idx
	Dim flgDspAddButton

	sHTML = ""
	sToStation = ""
	flgDspAddButton = True

	sSQL = "sp_GetDataNearbyStation '" & G_USERID & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.RecordCount > 1 Then flgDspAddButton = False
	End If

	'「追加」ボタン
	If flgDspAddButton = True Then
		sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit1_2.asp?flag=1"" method=""post"" style=""display:inline;"">"
		sHTML = sHTML & "<input type=""submit"" value=""追　加"">"
		sHTML = sHTML & "</form><br>" & vbCrLf
	Else
		sHTML = sHTML & "<input type=""button"" disabled=""disabled"" value=""追　加"" style=""margin:0px;"">&nbsp;"
		sHTML = sHTML & "<span style=""color:#333333;"">※最寄駅は２つまでです。</span>" & vbCrLf
	End If

	idx = 0
	Do While GetRSState(oRS) = True And idx <= 1
		dbSeq = oRS.Collect("ID")
		dbStationCode = oRS.Collect("StationCode")
		dbStationName = oRS.Collect("StationName") & "駅"
		dbToStationBusFlag = oRS.Collect("ToStationBusFlag")
		dbToStationCarFlag = oRS.Collect("ToStationCarFlag")
		dbToStationBicycleFlag = oRS.Collect("ToStationBicycleFlag")
		dbToStationWalkFlag = oRS.Collect("ToStationWalkFlag")
		dbOtherTransportation = oRS.Collect("OtherTransportation")
		dbToStationTime = oRS.Collect("ToStationTime")

		sHTML = sHTML & "<div style=""padding:3px 0px; border-bottom:1px dotted #333333;"">"
		sHTML = sHTML & "<p class=""m0"" style=""float:left; width:350px;"">"
		sHTML = sHTML & dbStationName

		sToStation = ""
		If dbToStationBusFlag & dbToStationCarFlag & dbToStationBicycleFlag & dbToStationWalkFlag & dbOtherTransportation <> "" Then
			If dbToStationBusFlag = "1" Then
				If sToStation <> "" Then sToStation = sToStation & ",&nbsp;"
				sToStation = sToStation & "バス"
			End If

			If dbToStationCarFlag = "1" Then
				If sToStation <> "" Then sToStation = sToStation & ",&nbsp;"
				sToStation = sToStation & "車"
			End If

			If dbToStationBicycleFlag = "1" Then
				If sToStation <> "" Then sToStation = sToStation & ",&nbsp;"
				sToStation = sToStation & "自転車"
			End If

			If dbToStationWalkFlag = "1" Then
				If sToStation <> "" Then sToStation = sToStation & ",&nbsp;"
				sToStation = sToStation & "徒歩"
			End If

			If dbOtherTransportation <> "" Then
				If sToStation <> "" Then sToStation = sToStation & ",&nbsp;"
				sToStation = sToStation & dbOtherTransportation
			End If
		End If

		If dbToStationTime <> "" And dbToStationTime > "0" Then
			sToStation = sToStation & "&nbsp;&nbsp;" & dbToStationTime & "分"
		End If

		If sToStation <> "" Then sToStation = "&nbsp;(" & sToStation & ")"

		sHTML = sHTML & sToStation
		sHTML = sHTML & "</p>" & vbCrLf

		sHTML = sHTML & "<div align=""right"" style=""float:right; width:140px;"">"
		sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit1_2.asp?flag=1&amp;seq=" & dbSeq & """ method=""post"" style=""display:inline;"">"
		sHTML = sHTML & "<input type=""submit"" value=""編　集"">"
		sHTML = sHTML & "</form>&nbsp;"
		sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit1_2.asp?flag=0&amp;seq=" & dbSeq & """ method=""post"" style=""display:inline;"" onsubmit=""confirm('「" & dbStationName & "」を削除しますか？');"">"
		sHTML = sHTML & "<input type=""submit"" value=""削　除"">"
		sHTML = sHTML & "</form><br>"
		sHTML = sHTML & "</div>" & vbCrLf

		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		sHTML = sHTML & "</div>" & vbCrLf

		oRS.MoveNext
		idx = idx + 1
	Loop
	Call RSClose(oRS)

	GetHtmlNearbyStation = sHTML
End Function

'******************************************************************************
'概　要：学歴・職歴一覧の１行分のＨＴＭＬを取得
'引　数：vYear	：１列目（年）
'　　　：vMonth	：２列目（月）
'　　　：vBody	：３列目（学歴内容、職歴内容）
'　　　：vRow	：４列目（学歴職歴の行数）
'使用元：しごとナビ/staff/edit/edit2.asp
'備　考：
'更　新：2008/02/28 LIS K.Kokubo 作成
'******************************************************************************
Function GetHtmlResumeListBase(ByVal vYear, ByVal vMonth, ByVal vBody, ByVal vRow)
	Dim sHTML

	sHTML = ""
	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td style=""vertical-align:top; font-size:10px;"">" & vRow & "行</td>"
	sHTML = sHTML & "<td style=""border:1px solid #333333; text-align:right; vertical-align:top;"">" & vYear & "</td>"
	sHTML = sHTML & "<td style=""border:1px solid #333333; text-align:right; vertical-align:top;"">" & vMonth & "</td>"
	sHTML = sHTML & "<td style=""border:1px solid #333333; text-align:left; vertical-align:top;"">" & vBody & "</td>"
	sHTML = sHTML & "</tr>" & vbCrLf

	GetHtmlResumeListBase = sHTML
End Function

'******************************************************************************
'概　要：スキル大分類１つ分のＨＴＭＬを取得
'引　数：rDB			：接続中ＤＢオブジェクト
'　　　：vStaffCode		：ログイン中求職者コード
'　　　：vCategoryCode	：スキル大分類コード
'使用元：しごとナビ/staff/edit/edit4.asp
'備　考：
'更　新：2008/03/06 LIS K.Kokubo 作成
'******************************************************************************
Function GetHtmlSkillList(ByRef rDB, ByVal vStaffCode, ByVal vCategoryCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbCategoryCode
	Dim dbSeq
	Dim dbCode
	Dim dbSkillName
	Dim dbStartDay
	Dim dbPeriod

	Dim sHTML
	Dim sCategoryName
	Dim iRecordCount
	Dim iMaxCount

	Select Case vCategoryCode
		Case "OS": sCategoryName = "ＯＳ": iMaxCount = 10
		Case "Application": sCategoryName = "アプリケーション": iMaxCount = 10
		Case "DevelopmentLanguage": sCategoryName = "開発言語": iMaxCount = 10
		Case "Database": sCategoryName = "データベース": iMaxCount = 8
	End Select

	sHTML = ""

	sSQL = "sp_GetDataSkill '" & vStaffCode & "', '" & vCategoryCode & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	iRecordCount = 0
	If GetRSState(oRS) = True Then
		'レコードセットの切断
		Set oRS.ActiveConnection = Nothing

		iRecordCount = oRS.RecordCount
	End If

	'ＯＳ、アプリケーション、開発言語は１０個まで登録可能
	'データベースは８個まで登録可能
	If iRecordCount < iMaxCount Then
		sHTML = sHTML & "<form class=""m0"" action=""" & HTTPS_CURRENTURL & "staff/edit/edit4_1.asp?staffcode=" & vStaffCode & "&amp;flag=0&amp;categorycode=" & vCategoryCode & """ method=""post"" style=""display:inline;"">"
		sHTML = sHTML & "<input type=""submit"" value=""追　加"">"
		sHTML = sHTML & "</form>&nbsp;"
	Else
		sHTML = sHTML & "<input type=""submit"" disabled=""disabled"" value=""追　加"">&nbsp;"
	End If
	sHTML = sHTML & "※" & sCategoryName & "は" & iMaxCount & "個まで登録できます。"

	If oRS.RecordCount = 0 Then
		sHTML = sHTML & "登録がありません"
	End If

	Do While GetRSState(oRS) = True
		dbCategoryCode = oRS.Collect("CategoryCode")
		dbSeq = oRS.Collect("ID")
		dbCode = oRS.Collect("Code")
		dbSkillName = oRS.Collect("SkillName")
		dbStartDay = oRS.Collect("StartDay")
		dbPeriod = oRS.Collect("Period")

		sHTML = sHTML & "<div style=""margin-bottom:3px; border-bottom:1px dotted #999999;"">"
		sHTML = sHTML & "<div style=""float:left; width:350px;"">" & vbCrLf
		sHTML = sHTML & "<p class=""m0"">"
		sHTML = sHTML & "<span style=""background-color:#ccccff;"">"
		sHTML = sHTML & dbSkillName
		If dbPeriod <> "" Then sHTML = sHTML & "&nbsp;使用期間(" & dbPeriod & "年)"
		sHTML = sHTML & "</span>"
		sHTML = sHTML & "</p>" & vbCrLf
		sHTML = sHTML & "</div>" & vbCrLf
		sHTML = sHTML & "<div align=""right"" style=""float:left; width:140px;"">" & vbCrLf

		sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit4_1.asp?staffcode=" & vStaffCode & "&amp;flag=1&amp;categorycode=" & dbCategoryCode & "&amp;seq=" & dbSeq & """ method=""post"" style=""display:inline;"">"
		sHTML = sHTML & "<input type=""submit"" value=""編　集"">"
		sHTML = sHTML & "</form>" & vbCrLf

		sHTML = sHTML & "<form action=""" & HTTPS_CURRENTURL & "staff/edit/edit4_1.asp?staffcode=" & vStaffCode & "&amp;flag=0&amp;categorycode=" & dbCategoryCode & "&amp;seq=" & dbSeq & """ method=""post"" style=""display:inline;"" onsubmit=""return confirm('「" & dbSkillName & "」を削除しますか？');"">"
		sHTML = sHTML & "<input type=""submit"" value=""削　除"">"
		sHTML = sHTML & "</form>" & vbCrLf

		sHTML = sHTML & "</div>" & vbCrLf
		sHTML = sHTML & "<div style=""clear:both;""></div>" & vbCrLf
		sHTML = sHTML & "</div>"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	GetHtmlSkillList = sHTML
End Function
%>
