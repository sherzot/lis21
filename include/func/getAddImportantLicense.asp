<%
'*******************************************************************************
'概　要：求職者が追加登録した重要資格名を取得
'引　数：
'戻り値：String
'備　考：
'履　歴：2010/08/25 LIS K.Kokubo 作成
'*******************************************************************************
Function getAddImportantLicense(ByRef rBefore,ByRef rAfter)
	Dim idx
	Dim tmpAry,sImportant

	getAddImportantLicense = ""
	tmpAry = rBefore.Items
	sImportant = ""

	idx = 1
	Do While rBefore.Exists("LicenseName"&idx) = True Or rAfter.Exists("LicenseName"&idx) = True
		If UBound(tmpAry) > 0 And rAfter("LicenseName" & idx) <> "" Then
			If rAfter("LicenseName" & idx) = "応用情報技術者" And UBound(Filter(tmpAry,"応用情報技術者")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "応用情報技術者"
			ElseIf rAfter("LicenseName" & idx) = "ＩＴストラテジスト" And UBound(Filter(tmpAry,"ＩＴストラテジスト")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "ＩＴストラテジスト"
			ElseIf rAfter("LicenseName" & idx) = "プロジェクトマネージャ" And UBound(Filter(tmpAry,"プロジェクトマネージャ")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "プロジェクトマネージャ"
			ElseIf rAfter("LicenseName" & idx) = "システムアーキテクト" And UBound(Filter(tmpAry,"システムアーキテクト")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "システムアーキテクト"
			ElseIf rAfter("LicenseName" & idx) = "ＩＴサービスマネージャ" And UBound(Filter(tmpAry,"ＩＴサービスマネージャ")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "ＩＴサービスマネージャ"
			ElseIf rAfter("LicenseName" & idx) = "情報セキュリティスペシャリスト" And UBound(Filter(tmpAry,"情報セキュリティスペシャリスト")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "情報セキュリティスペシャリスト"
			ElseIf rAfter("LicenseName" & idx) = "CCNP" And UBound(Filter(tmpAry,"CCNP")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "CCNP"
			ElseIf rAfter("LicenseName" & idx) = "LPIC Level 3" And UBound(Filter(tmpAry,"LPIC Level 3")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "LPIC Level 3"
			ElseIf rAfter("LicenseName" & idx) = "看護師" And UBound(Filter(tmpAry,"看護師")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "看護師"
			ElseIf rAfter("LicenseName" & idx) = "ケアマネージャー（介護支援専門員）" And UBound(Filter(tmpAry,"ケアマネージャー（介護支援専門員）")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "ケアマネージャー（介護支援専門員）"
			ElseIf rAfter("LicenseName" & idx) = "薬剤師" And UBound(Filter(tmpAry,"薬剤師")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "薬剤師"
			ElseIf rAfter("LicenseName" & idx) = "医師" And UBound(Filter(tmpAry,"医師")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "医師"
			ElseIf rAfter("LicenseName" & idx) = "臨床検査技師" And UBound(Filter(tmpAry,"臨床検査技師")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "臨床検査技師"
			ElseIf rAfter("LicenseName" & idx) = "MR認定資格" And UBound(Filter(tmpAry,"MR認定資格")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "MR認定資格"
			ElseIf rAfter("LicenseName" & idx) = "保育士" And UBound(Filter(tmpAry,"保育士")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "保育士"
			ElseIf rAfter("LicenseName" & idx) = "栄養士" And UBound(Filter(tmpAry,"栄養士")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "栄養士"
			ElseIf rAfter("LicenseName" & idx) = "准看護師" And UBound(Filter(tmpAry,"准看護師")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "准看護師"
			ElseIf rAfter("LicenseName" & idx) = "介護福祉士" And UBound(Filter(tmpAry,"介護福祉士")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "介護福祉士"
			ElseIf rAfter("LicenseName" & idx) = "理学療法士" And UBound(Filter(tmpAry,"理学療法士")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "理学療法士"
			ElseIf rAfter("LicenseName" & idx) = "作業療法士" And UBound(Filter(tmpAry,"作業療法士")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "作業療法士"
			ElseIf rAfter("LicenseName" & idx) = "保健師" And UBound(Filter(tmpAry,"保健師")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "保健師"
			End If
		End If

		idx = idx + 1
	Loop

	getAddImportantLicense = sImportant
End Function

Function setDicLicense(ByRef rDB,ByVal vUserCode,ByRef rDic)
	Dim sSQL,oRS,flgQE,sSQLErr
	Dim idx

	setDicLicense = False
	Set rDic = Server.CreateObject("scripting.dictionary")

	sSQL = "sp_GetDataLicense '" & vUserCode & "'"
	flgQE = QUERYEXE(rDB,oRS,sSQL,sSQLErr)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing
		idx = 1
		Do While GetRSState(oRS)
			Call rDic.Add("Code" & idx, oRS.Collect("GroupCode") & oRS.Collect("CategoryCode") & oRS.Collect("Code"))
			Call rDic.Add("LicenseName" & idx, oRS.Collect("LicenseName"))
			Call rDic.Add("LicenseNameDsp" & idx, oRS.Collect("LicenseNameDsp"))
			Call rDic.Add("GetDay" & idx, ChkStr(oRS.Collect("GetDay")))

			idx = idx + 1
			oRS.MoveNext
		Loop
		setDicLicense = True
	End If
	Call RSClose(oRS)
End Function
%>
