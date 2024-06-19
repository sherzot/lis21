<%
'*******************************************************************************
'概　要：代理店拠点一覧の<option></option>を取得
'引　数：vAgencyCode：代理店コード
'　　　：vBranchSeq	：代理店拠点番号
'　　　：vCode		：チェック中担当者番号
'　　　：vAttribute	：optionの追加属性
'戻り値：String
'備　考：
'履　歴：2010/03/17 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlAGCChargeOption(ByVal vAgencyCode, ByVal vBranchSeq, ByVal vCode, ByVal vAttribute)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sSQLErr

	Dim dbChargeSeq
	Dim dbPersonName

	Dim sHTML
	Dim aCode
	Dim aFilter
	Dim sSelected

	sHTML = ""

	If vAttribute <> "" Then vAttribute = " " & vAttribute

	sSQL = ""
	sSQL = sSQL & "/* 代理店拠点一覧 */" & vbCrLf
	sSQL = sSQL & "SELECT ChargeSeq,PersonName FROM AGCCharge WHERE AgencyCode = '" & ChkStr(vAgencyCode) & "' AND BranchSeq = '" & ChkStr(vBranchSeq) & "';"
	flgQE = QUERYEXE(dbconn,oRS,sSQL,sSQLErr)
	aCode = Split(ChkStr(vCode), ",")
	Do While GetRSState(oRS) = True
		dbChargeSeq = oRS.Collect("ChargeSeq")
		dbPersonName = oRS.Collect("PersonName")

		sSelected = ""
		If UBound(Filter(aCode, dbChargeSeq)) >= 0 Then sSelected = " selected"

		sHTML = sHTML & "<option value=""" & dbChargeSeq & """" & vAttribute & sSelected & ">" & dbPersonName & "</option>"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	htmlAGCChargeOption = sHTML
End Function
%>
