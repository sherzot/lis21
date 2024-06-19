<%
'*******************************************************************************
'概　要：都道府県一覧の<option></option>を取得
'引　数：vTime		：選択中の時刻
'　　　：vSepMinute	：分の区切り...15の場合00,15,30,45、30の場合00,30、60を割ったときに割り切れる数を入力
'戻り値：String
'備　考：
'履　歴：2009/08/31 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlTimeOption(ByVal vTime, ByVal vSepMinute)
	Dim sHTML
	Dim aMinute()
	Dim iSep
	Dim idx
	Dim idx2
	Dim sTime
	Dim sSelected

	If 60 Mod CInt(vSepMinute) <> 0 Then Exit Function

	iSep = 60 / CInt(vSepMinute)
	ReDim aMinute(iSep - 1)
	For idx = 0 To iSep - 1
		aMinute(idx) = Right("0" & CInt(vSepMinute) * idx, 2)
	Next

	sHTML = ""
	For idx = 0 To 23
		For idx2 = 0 To UBound(aMinute)
			sSelected = ""

			sTime = Right("0" & idx, 2) & aMinute(idx2)
			If sTime = vTime Then sSelected = " selected"

			sHTML = sHTML & "<option value=""" & sTime & """" & sSelected & ">" & Right("0" & idx, 2) & ":" & aMinute(idx2) & "</option>"
		Next
	Next

	htmlTimeOption = sHTML
End Function
%>
