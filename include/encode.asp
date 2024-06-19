<%
'*******************************************************************************
'概　要	：検索サイトからのキーワードを読める形に変換する関数群です。
'*******************************************************************************
Function GetYahooKeyWord()
'*******************************************************************************
'概　要	：
'引　数	：
'戻り値	：
'作　成	：2005/12/07(企画推進室　小久保)
'更　新	：
'*******************************************************************************
	On Error Resume Next
	Call Err.Clear()

	Dim sKey

	GetYahooKeyWord = ""

	If InStr(Request.ServerVariables("HTTP_REFERER"), "?") = 0 Then Exit Function
	If InStr(Request.ServerVariables("HTTP_REFERER"), "search.yahoo.co.jp") = 0 Then Exit Function
	'If InStr(Request.ServerVariables("HTTP_REFERER"), "ns1.shigotonavi.co.jp") = 0 Then Exit Function
	If InStr(Request.ServerVariables("HTTP_REFERER"), "OVKEY=") <> 0 Then Exit Function
	If InStr(Request.ServerVariables("HTTP_REFERER"), "p=") = 0 Then Exit Function
	If InStr(Request.ServerVariables("HTTP_REFERER"), "p=&") <> 0 Then Exit Function
	If Right(Request.ServerVariables("HTTP_REFERER"), 2) = "p=" Then Exit Function

	'▼▼▼▼▼ UTF-8 対応 ▼▼▼▼▼
	'履歴書
	If InStr(Request.ServerVariables("HTTP_REFERER"), "%E5%B1%A5%E6%AD%B4%E6%9B%B8") <> 0 _
	And InStr(Request.ServerVariables("HTTP_REFERER"), "UTF-8") <> 0 Then
		GetYahooKeyWord = "履歴書"
		Exit Function
	End If

	'職務経歴書
	If InStr(Request.ServerVariables("HTTP_REFERER"), "%E8%81%B7%E5%8B%99%E7%B5%8C%E6%AD%B4%E6%9B%B8") <> 0 _
	And InStr(Request.ServerVariables("HTTP_REFERER"), "UTF-8") <> 0 Then
		GetYahooKeyWord = "職務経歴書"
		Exit Function
	End If

	'志望動機
	If InStr(Request.ServerVariables("HTTP_REFERER"), "%E5%BF%97%E6%9C%9B%E5%8B%95%E6%A9%9F") <> 0 _
	And InStr(Request.ServerVariables("HTTP_REFERER"), "UTF-8") <> 0 Then
		GetYahooKeyWord = "志望動機"
		Exit Function
	End If

	'自己ＰＲ
	'　EUC  ：%BC%AB%B8%CA%A3%D0%A3%D2
	'　UTF-8：%E8%87%AA%E5%B7%B1%EF%BC%B0%EF%BC%B2
	'自己ｐｒ
	'　EUC  ：%BC%AB%B8%CA%A3%F0%A3%F2
	'　UTF-8：%E8%87%AA%E5%B7%B1%EF%BD%90%EF%BD%92
	'自己
	'　EUC  ：%BC%AB%B8%CAPR
	'　UTF-8：%E8%87%AA%E5%B7%B1
	If InStr(Request.ServerVariables("HTTP_REFERER"), "%E8%87%AA%E5%B7%B1") <> 0 _
	And InStr(Request.ServerVariables("HTTP_REFERER"), "UTF-8") <> 0 Then
		GetYahooKeyWord = "自己"
		Exit Function
	End If

	'退職願
	If InStr(Request.ServerVariables("HTTP_REFERER"), "%E9%80%80%E8%81%B7%E9%A1%98") <> 0 _
	And InStr(Request.ServerVariables("HTTP_REFERER"), "UTF-8") <> 0 Then
		GetYahooKeyWord = "退職願"
		Exit Function
	End If
	'▲▲▲▲▲ UTF-8 対応 ▲▲▲▲▲

	sKey = Mid(Request.ServerVariables("HTTP_REFERER"), InStr(Request.ServerVariables("HTTP_REFERER"), "p=") + 2)
	If InStr(sKey, "&") <> 0 Then
		sKey = Mid(sKey, 1, InStr(sKey, "&") - 1)
	End If

	If InStr(sKey, "%") = 0 Then
		Exit Function
	End If

	If Err.Number = 0 Then
		GetYahooKeyWord = EncEUC_SJIS(sKey)
	End If
End Function

Function EncEUC_SJIS(ByVal vsEUC)
'*******************************************************************************
'概　要	：
'引　数	：
'戻り値	：
'作　成	：2005/08/29(企画推進室　小久保)
'更　新	：
'*******************************************************************************
	Dim sText
	sText = EncEUC_JIS(vsEUC)
	sText = EncJIS_SJIS(sText)
	EncEUC_SJIS = sText
End Function

Function EncEUC_JIS(ByVal vsEUC)
'*******************************************************************************
'概　要	：
'引　数	：
'戻り値	：
'作　成	：2005/08/29(企画推進室　小久保)
'更　新	：
'*******************************************************************************
	Dim sJIS
	Dim sJISs
	Dim idx
	Dim bFlg
	Dim iCode
	Dim sEUC
	Dim bCharAsc

	If Left(vsEUC, 1) <> "%" Then
		If InStr(vsEUC, "%") > 0 Then
			sJIS = "%1B" & "%28" & "%42" & Left(vsEUC, InStr(vsEUC, "%") - 1)
			sEUC = Mid(vsEUC, InStr(vsEUC, "%") + 1)
		Else
			sJIS = Trim(vsEUC)
			Exit Function
		End If
		sJISs = Split(vsEUC, "%")
		bCharAsc = 1
	Else
		sJISs = Split(Mid(vsEUC, 2), "%")
		bCharAsc = 0
	End If

	sJIS = ""
	For idx = LBound(sJISs) To UBound(sJISs)
		If bCharAsc = 0 Then
			iCode = Enc16To10(Left(sJISs(idx), 2))

			If iCode >= Enc16To10("80") Then
				If bFlg = False Then
					sJIS = sJIS & "%1B" & "%24" & "%42"
					bFlg = True
				End If
				sJIS = sJIS & "%" & Enc10To16(iCode - Enc16To10("80"))
			Else
				If bFlg = True Then
					sJIS = sJIS & "%1B" & "%28" & "%42"
					bFlg = False
				End If
				sJIS = sJIS & "%" & sJISs(idx)
			End If
			If Len(sJISs(idx)) > 2 Then
				If bFlg = True Then
					sJIS = sJIS & "%1B" & "%28" & "%42" & Mid(sJISs(idx), 3)
				End If
				bFlg = False
			End If
		Else
			sJIS = sJIS & "%1B" & "%28" & "%42" & sJISs(idx)
			bFlg = False
			bCharAsc = 0
		End If
	Next

	EncEUC_JIS = sJIS
End Function

Function EncJIS_SJIS(ByVal vsJIS)
'*******************************************************************************
'概　要	：
'引　数	：
'戻り値	：
'作　成	：2005/08/29(企画推進室　小久保)
'更　新	：
'*******************************************************************************
	Dim sJIS
	Dim sSJIS
	Dim sSJIS2
	Dim sSJISs
	Dim idx
	Dim iFlg
	Dim iCode1
	Dim iCode2
	Dim sRest

	sJIS = Replace(LCase(vsJIS), "%1b%24%40", "%２文字")
	sJIS = Replace(LCase(sJIS), "%1b%24%42", "%２文字")
	sJIS = Replace(LCase(sJIS), "%1b%26%40%1b%21%42", "２文字")
	sJIS = Replace(LCase(sJIS), "%1b%24%28%44", "%２文字")
	sJIS = Replace(LCase(sJIS), "%1b%28%4A", "%１文字")
	sJIS = Replace(LCase(sJIS), "%1b%28%48", "%１文字")
	sJIS = Replace(LCase(sJIS), "%1b%28%42", "%１文字")
	sJIS = Replace(LCase(sJIS), "%1b%28%49", "%１文字")

	sSJISs = Split(Mid(sJIS, 2), "%")

	sSJIS = ""
	sSJIS2 = ""

	For idx = LBound(sSJISs) To UBound(sSJISs)
		sRest = ""

		If InStr(sSJISs(idx), "２文字") > 0 Then
			iFlg = 2
			If Len(sSJISs(idx)) > 3 Then
				sRest = Mid(sSJISs(idx), 4)
			End If
			idx = idx + 1
		ElseIf InStr(sSJISs(idx), "１文字") Then
			iFlg = 1
			If Len(sSJISs(idx)) > 3 Then
				sSJIS = sSJIS & Mid(sSJISs(idx), 4)
				sSJIS2 = sSJIS2 & Mid(sSJISs(idx), 4)
			End If

			idx = idx + 1

			If idx <= UBound(sSJISs) Then
				If InStr(sSJISs(idx), "２文字") > 0 Then
					iFlg = 2
					If Len(sSJISs(idx)) > 3 Then
						sRest = Mid(sSJISs(idx), 4)
					End If
					idx = idx + 1
				Else
					EncJIS_SJIS = sSJIS
					Exit Function
				End If
			Else
				EncJIS_SJIS = sSJIS
				Exit Function
			End If
		End If

		iCode1 = Enc16To10(Left(sSJISs(idx), 2))
		If Len(sSJISs(idx)) > 2 Then
			sRest = Mid(sSJISs, 3)
		End If

		If idx + 1 <= UBound(sSJISs) Then
			iCode2 = Enc16To10(Left(sSJISs(idx + 1), 2))
			If Len(sSJISs(idx + 1)) > 2 Then
				sRest = Mid(sSJISs(idx + 1), 3)
			End If
		Else
			EncJIS_SJIS = sSJIS2
			Exit Function
		End If

		If iFlg = 2 Then
			If iCode1 Mod 2 <> 0 Then
				iCode1 = (iCode1 + 1) / 2 + Enc16To10("70")
				iCode2 = iCode2 + Enc16To10("1F")
			Else
				iCode1 = iCode1 / 2 + Enc16To10("70")
				iCode2 = iCode2 + Enc16To10("7D")
			End If

			If iCode1 >= Enc16To10("A0") Then
				iCode1 = iCode1 + Enc16To10("40")
			End If
			If iCode2 >= Enc16To10("7F") Then
				iCode2 = iCode2 + 1
			End If
			sSJIS = sSJIS & Chr(Enc16To10(Enc10To16(iCode1) & Enc10To16(iCode2)))
			sSJIS2 = sSJIS2 & "%" & Enc10To16(iCode1) & "%" & Enc10To16(iCode2)
			idx = idx + 1
		ElseIf iFlg = 1 Then
			sSJIS = sSJIS & sRest
			sSJIS2 = sSJIS2 & sRest
		End If
	Next

	EncJIS_SJIS = sSJIS
End Function

Function EncSJIS_JIS(ByVal vsSJIS)
'*******************************************************************************
'概　要	：
'引　数	：
'戻り値	：
'作　成	：2005/08/29(企画推進室　小久保)
'更　新	：
'*******************************************************************************
	Dim sSJIS
	Dim sJIS
	Dim sJISs
	Dim idx
	Dim bFlg

	sSJIS = Mid(vsSJIS, 2)
	sJISs = Split(sSJIS, "%")

	sJIS = ""
	bFlg = False
	For idx = LBound(sJISs) To UBound(sJISs)
		If Enc16To10(sJISs(idx)) >= Enc16To10("E0") Then
			sJISs(idx) = Enc10To16(Enc16To10(sJISs(idx)) - Enc16To10("40"))
			bFlg = True
		End If

		If idx + 1 <= UBound(sJISs) Then
			If Enc16To10(sJISs(idx + 1)) >= Enc16To10("80") Then
				sJISs(idx + 1) = Enc10To16(Enc16To10(sJISs(idx + 1)) - 1)
				bFlg = True
			End If
			If Enc16To10(sJISs(idx + 1)) >= Enc16To10("9E") Then
				sJISs(idx) = Enc10To16((Enc16To10(sJISs(idx)) - Enc16To10("70")) * 2)
				sJISs(idx) = Enc10To16(Enc16To10(sJISs(idx + 1)) - Enc16To10("7D"))
			Else
				sJISs(idx) = Enc10To16((Enc16To10(sJISs(idx)) - Enc16To10("70")) * 2 - 1)
				sJISs(idx + 1) = Enc10To16((Enc16To10(sJISs(idx + 1)) - Enc16To10("1F")))
			End If
		End If

		If bFlg = True Then
			sJIS = sJIS & Chr(Enc16To10(sJISs(idx) & "00") + Enc16To10(sJISs(idx + 1)))
			idx = idx + 1
			bFlg = False
		Else
			sJIS = sJIS & Chr(Enc16To10(sJISs(idx)))
		End If
	Next

	EncSJIS_JIS = sJIS
End Function

Function Enc2To10(ByVal vs2txt)
'*******************************************************************************
'概　要	：2進数文字列から16進数文字列へ
'引　数	：
'戻り値	：
'作　成	：2005/08/29(企画推進室　小久保)
'更　新	：
'*******************************************************************************
	Dim idx
	Dim s2txt
	Dim iResult

	iResult = 0
	s2txt = StrReverse(vs2txt)
	For idx = 1 To Len(vs2txt)
		If idx = 1 Then
			iResult = iResult + CInt(Mid(s2txt, 1, 1))
		ElseIf Mid(s2txt, idx, 1) = "1" Then
			iResult = 2 ^ (idx - 1) + iResult
		End If
	Next

	Enc2To10 = iResult
End Function

Function Enc10To16(ByVal vi10)
'*******************************************************************************
'概　要	：整数から16進数文字列へ
'引　数	：
'戻り値	：
'作　成	：2005/08/29(企画推進室　小久保)
'更　新	：
'*******************************************************************************
	Dim sResult

	sResult = Enc10To2(vi10)
	sResult = Enc2To16(sResult)

	Enc10To16 = sResult
End Function

Function Enc10to2(ByVal vi10)
'*******************************************************************************
'概　要	：整数から2進数文字列へ
'引　数	：
'戻り値	：
'作　成	：2005/08/29(企画推進室　小久保)
'更　新	：
'*******************************************************************************
	Dim i10
	Dim s2txt

	i10 = vi10

	s2txt = ""
	Do While i10 > 1
		s2txt = (i10 Mod 2) & s2txt
		i10 = Int(i10 / 2)
	Loop
	s2txt = i10 & s2txt

	If Len(s2txt) Mod 8 <> 0 Then
		s2txt = String(8 - (Len(s2txt) Mod 8), "0") & s2txt
	End If

	Enc10to2 = s2txt
End Function

Function Enc16To10(ByVal vs16txt)
'*******************************************************************************
'概　要	：16進数文字列から整数へ
'引　数	：
'戻り値	：
'作　成	：2005/08/29(企画推進室　小久保)
'更　新	：
'*******************************************************************************
	Dim idx
	Dim iMaxIdx
	Dim sRev16txt
	Dim cOne
	Dim iNum

	iNum = 0
	iMaxIdx = Len(vs16txt)
	sRev16txt = StrReverse(LCase(vs16txt))

	For idx = 1 To iMaxIdx
		cOne = Mid(sRev16txt, idx, 1)
		if idx = 1 Then
			Select Case cOne
			Case "a":
				iNum = iNum + 10
			Case "b":
				iNum = iNum + 11
			Case "c":
				iNum = iNum + 12
			Case "d":
				iNum = iNum + 13
			Case "e":
				iNum = iNum + 14
			Case "f":
				iNum = iNum + 15
			Case Else:
				iNum = iNum + CInt(cOne)
			End Select
		Else
			Select Case cOne
			Case "a":
				iNum = 16 ^ (idx - 1) * 10 + iNum
			Case "b":
				iNum = 16 ^ (idx - 1) * 11 + iNum
			Case "c":
				iNum = 16 ^ (idx - 1) * 12 + iNum
			Case "d":
				iNum = 16 ^ (idx - 1) * 13 + iNum
			Case "e":
				iNum = 16 ^ (idx - 1) * 14 + iNum
			Case "f":
				iNum = 16 ^ (idx - 1) * 15 + iNum
			Case Else:
				iNum = 16 ^ (idx - 1) * CInt(cOne) + iNum
			End Select
		End If
	Next

	Enc16To10 = iNum
End Function

Function Enc16to2(ByVal vs16txt)
'*******************************************************************************
'概　要	：16進数文字列から2進数文字列へ
'引　数	：
'戻り値	：
'作　成	：2005/08/29(企画推進室　小久保)
'更　新	：
'*******************************************************************************
	Dim idx
	Dim s16txt
	Dim s2txt

	If Len(vs16txt) Mod 2 = 1 Then
		s16txt = StrReverse("0" & vs16txt)
	End If

	s2txt = ""
	For idx = 1 To Len(s16txt)
		Select Case Mid(s16txt, 1, 1)
		Case "0": s2txt = "0000" & s2txt
		Case "1": s2txt = "0001" & s2txt
		Case "2": s2txt = "0010" & s2txt
		Case "3": s2txt = "0011" & s2txt
		Case "4": s2txt = "0100" & s2txt
		Case "5": s2txt = "0101" & s2txt
		Case "6": s2txt = "0110" & s2txt
		Case "7": s2txt = "0111" & s2txt
		Case "8": s2txt = "1000" & s2txt
		Case "9": s2txt = "1001" & s2txt
		Case "a": s2txt = "1010" & s2txt
		Case "b": s2txt = "1011" & s2txt
		Case "c": s2txt = "1100" & s2txt
		Case "d": s2txt = "1101" & s2txt
		Case "e": s2txt = "1110" & s2txt
		Case "f": s2txt = "1111" & s2txt
		End Select
	Next

	Enc16to2 = s2txt
End Function

Function Enc2To16(ByVal vs2txt)
'*******************************************************************************
'概　要	：2進数文字列から16進数文字列へ
'引　数	：
'戻り値	：
'作　成	：2005/08/29(企画推進室　小久保)
'更　新	：
'*******************************************************************************
	Dim idx
	Dim s16txt

	If Len(vs2txt) Mod 8 <> 0 Then
		Enc2To16 = ""
		Exit Function
	End If

	s16txt = ""
	For idx = 1 To Len(vs2txt) Step 4
		Select Case Mid(vs2txt, idx, 4)
		Case "0000": s16txt = s16txt & "0"
		Case "0001": s16txt = s16txt & "1"
		Case "0010": s16txt = s16txt & "2"
		Case "0011": s16txt = s16txt & "3"
		Case "0100": s16txt = s16txt & "4"
		Case "0101": s16txt = s16txt & "5"
		Case "0110": s16txt = s16txt & "6"
		Case "0111": s16txt = s16txt & "7"
		Case "1000": s16txt = s16txt & "8"
		Case "1001": s16txt = s16txt & "9"
		Case "1010": s16txt = s16txt & "A"
		Case "1011": s16txt = s16txt & "B"
		Case "1100": s16txt = s16txt & "C"
		Case "1101": s16txt = s16txt & "D"
		Case "1110": s16txt = s16txt & "E"
		Case "1111": s16txt = s16txt & "F"
		End Select
	Next

	Enc2To16 = s16txt
End Function

Function BitCalculate(vsBits1, vsBits2, ByVal vsFlag, ByVal vsMode)
'*******************************************************************************
'概　要	：ビット演算
'引　数	：
'戻り値	：
'作　成	：2005/08/29(企画推進室　小久保)
'更　新	：
'*******************************************************************************
	Dim idx
	Dim sBits1
	Dim sBits2
	Dim sResult

	If vsFlag = "16" Then
		sBits1 = Enc16To2(vsBits1)
		sBits2 = Enc16To2(vsBits2)
	ElseIf vsFlag = "2" Then
		sBits1 = vsBits1
		sBits2 = vsBits2
	End If

	sResult = ""
	If vsMode = "and" Then
		For idx = 1 To Len(sBits1)
			If Mid(sBits1, idx, 1) = "1" And Mid(sBits2, idx, 1) = "1" Then
				sResult = sResult = "1"
			Else
				sResult = sResult & "0"
			End If
		Next
	ElseIf vsMode = "or" Then
		For idx = 1 To Len(sBits1)
			If Mid(sBits1, idx, 1) = "1" Or Mid(sBits2, idx, 1) = "1" Then
				sResult = sResult = "1"
			Else
				sResult = sResult & "0"
			End If
		Next
	End If

	If vsFlag = "16" Then
		sResult = Enc2To16(sResult)
	End If

	BitCalculate = sResult
End Function

Function GetSearchStr(vsURL)
'*******************************************************************************
'概　要	：
'引　数	：
'戻り値	：
'作　成	：
'更　新	：
'*******************************************************************************
	Const YAHOO = "http://search.yahoo.co.jp/bin/search?p="
	Const GOOGLE = "http://www.google.co.jp/search?hl=ja&q="

	Dim sURL

	sURL = ""
	If InStr(vsURL, YAHOO) > 0 Then
		'YAHOO
		sURL = Mid(vsURL, InStr(vsURL, YAHOO) + Len(YAHOO))
	ElseIf InStr(vsURL, GOOGLE) > 0 Then
		'GOOGLE
		sURL = Mid(vsURL, InStr(vsURL, GOOGLE) + Len(GOOGLE))
	End If

	If InStr(sURL, "&") > 0 Then
		sURL = Left(sURL, InStr(sURL, "&") - 1)
	End If
	GetSearchStr = sURL
End Function

Sub DebugWrite(ByVal vsFileName, ByVal vsText, ByVal vbMode, ByVal vbCrLf)
'*******************************************************************************
'概　要	：
'引　数	：
'戻り値	：
'作　成	：
'更　新	：
'*******************************************************************************
	Dim objFS
	Dim objTS

	Set objFS = CreateObject("scripting.filesystemobject")
	If vbMode = True Then
		Set objTS = objFS.OpenTextFile(vsFileName, 8, True)
	Else
		Set objTS = objFS.CreateTextFile(vsFileName, True)
	End If

	If vbCrLf = True Then
		Call objTS.WriteLine(vsText)
	Else
		Call objTS.Write(vsText)
	End If

	Call objTS.Close()

	Set objFS = Nothing
	Set objTS = Nothing
End Sub
%>
