<%
'*******************************************************************************
'T@vFEÒªÇÁo^µ½dvi¼ðæ¾
'ø@F
'ßèlFString
'õ@lF
'@ðF2010/08/25 LIS K.Kokubo ì¬
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
			If rAfter("LicenseName" & idx) = "pîñZpÒ" And UBound(Filter(tmpAry,"pîñZpÒ")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "pîñZpÒ"
			ElseIf rAfter("LicenseName" & idx) = "hsXgeWXg" And UBound(Filter(tmpAry,"hsXgeWXg")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "hsXgeWXg"
			ElseIf rAfter("LicenseName" & idx) = "vWFNg}l[W" And UBound(Filter(tmpAry,"vWFNg}l[W")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "vWFNg}l[W"
			ElseIf rAfter("LicenseName" & idx) = "VXeA[LeNg" And UBound(Filter(tmpAry,"VXeA[LeNg")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "VXeA[LeNg"
			ElseIf rAfter("LicenseName" & idx) = "hsT[rX}l[W" And UBound(Filter(tmpAry,"hsT[rX}l[W")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "hsT[rX}l[W"
			ElseIf rAfter("LicenseName" & idx) = "îñZLeBXyVXg" And UBound(Filter(tmpAry,"îñZLeBXyVXg")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "îñZLeBXyVXg"
			ElseIf rAfter("LicenseName" & idx) = "CCNP" And UBound(Filter(tmpAry,"CCNP")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "CCNP"
			ElseIf rAfter("LicenseName" & idx) = "LPIC Level 3" And UBound(Filter(tmpAry,"LPIC Level 3")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "LPIC Level 3"
			ElseIf rAfter("LicenseName" & idx) = "Åìt" And UBound(Filter(tmpAry,"Åìt")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "Åìt"
			ElseIf rAfter("LicenseName" & idx) = "PA}l[W[iîìxêåõj" And UBound(Filter(tmpAry,"PA}l[W[iîìxêåõj")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "PA}l[W[iîìxêåõj"
			ElseIf rAfter("LicenseName" & idx) = "òÜt" And UBound(Filter(tmpAry,"òÜt")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "òÜt"
			ElseIf rAfter("LicenseName" & idx) = "ãt" And UBound(Filter(tmpAry,"ãt")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "ãt"
			ElseIf rAfter("LicenseName" & idx) = "Õ°¸Zt" And UBound(Filter(tmpAry,"Õ°¸Zt")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "Õ°¸Zt"
			ElseIf rAfter("LicenseName" & idx) = "MRFèi" And UBound(Filter(tmpAry,"MRFèi")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "MRFèi"
			ElseIf rAfter("LicenseName" & idx) = "Ûçm" And UBound(Filter(tmpAry,"Ûçm")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "Ûçm"
			ElseIf rAfter("LicenseName" & idx) = "h{m" And UBound(Filter(tmpAry,"h{m")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "h{m"
			ElseIf rAfter("LicenseName" & idx) = "yÅìt" And UBound(Filter(tmpAry,"yÅìt")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "yÅìt"
			ElseIf rAfter("LicenseName" & idx) = "îìm" And UBound(Filter(tmpAry,"îìm")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "îìm"
			ElseIf rAfter("LicenseName" & idx) = "wÃ@m" And UBound(Filter(tmpAry,"wÃ@m")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "wÃ@m"
			ElseIf rAfter("LicenseName" & idx) = "ìÆÃ@m" And UBound(Filter(tmpAry,"ìÆÃ@m")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "ìÆÃ@m"
			ElseIf rAfter("LicenseName" & idx) = "Ût" And UBound(Filter(tmpAry,"Ût")) < 0 Then
				If sImportant <> "" Then sImportant = sImportant & ","
				sImportant = sImportant & "Ût"
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
