<%
'******************************************************************************
'概　要：メールの署名を取得
'作　者：2007/06/25 Lis K.Kokubo
'引　数：vType				：メール送信先種別 ["1"]一般の求人広告 ["2"]リス案件 ["3"]求人票が紐付いていないリス案件
'　　　：vStaffCode			：メール送信元求職者コード
'　　　：vCompanyName		：メール送信先企業名
'　　　：vOrderCode			：メールに紐づいた情報コード
'　　　：vContactPersonName	：メール送信先案件担当者
'　　　：vSubject			：メール件名
'　　　：vBody				：メール内容
'　　　：vMailURL			：求職者情報参照先ＵＲＬ
'　　　：vHeader			：メール内容に追加するヘッダ
'　　　：vFooter			：メール内容に追加するフッタ
'戻り値：
'備　考：
'使用元：しごとナビ/staff/mailtocompany.asp
'更　新：
'******************************************************************************
Function GetMailBodyStaff(ByVal vType, ByVal vStaffCode, ByVal vCompanyCode, ByVal vCompanyName, ByVal vOrderCode, ByVal vContactPersonName, ByVal vSubject, ByVal vBody, ByVal vMailURL, ByVal vHeader, ByVal vFooter)
	Dim sBody
	Dim iLen

	sBody = ""
	Select Case vType
		Case "1":
			'一般の求人広告
			'企業名＋求人票担当者名＋様
			sBody = vCompanyname & "　" & vContactPersonName & "様" & vbCrLf & vbCrLf
		Case "2":
			'リス案件
			'リス社員名＋様
			sBody = vContactPersonName & "様" & vbCrLf & vbCrLf
			vMailURL = "http://bi.lis21.co.jp/staff/staff_MailHistory.asp?Mail=pop&newopen=1"
		Case Else:
			'求人票が紐付いていない
			If Left(vCompanyCode, 1) = "L" Then
				vMailURL = "http://bi.lis21.co.jp/staff/staff_MailHistory.asp?Mail=pop&newopen=1"
			End If
	End Select
	sBody = sBody & vHeader & vbCrLf & vMailURL & vbCrLf

	'メール内容表示処理
	iLen = Len(vBody) * 0.3
	sBody = sBody & vbCrLf & vbCrLf & _
		"-----------------------■　メール情報　■-------------------------" & vbCrLf & _
		"【送信者コード】" & vStaffCode & vbCrLf & _
		"【対象情報コード】" & vOrderCode & vbCrLf & _
		"【メールタイトル】" & vSubject & vbCrLf & _
		"【メール内容】" & vbCrLf & vbCrLf & Left(vBody, iLen) & "..." & vbCrLf & _
		"------------------------------------------------------------------" & vbCrLf
	sBody = sBody & vFooter

	GetMailBodyStaff = sBody
End Function

'******************************************************************************
'概　要：メールの署名を取得
'作　者：2007/06/25 Lis K.Kokubo
'引　数：rDB		：接続中ＤＢオブジェクト
'　　　：vUserID	：ログイン中のユーザ種類
'戻り値：
'備　考：
'使用元：しごとナビ/staff/mailtoperson.asp
'更　新：
'******************************************************************************
Function GetMailSignatureStaff(ByRef rDB, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	sSQL = "sp_GetDataMailSignatureStaff '" & vUserID & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	sSignature = ""
	If GetRSState(oRS) = True Then
		sSignature = sSignature & "------------------------------" & vbCrLf
		sSignature = sSignature & "住所：" & oRS.Collect("Prefecture") & oRS.Collect("City") & oRS.Collect("Town") & oRS.Collect("Address") & vbCrLf
		sSignature = sSignature & "氏名：" & oRS.Collect("Name") & vbCrLf
		If oRS.Collect("HomeContactFlag") = "1" Then
			sSignature = sSignature & "自宅：" & oRS.Collect("HomeTelephoneNumber") & vbCrLf
		End If
		If oRS.Collect("PortableContactFlag") = "1" Then
			sSignature = sSignature & "携帯：" & oRS.Collect("PortableTelephoneNumber") & vbCrLf
		End If
		If oRS.Collect("FaxContactFlag") = "1" Then
			sSignature = sSignature & "FAX ：" & oRS.Collect("FaxNumber") & vbCrLf
		End If
		If oRS.Collect("MailContactFlag") = "1" Then
			sSignature = sSignature & "Mail：" & oRS.Collect("MailAddress") & vbCrLf
		End If
	End If
	Call RSClose(oRS)

	GetMailSignatureStaff = sSignature
End Function

'******************************************************************************
'概　要：求職者のメールテンプレート取得
'作　者：2007/06/25 Lis K.Kokubo
'引　数：rDB		：接続中ＤＢオブジェクト
'　　　：vUserType	：ログイン中のユーザ種類
'　　　：vSEQ		：番号
'戻り値：
'備　考：
'使用元：しごとナビ/staff/mailtocompany.asp
'更　新：
'******************************************************************************
Function GetStaffMailTemplateOptionHtml(ByRef rDB, ByVal vUserType, ByVal vSEQ)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sSelected

	GetStaffMailTemplateOptionHtml = ""

	sSQL = "sp_GetDataMailTemplate '1'"

	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sSelected = ""
		If CStr(oRS.Collect("Cd")) = CStr(vSEQ) Then sSelected = "selected=""true"""
		GetStaffMailTemplateOptionHtml = GetStaffMailTemplateOptionHtml & _
			"<option value=""" & oRS.Collect("Cd") & """ " & sSelected & ">" & oRS.Collect("Title") & "</option>"
		oRS.MoveNext
	Loop
	Call RSClose(oRS)
End Function
%>
