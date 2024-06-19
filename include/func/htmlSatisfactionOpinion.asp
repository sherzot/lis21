<%
'*******************************************************************************
'概　要：企業満足度アンケートコンテンツの「皆さんの意見」部分のHTMLを取得
'引　数：vQuestionnaireID	：企業満足度アンケートID
'戻り値：String
'備　考：
'更　新：2009/06/01 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlSatisfactionOpinion(ByVal vQuestionnaireID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	'DB
	Dim dbPenName
	Dim dbOpinion
	Dim dbRegistDay

	Dim sHTML

	sHTML = ""

	sSQL = ""
	sSQL = sSQL & "/* しごとナビ 企業満足度アンケート コメント一覧取得 */" & vbCrLf
	sSQL = sSQL & "EXEC up_LstCMPQuestionnaireOpinion '" & vQuestionnaireID & "','1','1';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sHTML = sHTML & "<div style=""margin:0 0 5px 15px; border-bottom:1px dashed #ccc;"">"
		sHTML = sHTML & "<div style=""padding:4px;font-size:16px; font-weight:bold;color:orange;"">皆さんからの意見</div>"

		Do While GetRSState(oRS) = True
			dbPenName = ChkStr(oRS.Collect("PenName"))
			dbOpinion = ChkStr(oRS.Collect("Opinion"))
			dbRegistDay = oRS.Collect("RegistDay")

			'コメント
			sHTML = sHTML & "<div style=""padding:4px;"">"
			sHTML = sHTML & "<p class=""m0"">ペンネーム：" & dbPenName & "</p>"
			sHTML = sHTML & "<div style=""margin-left:15px;"">"
			sHTML = sHTML & "<b>L</b><p class=""m0"" style=""float:right;width:675px;"">" & Replace(Trim(dbOpinion) & "(" & GetDateStr(dbRegistDay, "/") & ")", vbCrLf, "<br>") & "</p>"
			sHTML = sHTML & "</div>"
			sHTML = sHTML & "<div clear=""both""></div></div>"


			oRS.MoveNext
		Loop
		sHTML = sHTML & "</div>"
	End If
	Call RSClose(oRS)

	htmlSatisfactionOpinion = sHTML
End Function
%>
