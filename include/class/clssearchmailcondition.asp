<%
'******************************************************************************
'概　要：メール一覧の検索条件を保持するクラス
'関　数：■Private
'　　　：
'　　　：■Public
'　　　：Class_Initialize	：コンストラクタ
'　　　：GetSearchParam		：メール一覧のＧＥＴパラメータ取得
'　　　：GetSQLSearchMail	：メール一覧ＳＱＬ取得
'　　　：
'備　考：■■■ 検索用パラメータ (アドホックなＳＱＬ生成)
'　　　：DayFrom				：日付下限[YYYYMMDD]
'　　　：DayTo					：日付上限[YYYYMMDD]
'　　　：OrderCode				：情報コード
'　　　：SearchCode				：相手コード
'　　　：Evaluation				：評価
'　　　：Keyword				：キーワード
'　　　：MailContactPersonName	：求人担当者
'　　　：
'更　新：2007/12/25 LIS K.Kokubo 作成
'　　　：2009/03/27 LIS K.Kokubo 改修 MailHistory.RegistDay削除→SendDayで統一
'　　　：2009/07/30 LIS K.Kokubo スカウトアプローチフラグ検索追加
'　　　：2010/05/13 LIS K.Kokubo 未読メール検索追加(受信側のみ)
'******************************************************************************
Class clsSearchMailCondition
	Public UserCode

	'検索条件メンバ変数
	Public DayFrom
	Public DayTo
	Public OrderCode
	Public SearchCode
	Public Evaluation
	Public Keyword
	Public MailContactPersonName
	Public ScoutApproachFlag
	Public NotOpenFlag

	'その他メンバ変数
	Public HtmlStaffSearch	'検索条件出力ＨＴＭＬ文
	Public SQLStaffSearch	'検索ＳＱＬ
	Public SQLWriteLog		'ログ書き込みＳＱＬ

	'******************************************************************************
	'概　要：コンストラクタ
	'引　数：
	'備　考：
	'更　新：2007/12/26 LIS K.Kokubo 作成
	'******************************************************************************
	Private Sub Class_Initialize()
		UserCode = Session("userid")
		'パラメータから検索条件を取得
		If GetForm("sdf", 2) <> "" Then DayFrom = GetForm("sdf", 2)
		If GetForm("sdt", 2) <> "" Then DayTo = GetForm("sdt", 2)
		If GetForm("soc", 2) <> "" Then OrderCode = GetForm("soc", 2)
		If GetForm("sc", 2) <> "" Then SearchCode = GetForm("sc", 2)
		If GetForm("se", 2) <> "" Then Evaluation = GetForm("se", 2)
		If GetForm("skwd", 2) <> "" Then Keyword = GetForm("skwd", 2)
		If GetForm("mcpn", 2) <> "" Then MailContactPersonName = GetForm("mcpn", 2)
		If GetForm("ssaf", 2) <> "" Then ScoutApproachFlag = GetForm("ssaf", 2)
		If GetForm("snof", 2) <> "" Then NotOpenFlag = GetForm("snof", 2)
	End Sub

	'******************************************************************************
	'概　要：メール一覧のGETパラメータを生成して取得。
	'備　考：■制限
	'　　　：パラメータを含むURLは、IEの制限が2048文字までであるので、それに合わせる。
	'更　新：2007/12/26 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function GetSearchParam()
		GetSearchParam = ""
		If DayFrom <> "" Then GetSearchParam = GetSearchParam & "&amp;sdf=" & DayFrom
		If DayTo <> "" Then GetSearchParam = GetSearchParam & "&amp;sdt=" & DayTo
		If OrderCode <> "" Then GetSearchParam = GetSearchParam & "&amp;soc=" & OrderCode
		If SearchCode <> "" Then GetSearchParam = GetSearchParam & "&amp;sc=" & SearchCode
		If Evaluation <> "" Then GetSearchParam = GetSearchParam & "&amp;se=" & Evaluation
		If Keyword <> "" Then GetSearchParam = GetSearchParam & "&amp;skwd=" & Server.URLEncode(Keyword)
		If MailContactPersonName <> "" Then GetSearchParam = GetSearchParam & "&amp;mcpn=" & Server.URLEncode(MailContactPersonName)
		If ScoutApproachFlag <> "" Then GetSearchParam = GetSearchParam & "&amp;ssaf=" & ScoutApproachFlag
		If NotOpenFlag <> "" Then GetSearchParam = GetSearchParam & "&amp;snof=" & NotOpenFlag

		If GetSearchParam <> "" Then
			'頭の&amp;を削除
			GetSearchParam = Mid(GetSearchParam, 6)

			'ＩＥの仕様はパラメータの上限が２０４８バイト
			GetSearchParam = Left(GetSearchParam, 2048)
		End If
	End Function

	'******************************************************************************
	'概　要：メール一覧ＳＱＬ取得
	'引　数：vMode	：送受信フラグ	["1"]送信モード [<>"1"]受信モード
	'備　考：
	'更　新：2007/12/25 LIS K.Kokubo 作成
	'******************************************************************************
	Public Function GetSQLSearchMail(ByVal vMode)
		Dim sSQL
		Dim tmpSQL1
		Dim tmpSQL2
		Dim sDeclare
		Dim sParams
		Dim sWhere
		Dim sJoin
		Dim idx

		sDeclare = ""
		sParams = ""
		sWhere = ""
		sJoin = ""

		tmpSQL1 = ""
		tmpSQL2 = ""

		'ログイン中ユーザ
		sDeclare = sDeclare & "@vUserCode VARCHAR(8)"
		sParams = sParams & ",@vUserCode = N'" & UserCode & "'"

		If vMode = "1" Then
			'送信メール一覧

			'期間指定
			tmpSQL1 = ""
			If DayFrom & DayTo <> "" Then
				If DayFrom <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDayFrom VARCHAR(8)"
					sParams = sParams & ",@vDayFrom = N'" & DayFrom & "'"

					If tmpSQL1 <> "" Then tmpSQL1 = tmpSQL1 & "AND "
					tmpSQL1 = tmpSQL1 & "A.SendDay >= CONVERT(DATETIME, @vDayFrom) "
				End If

				If DayTo <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDayTo VARCHAR(8)"
					sParams = sParams & ",@vDayTo = N'" & DayTo & "'"

					If tmpSQL1 <> "" Then tmpSQL1 = tmpSQL1 & "AND "
					tmpSQL1 = tmpSQL1 & "A.SendDay < DATEADD(DAY, 1, CONVERT(DATETIME, @vDayTo)) "
				End If

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.SenderCode = @vUserCode AND " & tmpSQL1 & ") AS MDAY ON MH.ID = MDAY.ID "
			End If

			'情報コード
			If OrderCode <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOrderCode VARCHAR(8)"
				sParams = sParams & ",@vOrderCode = N'" & OrderCode & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.SenderCode = @vUserCode AND A.OrderCode = @vOrderCode) AS MORD ON MH.ID = MORD.ID "
			End If

			'相手コード
			If SearchCode <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSearchCode VARCHAR(8)"
				sParams = sParams & ",@vSearchCode = N'" & SearchCode & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.SenderCode = @vUserCode AND A.ReceiverCode = @vSearchCode) AS MSCD ON MH.ID = MSCD.ID "
			End If

			'評価
			If Evaluation <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vEvaluation VARCHAR(8)"
				sParams = sParams & ",@vEvaluation = N'" & Evaluation & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.SenderCode = @vUserCode AND A.SenderEvaluation = @vEvaluation) AS MEVL ON MH.ID = MEVL.ID "
			End If

			'キーワード
			If Keyword <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vKeyword VARCHAR(100)"
				sParams = sParams & ",@vKeyword = N'" & Keyword & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.SenderCode = @vUserCode AND A.Subject + ':' + ISNULL(A.SenderRemark, '') + ':' + A.Body LIKE '%' + @vKeyword + '%') AS MWRD ON MH.ID = MWRD.ID "
			End If

			'求人担当者
			If MailContactPersonName <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vMailContactPersonName VARCHAR(100)"
				sParams = sParams & ",@vMailContactPersonName = N'" & MailContactPersonName & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A INNER JOIN C_Contact AS B ON A.OrderCode = B.OrderCode AND B.PersonName = @vMailContactPersonName WHERE A.SenderCode = @vUserCode) AS MCPN ON MH.ID = MCPN.ID "
			End If

			'スカウトアプローチフラグ
			If ScoutApproachFlag <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vScoutApproachFlag VARCHAR(1)"
				sParams = sParams & ",@vScoutApproachFlag = N'" & ScoutApproachFlag & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.SenderCode = @vUserCode AND A.ScoutApproachFlag = @vScoutApproachFlag) AS SAF ON MH.ID = SAF.ID "
			End If

			sSQL = ""
			sSQL = sSQL & "SELECT MH.ID, MH.SendDay "
			sSQL = sSQL & "FROM MailHistory AS MH " & sJoin
			sSQL = sSQL & "WHERE MH.SenderCode = @vUserCode "
			sSQL = sSQL & "AND MH.SenderDelFlag = '0' "
			sSQL = sSQL & "OPTION(MAXDOP 1) "

			'パラメータクエリ化
			sSQL = "" & _
			"/*ナビ・送信メール一覧*/ " & vbCrLf & _
			"SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " & vbCrLf & _
			"EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
			If sDeclare <> "" Then sSQL = sSQL & ",N'" & sDeclare & "'" & sParams
		Else
			'受信メール一覧

			'期間指定
			tmpSQL1 = ""
			If DayFrom & DayTo <> "" Then
				If DayFrom <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDayFrom VARCHAR(8)"
					sParams = sParams & ",@vDayFrom = N'" & DayFrom & "'"

					If tmpSQL1 <> "" Then tmpSQL1 = tmpSQL1 & "AND "
					tmpSQL1 = tmpSQL1 & "A.SendDay >= CONVERT(DATETIME, @vDayFrom) "
				End If

				If DayTo <> "" Then
					If sDeclare <> "" Then sDeclare = sDeclare & ","
					sDeclare = sDeclare & "@vDayTo VARCHAR(8)"
					sParams = sParams & ",@vDayTo = N'" & DayTo & "'"

					If tmpSQL1 <> "" Then tmpSQL1 = tmpSQL1 & "AND "
					tmpSQL1 = tmpSQL1 & "A.SendDay < DATEADD(DAY, 1, CONVERT(DATETIME, @vDayTo)) "
				End If

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND " & tmpSQL1 & ") AS MDAY ON MH.ID = MDAY.ID "
			End If

			'情報コード
			If OrderCode <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vOrderCode VARCHAR(8)"
				sParams = sParams & ",@vOrderCode = N'" & OrderCode & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.OrderCode = @vOrderCode) AS MORD ON MH.ID = MORD.ID "
			End If

			'相手コード
			If SearchCode <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vSearchCode VARCHAR(8)"
				sParams = sParams & ",@vSearchCode = N'" & SearchCode & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.SenderCode = @vSearchCode) AS MSCD ON MH.ID = MSCD.ID "
			End If

			'評価
			If Evaluation <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vEvaluation VARCHAR(8)"
				sParams = sParams & ",@vEvaluation = N'" & Evaluation & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.ReceiverEvaluation = @vEvaluation) AS MEVL ON MH.ID = MEVL.ID "
			End If

			'キーワード
			If Keyword <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vKeyword VARCHAR(100)"
				sParams = sParams & ",@vKeyword = N'" & Keyword & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.Subject + ':' + ISNULL(A.ReceiverRemark, '') + ':' + A.Body LIKE '%' + @vKeyword + '%') AS MWRD ON MH.ID = MWRD.ID "
			End If

			'求人担当者
			If MailContactPersonName <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vMailContactPersonName VARCHAR(100)"
				sParams = sParams & ",@vMailContactPersonName = N'" & MailContactPersonName & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A INNER JOIN C_Contact AS B ON A.OrderCode = B.OrderCode AND B.PersonName = @vMailContactPersonName WHERE A.ReceiverCode = @vUserCode) AS MCPN ON MH.ID = MCPN.ID "
			End If

			'スカウトアプローチフラグ
			If ScoutApproachFlag <> "" Then
				If sDeclare <> "" Then sDeclare = sDeclare & ","
				sDeclare = sDeclare & "@vScoutApproachFlag VARCHAR(1)"
				sParams = sParams & ",@vScoutApproachFlag = N'" & ScoutApproachFlag & "'"

				sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.ScoutApproachFlag = @vScoutApproachFlag) AS SAF ON MH.ID = SAF.ID "
			End If

			'未読メール
			If NotOpenFlag <> "" Then
				If NotOpenFlag = "1" Then
					'未読
					sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.OpenDay IS NULL) AS NRF ON MH.ID = NRF.ID "
				ElseIf NotOpenFlag = "0" Then
					'既読
					sJoin = sJoin & "INNER JOIN (SELECT A.ID FROM MailHistory AS A WHERE A.ReceiverCode = @vUserCode AND A.OpenDay > 0) AS NRF ON MH.ID = NRF.ID "
				End If
			End If

			sSQL = ""
			sSQL = sSQL & "SELECT MH.ID, MH.SendDay "
			sSQL = sSQL & "FROM MailHistory AS MH " & sJoin
			sSQL = sSQL & "WHERE MH.ReceiverCode = @vUserCode "
			sSQL = sSQL & "AND MH.ReceiverDelFlag = '0' "
			sSQL = sSQL & "OPTION(MAXDOP 1) "

			'パラメータクエリ化
			sSQL = "" & _
			"/*ナビ・受信メール一覧*/ " & vbCrLf & _
			"SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " & vbCrLf & _
			"EXEC sp_executesql N'" & Replace(sSQL, "'", "''") & "'"
			If sDeclare <> "" Then sSQL = sSQL & ",N'" & sDeclare & "'" & sParams
		End If

		GetSQLSearchMail = sSQL

'Response.Write GetSQLSearchMail
	End Function
End Class
%>
