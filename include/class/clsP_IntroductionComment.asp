<%
'******************************************************************************
'名　称：clsP_IntroductionComment
'概　要：formで飛んできたP_IntroductionCommentテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_IntroductionComment
	Public StaffCode
	Public Comment(16)
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_IntroductionComment クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim sidx
		Dim idx

		IsData = False
		MaxIndex = UBound(Comment)
		StaffCode = Request.Form("CONF_StaffCode")

		For idx = 1 To UBound(Comment)
			If idx <= 9 Then
				sidx = "00" & idx
			Else
				sidx = "0" & idx
			End If

			Comment(idx) = Request.Form("CONF_IntroductionComment" & sidx)
			If Comment(idx) <> "" Then IsData = True
		Next
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_IntroductionComment 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx
		Dim sidx

		If IsData = False Then Exit Function

		GetRegSQL = ""
		For idx = 1 To MaxIndex
			If Comment(idx) <> "" Then
				If idx <= 9 Then
					sidx = "00" & idx
				Else
					sidx = "0" & idx
				End If

				GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_IntroductionComment" & _
					" '" & ChkSQLStr(vStaffCode) & "'" & _
					",'IntroductionComment'" & _
					",'" & ChkSQLStr(sidx) & "'" & _
					",'" & ChkSQLStr(Comment(idx)) & "'" & vbCrLf
			End If
		Next
	End Function
End Class
%>
