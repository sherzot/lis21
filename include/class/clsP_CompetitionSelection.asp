<%
'******************************************************************************
'名　称：clsP_CompetitionSelection
'概　要：formで飛んできたP_テーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_CompetitionSelection
	Public StaffCode
	Public IndustryTypeCode()
	Public JobTypeCode()
	Public CompanyName()
	Public SelectionTypeCode()
	Public MediaCode()
	Public OtherMedia()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_CompetitionSelection クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If Request.Form("StaffCode") <> "" Then StaffCode = Request.Form("StaffCode")

		Do While True
			If ExistsForm("CONF_SelectionTypeCode" & idx) = False Then Exit Do

			If Request.Form("CONF_SelectionTypeCode" & idx) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve IndustryTypeCode(MaxIndex) : IndustryTypeCode(MaxIndex) = Request.Form("CONF_IndustryTypeCode_S" & idx)
				ReDim Preserve JobTypeCode(MaxIndex) : JobTypeCode(MaxIndex) = Request.Form("CONF_JobTypeCode_S" & idx)
				ReDim Preserve CompanyName(MaxIndex) : CompanyName(MaxIndex) = Request.Form("CONF_CompanyName_S" & idx)
				ReDim Preserve SelectionTypeCode(MaxIndex) : SelectionTypeCode(MaxIndex) = Request.Form("CONF_SelectionTypeCode" & idx)
				ReDim Preserve MediaCode(MaxIndex) : MediaCode(MaxIndex) = Request.Form("CONF_MediaCode" & idx)
				ReDim Preserve OtherMedia(MaxIndex) : OtherMedia(MaxIndex) = Request.Form("CONF_OtherMedia" & idx)
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_CompetitionSelection 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_CompetitionSelection '" & ChkSQLStr(vStaffCode) & "'"
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
		GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_CompetitionSelection" & _
			" '" & ChkSQLStr(vStaffCode) & "'" & _
			",''" & _
			",'" & ChkSQLStr(IndustryTypeCode(idx)) & "'" & _
			",'" & ChkSQLStr(JobTypeCode(idx)) & "'" & _
			",'" & ChkSQLStr(CompanyName(idx)) & "'" & _
			",'" & ChkSQLStr(SelectionTypeCode(idx)) & "'" & _
			",'" & ChkSQLStr(MediaCode(idx)) & "'" & _
			",'" & ChkSQLStr(OtherMedia(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
