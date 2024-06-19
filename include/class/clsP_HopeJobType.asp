<%
'******************************************************************************
'名　称：clsP_HopeJobType
'概　要：formで飛んできたP_HopeJobTypeテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_HopeJobType
	Public StaffCode
	Public JobTypeCode()
	Public JobTypeDetail()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_HopeJobType クラスの初期化関数
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

		Err = ""

		Do While True
			If ExistsForm("CONF_JobTypeCode" & idx) = False Then Exit Do

			If Request.Form("CONF_JobTypeCode" & idx) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve JobTypeCode(MaxIndex) : JobTypeCode(MaxIndex) = Request.Form("CONF_JobTypeCode" & idx)
				ReDim Preserve JobTypeDetail(MaxIndex) : JobTypeDetail(MaxIndex) = Request.Form("CONF_JobTypeDetail" & idx)

				If JobTypeCode(MaxIndex) <> "" And IsNumber(JobTypeCode(MaxIndex), 3, False) = False Then Err = Err & "JobTypeCode(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_HopeJobType 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_HopeJobType '" & ChkSQLStr(vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_HopeJobType" & _
				" '" & ChkSQLStr(vStaffCode )& "'"  & _
				",''" & _
				",'" & ChkSQLStr(JobTypeCode(idx)) & "'" & _
				",'" & ChkSQLStr(JobTypeDetail(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
