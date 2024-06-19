<%
'******************************************************************************
'名　称：clsP_CareerHistoryIT
'概　要：formで飛んできたP_CareerHistoryITテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/04/05
'更　新：
'******************************************************************************
Class clsP_CareerHistoryIT
	Public StaffCode
	Public StartDay()
	Public EndDay()
	Public Number()
	Public PMFlag()
	Public PLFlag()
	Public SEFlag()
	Public PGFlag()
	Public TSFlag()
	Public SystemAnalysisFlag()
	Public DesignFlag()
	Public DevelopmentFlag()
	Public TestFlag()
	Public MaintenanceFlag()
	Public DevelopmentRemark()
	Public DevelopmentDetail()
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：clsP_CareerHistoryIT クラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim sStartDay
		Dim sEndDay
		Dim idx	: idx = 1
		Dim flg	: flg = False

		IsData = False
		MaxIndex = -1
		If Request.Form("StaffCode") <> "" Then StaffCode = Request.Form("StaffCode")

		Err = ""

		Do While True
			If ExistsForm("CONF_DevelopmentDetail" & idx) = False Then Exit Do
			sStartDay = ""
			sEndDay = ""

			sStartDay = Request.Form("CONF_StartDayITY" & idx) & "/"
			If Len(Request.Form("CONF_StartDayITM" & idx)) = 1 Then sStartDay = sStartDay & "0"
			sStartDay = sStartDay & Request.Form("CONF_StartDayITM" & idx) & "/01"
			If IsDate(sStartDay) = False Then sStartDay = ""

			sEndDay = Request.Form("CONF_EndDayITY" & idx) & "/"
			If Len(Request.Form("CONF_EndDayITM" & idx)) = 1 Then sEndDay = sEndDay & "0"
			sEndDay = sEndDay & Request.Form("CONF_EndDayITM" & idx) & "/01"
			If IsDate(sEndDay) = False Then sEndDay = ""

			If IsDate(sStartDay) = True Then sStartDay = Replace(sStartDay, "/", "")
			If IsDate(sEndDay) = True Then sEndDay = Replace(sEndDay, "/", "")

			If Request.Form("CONF_DevelopmentDetail" & idx) <> "" Then
				IsData = True
				MaxIndex = MaxIndex + 1

				ReDim Preserve StartDay(MaxIndex) : StartDay(MaxIndex) = sStartDay
				ReDim Preserve EndDay(MaxIndex) : EndDay(MaxIndex) = sEndDay
				ReDim Preserve Number(MaxIndex) : Number(MaxIndex) = Request.Form("CONF_Number_IT" & idx)
				ReDim Preserve PMFlag(MaxIndex) : PMFlag(MaxIndex) = Request.Form("CONF_PMFlag" & idx)
				ReDim Preserve PLFlag(MaxIndex) : PLFlag(MaxIndex) = Request.Form("CONF_PLFlag" & idx)
				ReDim Preserve SEFlag(MaxIndex) : SEFlag(MaxIndex) = Request.Form("CONF_SEFlag" & idx)
				ReDim Preserve PGFlag(MaxIndex) : PGFlag(MaxIndex) = Request.Form("CONF_PGFlag" & idx)
				ReDim Preserve TSFlag(MaxIndex) : TSFlag(MaxIndex) = Request.Form("CONF_TSFlag" & idx)
				ReDim Preserve SystemAnalysisFlag(MaxIndex) : SystemAnalysisFlag(MaxIndex) = Request.Form("CONF_SystemAnalysisFlag" & idx)
				ReDim Preserve DesignFlag(MaxIndex) : DesignFlag(MaxIndex) = Request.Form("CONF_DesignFlag" & idx)
				ReDim Preserve DevelopmentFlag(MaxIndex) : DevelopmentFlag(MaxIndex) = Request.Form("CONF_DevelopmentFlag" & idx)
				ReDim Preserve TestFlag(MaxIndex) : TestFlag(MaxIndex) = Request.Form("CONF_TestFlag" & idx)
				ReDim Preserve MaintenanceFlag(MaxIndex) : MaintenanceFlag(MaxIndex) = Request.Form("CONF_MaintenanceFlag" & idx)
				ReDim Preserve DevelopmentRemark(MaxIndex) : DevelopmentRemark(MaxIndex) = Request.Form("CONF_DevelopmentRemark" & idx)
				ReDim Preserve DevelopmentDetail(MaxIndex) : DevelopmentDetail(MaxIndex) = Request.Form("CONF_DevelopmentDetail" & idx)

				If StartDay(MaxIndex) <> "" And IsDay(StartDay(MaxIndex)) = False Then Err = Err & "StartDay(" & MaxIndex & ")" & vbCrLf
				If EndDay(MaxIndex) <> "" And IsDay(EndDay(MaxIndex)) = False Then Err = Err & "EndDay(" & MaxIndex & ")" & vbCrLf
				If Number(MaxIndex) <> "" And IsNumber(Number(MaxIndex), 0, False) = False Then Err = Err & "Number(" & MaxIndex & ")" & vbCrLf
				If PMFlag(MaxIndex) <> "" And IsFlag(PMFlag(MaxIndex)) = False Then Err = Err & "PMFlag(" & MaxIndex & ")" & vbCrLf
				If PLFlag(MaxIndex) <> "" And IsFlag(PLFlag(MaxIndex)) = False Then Err = Err & "PLFlag(" & MaxIndex & ")" & vbCrLf
				If SEFlag(MaxIndex) <> "" And IsFlag(SEFlag(MaxIndex)) = False Then Err = Err & "SEFlag(" & MaxIndex & ")" & vbCrLf
				If PGFlag(MaxIndex) <> "" And IsFlag(PGFlag(MaxIndex)) = False Then Err = Err & "PGFlag(" & MaxIndex & ")" & vbCrLf
				If TSFlag(MaxIndex) <> "" And IsFlag(TSFlag(MaxIndex)) = False Then Err = Err & "TSFlag(" & MaxIndex & ")" & vbCrLf
				If SystemAnalysisFlag(MaxIndex) <> "" And IsFlag(SystemAnalysisFlag(MaxIndex)) = False Then Err = Err & "SystemAnalysisFlag(" & MaxIndex & ")" & vbCrLf
				If DesignFlag(MaxIndex) <> "" And IsFlag(DesignFlag(MaxIndex)) = False Then Err = Err & "DesignFlag(" & MaxIndex & ")" & vbCrLf
				If DevelopmentFlag(MaxIndex) <> "" And IsFlag(DevelopmentFlag(MaxIndex)) = False Then Err = Err & "DevelopmentFlag(" & MaxIndex & ")" & vbCrLf
				If TestFlag(MaxIndex) <> "" And IsFlag(TestFlag(MaxIndex)) = False Then Err = Err & "TestFlag(" & MaxIndex & ")" & vbCrLf
				If MaintenanceFlag(MaxIndex) <> "" And IsFlag(MaintenanceFlag(MaxIndex)) = False Then Err = Err & "MaintenanceFlag(" & MaxIndex & ")" & vbCrLf
			End If
			idx = idx + 1
		Loop
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：sp_Reg_P_CareerHistoryIT 実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/03/24
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vStaffCode)
		Dim idx

		GetRegSQL = "EXEC sp_Del_P_CareerHistoryIT '" & ChkSQLStr(vStaffCode) & "'" & vbCrLf
		If IsData = False Then Exit Function
		For idx = 0 To MaxIndex
			GetRegSQL = GetRegSQL & "EXEC sp_Reg_P_CareerHistoryIT" & _
				" '" & ChkSQLStr(vStaffCode) & "'" & _
				",''" & _
				",'" & ChkSQLStr(StartDay(idx)) & "'" & _
				",'" & ChkSQLStr(EndDay(idx)) & "'" & _
				",'" & ChkSQLStr(Number(idx)) & "'" & _
				",'" & ChkSQLStr(PMFlag(idx)) & "'" & _
				",'" & ChkSQLStr(PLFlag(idx)) & "'" & _
				",'" & ChkSQLStr(SEFlag(idx)) & "'" & _
				",'" & ChkSQLStr(PGFlag(idx)) & "'" & _
				",'" & ChkSQLStr(TSFlag(idx)) & "'" & _
				",'" & ChkSQLStr(SystemAnalysisFlag(idx)) & "'" & _
				",'" & ChkSQLStr(DesignFlag(idx)) & "'" & _
				",'" & ChkSQLStr(DevelopmentFlag(idx)) & "'" & _
				",'" & ChkSQLStr(TestFlag(idx)) & "'" & _
				",'" & ChkSQLStr(MaintenanceFlag(idx)) & "'" & _
				",'" & ChkSQLStr(DevelopmentRemark(idx)) & "'" & _
				",'" & ChkSQLStr(DevelopmentDetail(idx)) & "'" & vbCrLf
		Next
	End Function
End Class
%>
