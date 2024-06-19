<%
'******************************************************************************
'名　称：cls_Schedule
'概　要：formで飛んできたP_UserInfoテーブル用のデータを持つためのクラス
'備　考：
'作成者：Lis Kokubo
'作成日：2006/10/13
'更　新：
'******************************************************************************
Class cls_Schedule
	Public ScheduleID
	Public UserCode
	Public BriefingSessionID
	Public PublicType
	Public ScheduleTypeCode
	Public StartDay
	Public EndDay
	Public Subject
	Public Body
	Public PlaceName
	Public Longitude
	Public Latitude
	Public IsData
	Public MaxIndex
	Public Err

	'******************************************************************************
	'名　称：Initialize
	'概　要：cls_Scheduleクラスの初期化関数
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/10/13
	'更　新：
	'******************************************************************************
	Public Sub Initialize()
		Dim sDate
		Dim sTime
		IsData = False
		MaxIndex = -1

		If IsDay(GetForm("conf_startday", 1)) = True And ChkTime(GetForm("conf_starttime", 1)) = True Then
			sDate = GetForm("conf_startday", 1)
			sDate = Left(sDate, 4) & "-" & Mid(sDate, 5, 2) & "-" & Right(sDate, 2)
			sTime = GetForm("conf_starttime", 1)
			sTime = Left(sTime, 2) & ":" & Right(sTime, 2) & ":00.000"
			StartDay = sDate & " " & sTime
		End If

		If IsDay(GetForm("conf_endday", 1)) = True And ChkTime(GetForm("conf_endtime", 1)) = True Then
			sDate = GetForm("conf_endday", 1)
			sDate = Left(sDate, 4) & "-" & Mid(sDate, 5, 2) & "-" & Right(sDate, 2)
			sTime = GetForm("conf_endtime", 1)
			sTime = Left(sTime, 2) & ":" & Right(sTime, 2) & ":00.000"
			EndDay = sDate & " " & sTime
		End If

		If IsNumber(GetForm("sid", 2), 0, False) = True Then ScheduleID = GetForm("sid", 2)
		If IsNumber(GetForm("conf_briefingsessionid", 1), 0, False) = True Then BriefingSessionID = GetForm("conf_briefingsessionid", 1)
		If IsRE(GetForm("conf_publictype", 1), "^\d$", False) = True Then PublicType = GetForm("conf_publictype", 1)
		If IsRE(GetForm("conf_scheduletypecode", 1), "^\d\d\d$", False) = True Then ScheduleTypeCode = GetForm("conf_scheduletypecode", 1)
		If GetForm("conf_subject", 1) <> "" Then Subject = GetForm("conf_subject", 1)
		If GetForm("conf_body", 1) <> "" Then Body = GetForm("conf_body", 1)
		If GetForm("conf_placename", 1) <> "" Then PlaceName = GetForm("conf_placename", 1)
		If IsNumber(GetForm("conf_longitude", 1), 0, True) = True Then Longitude = GetForm("conf_longitude", 1)
		If IsNumber(GetForm("conf_latitude", 1), 0, True) = True Then Latitude = GetForm("conf_latitude", 1)

		If PublicType <> "" And StartDay <> "" And Subject <> "" And Body <> "" Then IsData = True
		'ScheduleIDチェック
		If GetForm("sid", 2) <> "" And IsNumber(GetForm("sid", 2), 0, False) = False Then
			'idあるけど数字じゃないから登録させない！
			IsData = False
		End If
	End Sub

	'******************************************************************************
	'名　称：GetRegSQL
	'概　要：up_Reg_Schedule UserInfo実行SQL取得
	'備　考：
	'作成者：Lis Kokubo
	'作成日：2006/10/13
	'更　新：
	'******************************************************************************
	Public Function GetRegSQL(ByVal vUserCode)
		If IsData = False Then Exit Function

		GetRegSQL = "up_Reg_Schedule '" & ScheduleID & "'" & _
			",'" & vUserCode & "'" & _
			",'" & BriefingSessionID & "'" & _
			",'" & PublicType & "'" & _
			",'" & ScheduleTypeCode & "'" & _
			",'" & StartDay & "'" & _
			",'" & EndDay & "'" & _
			",'" & Subject & "'" & _
			",'" & Body & "'" & _
			",'" & PlaceName & "'" & _
			",'" & Longitude & "'" & _
			",'" & Latitude & "'" & vbCrLf
	End Function
End Class
%>
