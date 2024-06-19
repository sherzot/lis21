<%
Class clsCalendar
	Private DaysClass()'各日付に適用するstyle

	Public ClassSun
	Public ClassSat
	Public ClassDay
	Public ClassSunOff
	Public ClassSatOff
	Public ClassDayOff
	Public Days()		'日付
	Public DaysLink()	'日付をクリックした時のリンク先ＵＲＬ配列。リンク配列(0)はカレンダーの最初の日曜日のもの, 対象月の１日目のものではない！
	Public DaysBody()	'日付で表示するテキスト。内容配列(0)はカレンダーの最初の日曜日のもの, 対象月の１日目のものではない！
	Public Weekdays()	'曜日値
	Public Holidays()	'祭日

	'*******************************************************************************
	'概　要：カレンダーの表示最大日数を取得
	'備　考：７日×６週間＝４２日
	'履　歴：2009/08/25 LIS K.Kokubo 作成
	'*******************************************************************************
	Public Property Get MAXCALENDARNUM
		MAXCALENDARNUM = 41
	End Property

	'*******************************************************************************
	'概　要：コンストラクタ
	'引　数：
	'戻り値：
	'備　考：
	'履　歴：2009/08/25 LIS K.Kokubo 作成
	'*******************************************************************************
	Private Sub Class_Initialize()
		ClassSun = "sun"
		ClassSat = "sat"
		ClassDay = "day"
		ClassSunOff = "offsun"
		ClassSatOff = "offsat"
		ClassDayOff = "offday"

		ReDim DaysLink(MAXCALENDARNUM)
		ReDim DaysBody(MAXCALENDARNUM)
		ReDim DaysClass(MAXCALENDARNUM)
		ReDim Weekdays(MAXCALENDARNUM)
		ReDim Holidays(MAXCALENDARNUM)
		ReDim Days(MAXCALENDARNUM)
	End Sub

	'*******************************************************************************
	'概　要：祝日ディクショナリを取得
	'引　数：vStartDay	：祝日を取得する下限日付
	'戻り値：
	'備　考：vStartDay～MAXCALENDARNUM日間中にある祝日を取得
	'履　歴：2009/08/25 LIS K.Kokubo 作成
	'*******************************************************************************
	Private Function dicHoliday(vStartDay)
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sSQLErr
		Dim idx

		Set dicHoliday = Server.CreateObject("scripting.dictionary")

		sSQL = "/* カレンダー関数 */"
		sSQL = sSQL & "SELECT HoliDay, Comment FROM L_TimeTable_HolidayRegist WHERE PublicFlag = '1' AND (HoliDay BETWEEN '" & GetDateStr(vStartDay, "") & "' AND '" & GetDateStr(DateAdd("d",MAXCALENDARNUM-1,vStartDay),"") & "');"
		flgQE = QUERYEXE(dbconn, oRS, sSQL, sSQLErr)
		Do While GetRSState(oRS) = True
			Call dicHoliday.Add(oRS.Collect("HoliDay"), oRS.Collect("Comment"))

			oRS.MoveNext
		Loop
		Call RSClose(oRS)

		For idx = 0 To MAXCALENDARNUM
			If dicHoliday.Exists(GetDateStr(Days(idx),"")) = True Then Holidays(idx) = dicHoliday.Item(GetDateStr(Days(idx),""))
		Next
	End Function

	'*******************************************************************************
	'概　要：祝日ディクショナリを取得
	'引　数：vStartDay	：祝日を取得する下限日付
	'戻り値：
	'備　考：vStartDay～MAXCALENDARNUM日間中にある祝日を取得
	'履　歴：2009/08/25 LIS K.Kokubo 作成
	'*******************************************************************************
	Public Function getStartDay(vYM)
		Dim sYM
		Dim dStartDay

		If ChkDate8(vYM&"01") = True Then
			sYM = vYM
		Else
			sYM = Year(Date) & Right("0" & Month(Date),2)
		End If

		dStartDay = CDate(Left(sYM,4) & "/" & Right(sYM,2) & "/" & "1")
		dStartDay = DateAdd("d", (Weekday(dStartDay) - 1) * -1, dStartDay)

		getStartDay = dStartDay
	End Function

	'*******************************************************************************
	'概　要：DaysBodyに値を設定する
	'引　数：vDay
	'　　　：vBody
	'戻り値：
	'履　歴：2009/08/28 LIS K.Kokubo 作成
	'*******************************************************************************
	Public Function setDaysBody(ByVal vDay, ByVal vBody)
		Dim idx

		If ChkDate8(vDay) = True Then vDay = CDate(Left(vDay,4) & "/" & Mid(vDay,5,2) & "/" & Right(vDay,2))

		For idx = 0 To MAXCALENDARNUM
			If Days(idx) = vDay Then
				DaysBody(idx) = DaysBody(idx) & vBody
			End If
		Next
	End Function

	'*******************************************************************************
	'概　要：スケジュールカレンダーHTMLを生成して出力する
	'引　数：vYMKey	：年月キー名
	'　　　：vYM	：年月(YYYYMM)
	'戻り値：String	：ＨＴＭＬ
	'履　歴：2009/08/25 LIS K.Kokubo 作成
	'*******************************************************************************
	Public Function htmlCalendarMonthLink(ByVal vYMKey, ByVal vYM)
		Dim dHeadDay
		Dim sLinkBefore
		Dim sLinkNext
		Dim sHref
		Dim sHTML

		If ChkDate8(vYM&"01") = False Then
			vYM = Year(Date) & Right("0"&Month(Date),2)
		End If

		dHeadDay = CDate(Left(vYM,4) & "/" & Right(vYM,2) & "/01")

		sLinkBefore = Year(DateAdd("m",-1,dHeadDay)) & Right("0" & Month(DateAdd("m",-1,dHeadDay)),2)
		sLinkBefore = "<a href=""?" & vYMKey & "=" & sLinkBefore & """>" & Left(sLinkBefore,4) & "年" & CInt(Right(sLinkBefore,2)) & "月&nbsp;&lt;&lt;</a>"

		sLinkNext = Year(DateAdd("m",1,dHeadDay)) & Right("0" & Month(DateAdd("m",1,dHeadDay)),2)
		sLinkNext = "<a href=""?" & vYMKey & "=" & sLinkNext & """>&gt;&gt;&nbsp;" & Left(sLinkNext,4) & "年" & CInt(Right(sLinkNext,2)) & "月</a>"

		sHTML = ""
		sHTML = sHTML & "<div style=""margin-bottom:3px;"">"
		sHTML = sHTML & "<div style=""float:left;width:33%;"">" & sLinkBefore & "</div>"
		sHTML = sHTML & "<div style=""float:left;width:34%;text-align:center;"">" & Left(vYM,4) & "年" & CInt(Right(vYM,2)) & "月" & "</div>"
		sHTML = sHTML & "<div style=""float:left;width:33%;text-align:right;"">" & sLinkNext & "</div>"
		sHTML = sHTML & "<div style=""clear:both;""></div>"
		sHTML = sHTML & "</div>" & vbCrLf

		htmlCalendarMonthLink = sHTML
	End Function

	Public Function setDays(ByVal vYM)
		Dim sYM
		Dim dStartDay
		Dim idx
		Dim oDicHoliday

		If ChkDate8(vYM&"01") = True Then
			sYM = vYM
		Else
			sYM = Year(Date) & Right("0" & Month(Date),2)
		End If

		dStartDay = getStartDay(sYM)

		For idx = 0 To MAXCALENDARNUM
			Days(idx) = DateAdd("d", idx, dStartDay)
			Weekdays(idx) = Weekday(Days(idx))
		Next

		Set oDicHoliday = dicHoliday(dStartDay)

		For idx = 0 To MAXCALENDARNUM
			If CInt(Right(sYM,2)) <> Month(Days(idx)) Then
				Select Case Weekday(Days(idx))
				Case 1: DaysClass(idx) = ClassSunOff
				Case 7: DaysClass(idx) = ClassSatOff
				Case Else: DaysClass(idx) = ClassDayOff
				End Select
			Else
				If oDicHoliday.Exists(GetDateStr(Days(idx), "")) = True Then
					DaysClass(idx) = ClassSun
				Else
					Select Case Weekday(Days(idx))
					Case 1: DaysClass(idx) = ClassSun
					Case 7: DaysClass(idx) = ClassSat
					Case Else: DaysClass(idx) = ClassDay
					End Select
				End If
			End If
		Next
	End Function

	'*******************************************************************************
	'概　要：スケジュールカレンダーHTMLを生成して出力する
	'引　数：vDay	：基本日付（yyyymmdd）
	'　　　：vWidth	：カレンダーの幅： "760px","100%"など。
	'戻り値：String	：ＨＴＭＬ
	'履　歴：2009/08/25 LIS K.Kokubo 作成
	'*******************************************************************************
	Public Function htmlCalendar(ByVal vYM, ByVal vWidth)
		Dim sYM
		Dim dStartDay
		Dim idx
		Dim iThisYear
		Dim iThisMonth
		Dim sTodayStyle
		Dim sHref
		Dim sHTML

		If ChkDate8(vYM&"01") = True Then
			sYM = vYM
		Else
			sYM = Year(Date) & Right("0" & Month(Date),2)
		End If

		iThisYear = CInt(Left(sYM,4))
		iThisMonth = CInt(Right(sYM,2))
		dStartDay = getStartDay(sYM)

		sHTML = sHTML & "<table class=""calendar"" border=""0"" style=""width:" & vWidth & "; border-collapse:collapse;"">"
		sHTML = sHTML & "<colgroup>"
		sHTML = sHTML & "<col style=""width:15%;""></col>"
		sHTML = sHTML & "<col style=""width:14%;""></col>"
		sHTML = sHTML & "<col style=""width:14%;""></col>"
		sHTML = sHTML & "<col style=""width:14%;""></col>"
		sHTML = sHTML & "<col style=""width:14%;""></col>"
		sHTML = sHTML & "<col style=""width:14%;""></col>"
		sHTML = sHTML & "<col style=""width:15%;""></col>"
		sHTML = sHTML & "</colgroup>"
		sHTML = sHTML & "<thead>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<th class=""" & ClassSun & """>Sun</th>"
		sHTML = sHTML & "<th class=""" & ClassDay & """>Mon</th>"
		sHTML = sHTML & "<th class=""" & ClassDay & """>Tue</th>"
		sHTML = sHTML & "<th class=""" & ClassDay & """>Wed</th>"
		sHTML = sHTML & "<th class=""" & ClassDay & """>Thu</th>"
		sHTML = sHTML & "<th class=""" & ClassDay & """>Fri</th>"
		sHTML = sHTML & "<th class=""" & ClassSat & """>Sat</th>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</thead>"
		sHTML = sHTML & "<tbody>" & vbCrLf

		For idx = 0 To MAXCALENDARNUM
			If Weekdays(idx) = 1 Then sHTML = sHTML & "<tr>"
			If Date = Days(idx) Then
				sTodayStyle = " style=""background-color:#ffff99;"""
			Else
				sTodayStyle = ""
			End If

			sHTML = sHTML & "<td class=""" & DaysClass(idx) & """ " & sTodayStyle & ">"
			sHTML = sHTML & "<a href=""" & DaysLink(idx) & """>" & Day(Days(idx)) & "</a>"
			If Holidays(idx) <> "" Then sHTML = sHTML & "&nbsp;<span style=""font-size:80%;"">" & Holidays(idx) & "</span>"
			sHTML = sHTML & "<br>"
			sHTML = sHTML & DaysBody(idx)
			sHTML = sHTML & "</td>"

			If Weekdays(idx) = 7 Then
				sHTML = sHTML & "</tr>" & vbCrLf
				'<行調節>
				If idx < MAXCALENDARNUM Then
					If Month(Days(idx)) <> Month(Days(idx + 1)) Then Exit For
					If ChkDateAdd("m", 1, Days(idx)) Then
						If iThisMonth <> Month(Days(idx)) And iThisMonth <> Month(Days(idx + 1)) Then Exit For
					End If
				End If
				'</行調節>
			End If
		Next

		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>" & vbCrLf

		htmlCalendar = sHTML
	End Function
End Class
%>
