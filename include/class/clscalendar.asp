<%
Class clsCalendar
	Private DaysClass()'�e���t�ɓK�p����style

	Public ClassSun
	Public ClassSat
	Public ClassDay
	Public ClassSunOff
	Public ClassSatOff
	Public ClassDayOff
	Public Days()		'���t
	Public DaysLink()	'���t���N���b�N�������̃����N��t�q�k�z��B�����N�z��(0)�̓J�����_�[�̍ŏ��̓��j���̂���, �Ώی��̂P���ڂ̂��̂ł͂Ȃ��I
	Public DaysBody()	'���t�ŕ\������e�L�X�g�B���e�z��(0)�̓J�����_�[�̍ŏ��̓��j���̂���, �Ώی��̂P���ڂ̂��̂ł͂Ȃ��I
	Public Weekdays()	'�j���l
	Public Holidays()	'�Փ�

	'*******************************************************************************
	'�T�@�v�F�J�����_�[�̕\���ő�������擾
	'���@�l�F�V���~�U�T�ԁ��S�Q��
	'���@���F2009/08/25 LIS K.Kokubo �쐬
	'*******************************************************************************
	Public Property Get MAXCALENDARNUM
		MAXCALENDARNUM = 41
	End Property

	'*******************************************************************************
	'�T�@�v�F�R���X�g���N�^
	'���@���F
	'�߂�l�F
	'���@�l�F
	'���@���F2009/08/25 LIS K.Kokubo �쐬
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
	'�T�@�v�F�j���f�B�N�V���i�����擾
	'���@���FvStartDay	�F�j�����擾���鉺�����t
	'�߂�l�F
	'���@�l�FvStartDay�`MAXCALENDARNUM���Ԓ��ɂ���j�����擾
	'���@���F2009/08/25 LIS K.Kokubo �쐬
	'*******************************************************************************
	Private Function dicHoliday(vStartDay)
		Dim sSQL
		Dim oRS
		Dim flgQE
		Dim sSQLErr
		Dim idx

		Set dicHoliday = Server.CreateObject("scripting.dictionary")

		sSQL = "/* �J�����_�[�֐� */"
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
	'�T�@�v�F�j���f�B�N�V���i�����擾
	'���@���FvStartDay	�F�j�����擾���鉺�����t
	'�߂�l�F
	'���@�l�FvStartDay�`MAXCALENDARNUM���Ԓ��ɂ���j�����擾
	'���@���F2009/08/25 LIS K.Kokubo �쐬
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
	'�T�@�v�FDaysBody�ɒl��ݒ肷��
	'���@���FvDay
	'�@�@�@�FvBody
	'�߂�l�F
	'���@���F2009/08/28 LIS K.Kokubo �쐬
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
	'�T�@�v�F�X�P�W���[���J�����_�[HTML�𐶐����ďo�͂���
	'���@���FvYMKey	�F�N���L�[��
	'�@�@�@�FvYM	�F�N��(YYYYMM)
	'�߂�l�FString	�F�g�s�l�k
	'���@���F2009/08/25 LIS K.Kokubo �쐬
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
		sLinkBefore = "<a href=""?" & vYMKey & "=" & sLinkBefore & """>" & Left(sLinkBefore,4) & "�N" & CInt(Right(sLinkBefore,2)) & "��&nbsp;&lt;&lt;</a>"

		sLinkNext = Year(DateAdd("m",1,dHeadDay)) & Right("0" & Month(DateAdd("m",1,dHeadDay)),2)
		sLinkNext = "<a href=""?" & vYMKey & "=" & sLinkNext & """>&gt;&gt;&nbsp;" & Left(sLinkNext,4) & "�N" & CInt(Right(sLinkNext,2)) & "��</a>"

		sHTML = ""
		sHTML = sHTML & "<div style=""margin-bottom:3px;"">"
		sHTML = sHTML & "<div style=""float:left;width:33%;"">" & sLinkBefore & "</div>"
		sHTML = sHTML & "<div style=""float:left;width:34%;text-align:center;"">" & Left(vYM,4) & "�N" & CInt(Right(vYM,2)) & "��" & "</div>"
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
	'�T�@�v�F�X�P�W���[���J�����_�[HTML�𐶐����ďo�͂���
	'���@���FvDay	�F��{���t�iyyyymmdd�j
	'�@�@�@�FvWidth	�F�J�����_�[�̕��F "760px","100%"�ȂǁB
	'�߂�l�FString	�F�g�s�l�k
	'���@���F2009/08/25 LIS K.Kokubo �쐬
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
				'<�s����>
				If idx < MAXCALENDARNUM Then
					If Month(Days(idx)) <> Month(Days(idx + 1)) Then Exit For
					If ChkDateAdd("m", 1, Days(idx)) Then
						If iThisMonth <> Month(Days(idx)) And iThisMonth <> Month(Days(idx + 1)) Then Exit For
					End If
				End If
				'</�s����>
			End If
		Next

		sHTML = sHTML & "</tbody>"
		sHTML = sHTML & "</table>" & vbCrLf

		htmlCalendar = sHTML
	End Function
End Class
%>
