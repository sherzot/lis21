<%
'**********************************************************************************************************************
'概　要：人材一覧 /staff/person_result.asp
'　　　：プロフィール /staff/person_detail.asp
'　　　：上記ページで出力用の関数群をこのファイルに用意する。
'　　　：
'　　　：■■■　前提条件　■■■
'　　　：要事前インクルード
'　　　：/config/personel.asp
'　　　：/include/commonfunc.asp
'一　覧：■■■　共通　■■■
'　　　：GetStrLastAccessDay				：最終アクセス日の文言を取得
'　　　：■■■　人材一覧ページ出力用　■■■
'　　　：GetHtmlPageControl					：求職者一覧ページのページコントロールＨＴＭＬを取得
'　　　：DspStaffOne						：求職者一覧表示の１人分の枠を生成
'　　　：■■■　プロフィールページ出力用　■■■
'　　　：DspProfileBase						：プロフィールページの基本情報部分を出力
'　　　：DspProfileNearbyStation			：プロフィールページの最寄駅部分を出力
'　　　：DspProfileEducateHistory			：プロフィールページの学歴情報部分を出力
'　　　：DspProfileCareerHistory			：プロフィールページの職歴情報部分を出力
'　　　：DspProfileCareerHistoryIT			：プロフィールページのＩＴ職歴情報部分を出力
'　　　：DspProfileSkill					：プロフィールページのスキル部分を出力
'　　　：DspProfileSkillSimple					：プロフィールページのスキル部分を出力(簡易登録)
'　　　：DspProfileHope						：プロフィールページの希望条件部分を出力
'　　　：DspProfileHopeWorkingType			：プロフィールページの希望条件勤務形態部分を出力
'　　　：DspProfileHopeIndustry				：プロフィールページの希望条件業種部分を出力
'　　　：DspProfileHopeJobType				：プロフィールページの希望条件業種部分を出力
'　　　：DspProfileHopeWorkingPlace			：プロフィールページの希望条件勤務地部分を出力
'　　　：DspProfileHopeSalary				：プロフィールページの希望条件給与部分を出力
'　　　：DspProfileHopeSpan					：プロフィールページの希望条件期間・時間部分を出力
'　　　：DspProfileHopeWelfare				：プロフィールページの希望条件福利厚生部分を出力
'　　　：DspCareerAnalyzer					：プロフィールページの将来的な理想像を出力
'　　　：DspProfileMail						：プロフィールページの最新送信メール状況部分を出力
'　　　：DspProfileStaffCode				：プロフィールページの求職者コードを出力
'　　　：DspProfileEditButton				：プロフィールページの各編集ボタンを出力
'　　　：DspProfileUpdateDay				：プロフィールページの最終更新日を出力
'　　　：DspProfileAttentionCareerHistory	：プロフィールページの職歴入力無しに対する文言を出力
'　　　：■■■　ＤＢ書き込み　■■■
'　　　：RegMailMagazineAccess				：メールマガジンからのプロフィールページへのアクセスをログに記録
'　　　：UpdAccessCount						：プロフィールページのアクセス回数のカウントアップ
'　　　：SendMailStaffEdit					：情報更新の通知をリスの社員にメールする
'　　　：RegAccessHistoryStaff				：企業が求職者詳細を閲覧したらログに書き込む
'　　　：RegAccessHistoryStaffList			：企業が求職者一覧を閲覧したらログに書き込む
'**********************************************************************************************************************

'******************************************************************************
'概　要：最終アクセス日の文言を取得
'引　数：vLastAccessDay	：最終アクセス日
'備　考：
'履　歴：2009/04/27 LIS K.Kokubo 作成
'　　　：2009/06/23 LIS K.Kokubo 「３～６ヶ月アクセスなし」を追加
'******************************************************************************
Function GetStrLastAccessDay(ByVal vLastAccessDay)
	Dim sTxt

	If DateAdd("m", -3, Date) <= vLastAccessDay Then
		sTxt = GetDateStr(vLastAccessDay, "/")
	ElseIf DateAdd("m", -6, Date) <= vLastAccessDay Then
		sTxt = GetDateStr(vLastAccessDay, "/") & "（３～６ヶ月アクセスなし）"
	Else
		sTxt = "６ヶ月以上アクセスなし"
	End If
	'<ログイン日まるめ表示Ver.>
'	If DateDiff("d", rRS.Collect("LastAccessDay"), Date) <= 2 Then
'		sLastAccess = "３日以内"
'	ElseIf DateDiff("d", rRS.Collect("LastAccessDay"), Date) <= 6 Then
'		sLastAccess = "７日以内"
'	ElseIf DateDiff("m", rRS.Collect("LastAccessDay"), Date) = 0 Then
'		sLastAccess = "１ヶ月以内"
'	ElseIf DateDiff("m", rRS.Collect("LastAccessDay"), Date) <= 2 Then
'		sLastAccess = "３ヶ月以内"
'	ElseIf DateDiff("m", rRS.Collect("LastAccessDay"), Date) <= 5 Then
'		sLastAccess = "半年以内"
'	ElseIf DateDiff("m", rRS.Collect("LastAccessDay"), Date) <= 11 Then
'		sLastAccess = "１年以内"
'	Else
'		sLastAccess = "１年以上×"
'	End If
	'</ログイン日まるめ表示Ver.>

	GetStrLastAccessDay = sTxt
End Function

'******************************************************************************
'概　要：求職者一覧ページのページコントロールＨＴＭＬを取得
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：求職者検索結果を保持するのレコードセットオブジェクト
'　　　：vPageSize		：１ページあたりの表示件数
'　　　：vPage			：表示中ページ
'使用元：しごとナビ/staff/person_list.asp
'　　　：しごとナビ/order/company_order.asp
'備　考：
'更　新：2007/02/11 LIS K.Kokubo 作成
'******************************************************************************
Function GetHtmlPageControl(ByRef rDB, ByRef rRS, ByVal vPageSize, ByVal vPage)
	Dim iMaxPage
	Dim iLine
	Dim S_Page
	Dim E_Page
	Dim Sort
	Dim idx

	If GetRSState(rRS) = False Then Exit Function

	If vPage <> "" Then vPage = CInt(vPage)

	'ページあたりの表示件数
	rRS.PageSize = vPageSize

	iMaxPage = rRS.PageCount
	If vPage > iMaxPage Then vPage = iMaxPage
	rRS.AbsolutePage = vPage

	'画面上に表示する開始・終了ページ番号を設定
	'表示開始ページ番号を指定
	S_Page = vPage - 5
	If S_Page < 1 Then
		S_Page = 1
	End If

	'表示終了ページ番号を指定
	E_Page = vPage + 4
	If E_Page < 10 Then E_Page = 10
	If E_Page > iMaxPage Then
		E_Page = iMaxPage
	End If
	If S_Page > iMaxPage - 9 And iMaxPage - 9 > 0 Then S_Page = iMaxPage - 9

	GetHtmlPageControl = ""
	GetHtmlPageControl = GetHtmlPageControl & "<table class=""cw"" style=""margin:10px 0px;"">"
	GetHtmlPageControl = GetHtmlPageControl & "<tbody>"
	GetHtmlPageControl = GetHtmlPageControl & "<tr>"
	GetHtmlPageControl = GetHtmlPageControl & "<td style=""width:88px; padding:5px; border:1px dotted #666699; border-width:1px 0px 1px 1px; text-align:center; background-color:#e8e8ff;"">"

	If vPage > 1 Then GetHtmlPageControl = GetHtmlPageControl & "<a href='javascript:ChgPage(" & vPage - 1 & ");'>前のページ</a>"
	GetHtmlPageControl = GetHtmlPageControl & "</td>"
	GetHtmlPageControl = GetHtmlPageControl & "<td style=""width:389px; padding:5px; border:1px dotted #666699; border-width:1px 0px 1px 0px; text-align:center; background-color:#e8e8ff;"">"

	If S_Page <> 1 Then GetHtmlPageControl = GetHtmlPageControl & "…"

	For idx = S_Page To E_Page	'ページ番号を表示
		GetHtmlPageControl = GetHtmlPageControl & "　"
		If idx = vPage Then		'指定ページの表示
			GetHtmlPageControl = GetHtmlPageControl & "[" & idx & "]"
		Else
			GetHtmlPageControl = GetHtmlPageControl & "<a href='javascript:ChgPage(" & idx & ");'>" & idx & "</a>"
		End If
	Next

	If E_Page < iMaxPage Then GetHtmlPageControl = GetHtmlPageControl & "　…"

	GetHtmlPageControl = GetHtmlPageControl & "</td>"
	GetHtmlPageControl = GetHtmlPageControl & "<td style=""width:89px; padding:5px; border:1px dotted #666699; border-width:1px 1px 1px 0px; text-align:center; background-color:#e8e8ff;"">"

	If vPage < iMaxPage Then GetHtmlPageControl = GetHtmlPageControl & "<a href='javascript:ChgPage(" & vPage + 1 & ");'>次のページ</a>"

	GetHtmlPageControl = GetHtmlPageControl & "</td>"
	GetHtmlPageControl = GetHtmlPageControl & "</tr>"
	GetHtmlPageControl = GetHtmlPageControl & "</tbody>"
	GetHtmlPageControl = GetHtmlPageControl & "</table>"
End Function

'******************************************************************************
'概　要：求職者一覧表示の１人分の枠を生成
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：求職者検索結果を保持するのレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [vUserID]
'使用元：ナビ/staff/person_list.asp
'備　考：
'履　歴：2007/04/09 LIS K.Kokubo 作成
'　　　：2009/04/27 LIS K.Kokubo 最終アクセス日を表示
'　　　：2015/10/29 LIS K.Kimura 個人情報保護のため、スタッフコード、最終アクセス日、自己PRを非表示に変更
'******************************************************************************
Function DspStaffOne(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim dbViewStaffDay		'求職者閲覧日

	Dim sStaffCode
	Dim sClassName
	Dim sNewHtml			'新着アイコンＨＴＭＬ
	Dim sHotHtml			'ＨＯＴアイコンＨＴＭＬ
	Dim sWorkerAlarmHtml	'WorkerAlarmアイコンＨＴＭＬ
	Dim sScoutHtml			'メール送信済み（他案件含む）アイコンＨＴＭＬ
	Dim sViewStaffHtml		'閲覧済み（他案件含む）アイコンＨＴＭＬ
	Dim sLastAccess			'ログイン日表示
	Dim wSelfPR
	Dim wRow
	Dim sStaffDetailURL

	'変数初期化 start
	sNewHtml = ""
	sHotHtml = ""
	sWorkerAlarmHtml = ""
	sScoutHtml = ""
	sViewStaffHtml = ""
	'変数初期化 end

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")
	If vUserType = "staff" Or vUserType = "" Then
		sClassName = "pattern9"
	Else
		sClassName = "pattern8"
	End If

	If InStr(LCase(Request.ServerVariables("URL")), "jinzai") <> 0 Then
		'人材採用ページの求職者詳細ＵＲＬ
		sStaffDetailURL = "./person_detail.asp?staffcode=" & sStaffCode
	Else
		'しごとナビの求職者詳細ＵＲＬ
		sStaffDetailURL = "/staff/person_detail.asp?staffcode=" & sStaffCode & "&amp;ordercode=" & vOrderCode
	End If

	If sStaffCode <> "" Then
		wRow = 4
		'自己PR
		wSelfPR = ""
		sSQL = "EXEC up_DtlStaffList '" & sStaffCode & "', '" & vUserID & "', '" & vOrderCode & "';"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			Set oRS.ActiveConnection = Nothing

			dbViewStaffDay = oRS.Collect("ViewStaffDay")

			'自己ＲＲ行追加
			'If ChkStr(oRS.Collect("SelfPR")) <> "" Then
			'	wSelfPR = oRS.Collect("SelfPR")
			'	wRow = wRow + 1
			'End If

			'職歴行追加
			If ChkStr(oRS.Collect("CareerJobType")) <> "" Then
				wRow = wRow + 1
			End If

			'<アイコン設定>
			'新着
			If oRS.Collect("NewFlag") = "1" Then sNewHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/c_new.gif"" alt="""" border=""0""></a>&nbsp;"
			'閲覧回数多い
			If oRS.Collect("HotFlag") = "1" Then sHotHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/c_hot.gif"" alt="""" border=""0""></a>&nbsp;"
			'WorkerAlarm
			If oRS.Collect("WorkerAlarmFlag") = "1" Then sWorkerAlarmHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/workeralarm.gif"" alt=""すぐに働ける求職者マーク"" border=""0""></a>&nbsp;"
			If oRS.Collect("ViewStaffFlag") = "1" Then
				'閲覧済み
				sViewStaffHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/viewstaff.gif"" border=""0"" alt=""""></a>&nbsp;"
			ElseIf oRS.Collect("ViewStaffOtherFlag") = "1" Then
				'他案件で閲覧済み
				sViewStaffHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/viewstaffother.gif"" border=""0"" alt=""""></a>&nbsp;"
			End If
			If oRS.Collect("OrderScoutFlag") = "1" Then
				'メール送信済み
				sScoutHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/contact.gif"" border=""0"" alt=""""></a>&nbsp;"
			ElseIf oRS.Collect("CompanyScoutFlag") = "1" Then
				'他案件メール送信済み
				sScoutHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/other_send.gif"" border=""0"" alt=""""></a>&nbsp;"
			End If
			If oRS.Collect("MailReceiveFlag") = "1" Then
				'メール受信済み
				sViewStaffHtml = "<a href=""javascript:void(0)"" onclick='window.open(""/staff/icon_navi.html"",""count"",""menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300"")'><img src=""/img/mailreceive.gif"" border=""0"" alt=""""></a>&nbsp;"
			End If
			'</アイコン設定>

			'<日付情報,一括メール>
			If vUserType = "company" Then
				Response.Write "<div class=""m0"" style=""float:left;width:20%;"">"
				If InStr(LCase(Request.ServerVariables("URL")), "jinzai") = 0 Then
					If oRS.Collect("LumpMailFlag") = "0" And oRS.Collect("PublicFlag") = "1" Then
						Response.Write "<input id=""lumpmail" & sStaffCode & """ class=""btn1"" type=""button"" value=""一括ﾒｰﾙ予約"" style=""width:120px;"" onclick=""open('/company/lumpmail/reg.asp?staffcode=" & sStaffCode & "&amp;ordercode=" & vOrderCode & "','lumpmail','menuber=no,toolber=no,location=no,directories=no,status=no,scrollbars=yes,resizable=yes,width=420,height=300');"">"
					ElseIf oRS.Collect("LumpMailFlag") = "1" Then
						Response.Write "<input class=""btn1"" type=""button"" value=""一括ﾒｰﾙ予約済"" disabled style=""width:120px;"">"
					Else
						Response.Write "<input class=""btn1"" type=""button"" value=""一括ﾒｰﾙ予約不可"" disabled style=""width:120px;"">"
					End If
				End If
				Response.Write "</div>"
			End If
			Response.Write "<p class=""m0"" style=""float:left;width:80%;text-align:right;font-size:10px;"">"
			If ChkStr(dbViewStaffDay) <> "" Then
				If dbViewStaffDay < rRS.Collect("UpdateDay") Then
					Response.Write "<span style=""color:#ff0000;"">※求職者が情報を更新しました</span>&nbsp;&nbsp;"
					Response.Write "詳細閲覧日：" & GetDateStr(dbViewStaffDay, "/") & "&nbsp;&nbsp;"
					Response.Write "更新日：" & GetDateStr(rRS.Collect("UpdateDay"), "/") & "&nbsp;&nbsp;"
				End If
			End If
			'Response.Write "最終アクセス日：" & GetStrLastAccessDay(rRS.Collect("LastAccessDay"))
			Response.Write "</p>"
			Response.Write "<div style=""clear:both;""></div>"
			'</日付情報,一括メール>

			Response.Write "<table class=""" & sClassName & " cw"" border=""0"">"
			'<１行目>
			Response.Write "<thead>"
			Response.Write "<tr>"
			Response.Write "<th colspan=""4"" style=""text-align:left;"">"
			'Response.Write oRS.Collect("StaffCode") & "　"

			'If Len(oRS.Collect("PrefectureName") & oRS.Collect("City")) > 10 Then
			'	Response.Write Left(oRS.Collect("PrefectureName") & oRS.Collect("City"),10) & "..."
			'Else
				Response.Write "　" & oRS.Collect("PrefectureName")
			'End If

			'Response.Write "在住（" & Left(oRS.Collect("Age"),1) & "歳／" & oRS.Collect("Sex") & "性" & "）"
			Response.Write "在住（" & Left(oRS.Collect("Age"),1) & "0代／" & oRS.Collect("Sex") & "性" & "）"
			'<求職者の状態フラグ>
			Response.Write sNewHtml
			Response.Write sHotHtml
			Response.Write sWorkerAlarmHtml
			Response.Write sScoutHtml
			Response.Write sViewStaffHtml
			'Response.Write sLastAccess
			'</求職者の状態フラグ>
			Response.Write "</th>"
			Response.Write "</tr>"
			Response.Write "</thead>"
			'<１行目>
			Response.Write "<tbody>"

			'<２行目>
			Response.Write "<tr>"
			Response.Write "<th align=""center"" rowspan=""" & wRow & """>"
			'Response.Write "<a href=""javascript:PersonDetail('" & oRS.Collect("StaffCode") & "');"">"
			Response.Write "<a href=""" & sStaffDetailURL & """ target=""_blank"">"
			Response.Write "<img src=""/img/shousai.gif"" border=""0"" alt=""""><br>"
			Response.Write "<b>詳細</b>"
			Response.Write "</a>"
			Response.Write "</th>"
			Response.Write "<th colspan=""2"">"
			Response.Write "現在の状況"
			Response.Write "</th>"
			Response.Write "<td>"
			Response.Write oRS.Collect("OperateClassWebName")
			Response.Write "</td>"
			Response.Write "</tr>"
			'</２行目>

			'<３行目>
			If ChkStr(oRS.Collect("CareerJobType")) <> "" Then
				Response.Write "<tr>"
				Response.Write "<th colspan=""2"">経験職種</th>"
				Response.Write "<td>"
				'Response.Write Replace(ChkStr(oRS.Collect("CareerJobType")), vbCrLf, "<br>")
				Response.Write Replace(RegExpReplace(oRS.Collect("CareerJobType"), "(\(\d).*?(年\))"), vbCrLf, "<br>")
				Response.Write "</td>"
				Response.Write "</tr>"
			End If
			'</３行目>

			'<４行目>
			Response.Write "<tr>"
			Response.Write "<th rowspan=""3"">希望条件</th>"
			Response.Write "<th>職種</th>"
			Response.Write "<td>"
			Response.Write Replace(ChkStr(oRS.Collect("HopeJobType")), vbCrLf, "<br>")
			Response.Write "</td>"
			Response.Write "</tr>"
			'</４行目>

			'<５行目>
			Response.Write "<tr>"
			Response.Write "<th>勤務地</th>"
			Response.Write "<td>"
			Response.Write Replace(ChkStr(oRS.Collect("HopeWorkingPlace")), vbCrLf, "<br>")
			Response.Write "</td>"
			Response.Write "</tr>"
			'</５行目>

			'<６行目>
			Response.Write "<tr>"
			Response.Write "<th>雇用形態</th>"
			Response.Write "<td>"
			Response.Write Replace(ChkStr(oRS.Collect("HopeWorkingType")), vbCrLf, "<br>")
			Response.Write "</td>"
			Response.Write "</tr>"
			'</６行目>

			'<７行目>
			'If Trim(ChkStr(wSelfPR)) <> "" Then
			'	Response.Write "<tr>"
			'	Response.Write "<th colspan=""2"">自己PR</th>"
			'	Response.Write "<td>"

			'	If Len(wSelfPR) > 100 Then
			'		Response.Write Left(wSelfPR,100) & "..."
			'	Else
			'		Response.Write wSelfPR
			'	End If

			'	Response.Write "</td>"
			'	Response.Write "</tr>"
			'End If
			'</７行目>

			Response.Write "<tr>"
			Response.Write "<td style=""padding:0px;border-width:0px;width:50px;""></td>"
			Response.Write "<td style=""padding:0px;border-width:0px;width:75px;""></td>"
			Response.Write "<td style=""padding:0px;border-width:0px;width:75px;""></td>"
			Response.Write "<td style=""padding:0px;border-width:0px;width:400px;""></td>"
			Response.Write "</tr>"
			Response.Write "</tbody>"
			Response.Write "</table>"
			Response.Write "<br>"
		End If
	End If

	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：プロフィールページの基本情報部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：しごとナビ/staff/person_list.asp
'　　　：しごとナビ/order/company_order.asp
'備　考：
'履　歴：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileBase(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim sRS
	Dim oRS2
	Dim oRS3
	Dim sErrer
	Dim flgQE

	Dim sStaffCode
	Dim sTableClass
	Dim sComment

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

%>
	<table class="profileSmart smartBlock" style="display:none;">
		<thead>
        	<tr>
    			<th colspan="2">基本情報</th>
    		</tr>
    	</thead>
        <tbody>
        	<tr>
            	<th colspan="2" class="promidasi">稼動区分</th>
            </tr>
            <tr>
                <td colspan="2">
                <%
                    Response.Write rRS.Collect("OperateClassWeb")
                    If rRS.Collect("HopeWorkStartDay") <> "" Then
                        Response.Write "(勤務開始予定日：" & GetDateStr(rRS.Collect("HopeWorkStartDay"), "/") & ")"
                    End If
                %>
                </td>
            </tr>
            <tr>
            	<th colspan="2" class="promidasi">性別・年齢</th>
            </tr>
            <tr>
            	<td colspan="2">
                <%
					Response.Write rRS.Collect("Sex")
					Response.Write "・" & rRS.Collect("Age") & "歳"
                %>
                </td>
            </tr>
            <tr>
            	<th colspan="2" class="promidasi">住所地</th>
            </tr>
            <tr>
            	<td colspan="2"><%= rRS.Collect("PrefectureName") & rRS.Collect("City") %></td>
            </tr>
            <tr>
            	<th colspan="2" class="promidasi">自己PR</th>
            </tr>
            <tr>
            	<td colspan="2">
                <%
					Response.Write Replace(ChkStr(rRS.Collect("SelfPR")), vbCrLf, "<br>")
				%>
                </td>
            </tr>
            
                <%
					sComment = Array("")
					
					sSQL = "SELECT ResultId FROM P_MyNaviResult WITH(NOLOCK) WHERE StaffCode = '" & sStaffCode & "';"
					flgQE = QUERYEXE(dbconn, oRS2, sSQL, sError)
					If GetRSState(oRS2) = True Then
						sSQL  = "SELECT specialcomment FROM MyNavi_Result WITH(NOLOCK) WHERE id = '" & oRS2.Collect("resultid") & "';"
						flgQE = QUERYEXE(dbconn, oRS3, sSQL, sError)
						If GetRSState(oRS3) = True Then
				
							sSQL = "SELECT point1,point2,point3,point4,point5,point6 FROM P_MyNaviResult WITH(NOLOCK) WHERE StaffCode = '" & sStaffCode & "';"
							flgQE = QUERYEXE(dbconn, sRS, sSQL, sError)
							If GetRSState(sRS) = True Then
								%>
                                <tr>
                                    <th colspan="2" class="promidasi">適職診断「じぶんナビ」結果</th>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                <%
								
								sComment = Split(Replace(oRS3.Collect("specialcomment"),"あなた","この方"),"<br>")
								Response.Write sComment(0)
								
								Response.Write "</td>"
								Response.Write "</tr>"
							End If
						End If
					End If
					
					'<最寄駅 />
					Call DspProfileNearbyStationSmart(rDB, rRS, vUserType, vUserID)
					
					If (G_USERTYPE = "company" Or G_USERTYPE = "dispatch" Or G_USERTYPE = "talent" Or G_USERTYPE = "staff") And Left(sStaffCode,1) = "T" Then
				%>
            
            <tr>
            	<th colspan="2" class="promidasi">人事担当推薦文</th>
			</tr>
            <tr>
            	<td colspan="2">
					<%= rRS.Collect("RecommendationLetter") %>
				</td>
            </tr>
			<% End If %>
            <% If ChkStr(rRS.Collect("Learn")) <> "" Then %>
            <tr>
            	<th colspan="2" class="promidasi">学生時代に身につけたこと</th>
            </tr>
            <tr>
            	<td colspan="2"><%= rRS.Collect("Learn") %></td>
            </tr>
    		<% End If %>
    	</tbody>
    </table>
<%

	Response.Write "<table class=""" & sTableClass & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
	Response.Write "<colgroup>"
	Response.Write "<col style=""width:100px;"">"
	Response.Write "<col style=""width:100px;"">"
	Response.Write "<col style=""width:400px;"">"
	Response.Write "</colgroup>"
	Response.Write "<thead>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""3"">基本情報</th>"
	Response.Write "</tr>"
	Response.Write "</thead>"
	Response.Write "<tbody>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">稼動区分</td>"
	Response.Write "<td>"
	Response.Write rRS.Collect("OperateClassWeb")
	If rRS.Collect("HopeWorkStartDay") <> "" Then
		Response.Write "(勤務開始予定日：" & GetDateStr(rRS.Collect("HopeWorkStartDay"), "/") & ")"
	End If
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">性別・年齢</th>"
	Response.Write "<td>"
	Response.Write rRS.Collect("Sex")
	Response.Write "・" & rRS.Collect("Age") & "歳"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">住所地</th>"
	Response.Write "<td>"
	Response.Write rRS.Collect("PrefectureName") & rRS.Collect("City")
	Response.Write "</td>"
	Response.Write "</tr>"

	'<自己ＰＲ>
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">自己PR</th>"
	Response.Write "<td>"
	Response.Write Replace(ChkStr(rRS.Collect("SelfPR")), vbCrLf, "<br>")
	Response.Write "</td>"
	Response.Write "</tr>"
	'</自己ＰＲ>

	'<じぶんナビ結果>
	sComment = Array("")

	sSQL = "SELECT ResultId FROM P_MyNaviResult WITH(NOLOCK) WHERE StaffCode = '" & sStaffCode & "';"
	flgQE = QUERYEXE(dbconn, oRS2, sSQL, sError)
	If GetRSState(oRS2) = True Then
		sSQL  = "SELECT specialcomment FROM MyNavi_Result WITH(NOLOCK) WHERE id = '" & oRS2.Collect("resultid") & "';"
		flgQE = QUERYEXE(dbconn, oRS3, sSQL, sError)
		If GetRSState(oRS3) = True Then

			sSQL = "SELECT point1,point2,point3,point4,point5,point6 FROM P_MyNaviResult WITH(NOLOCK) WHERE StaffCode = '" & sStaffCode & "';"
			flgQE = QUERYEXE(dbconn, sRS, sSQL, sError)
			If GetRSState(sRS) = True Then
				Response.Write "<tr>"
				Response.Write "<th colspan=""2"">適職診断「じぶんナビ」結果<br></th>"
				Response.Write "<td>"
				Response.Write "<div style=""width:120px; float:right; font-size:10px; color:#666699; height:100px;"" class=""smartFloat"">"
				Response.Write "<embed src=""/img/order/jinzai_radar.swf?point1=" & sRS.Collect("point1") & "&point2=" & sRS.Collect("point2") & "&point3=" & sRS.Collect("point3") & "&point4=" & sRS.Collect("point4") & "&point5=" & sRS.Collect("point5") & "&point6=" & sRS.Collect("point6") & """ quality=""high"" width=""120px"" height=""100px"" bgcolor=""#ffffff"" name=""mynavi_radar"" align=""middle"" menu=""false"">"
				Response.Write "</div>"
				sComment = Split(Replace(oRS3.Collect("specialcomment"),"あなた","この方"),"<br>")
				Response.Write sComment(0)
				Response.Write "</td>"
				Response.Write "</tr>"
			End If
		End If
	End If
	'</じぶんナビ結果>

	'<最寄駅 />
	Call DspProfileNearbyStation(rDB, rRS, vUserType, vUserID)

	'<りすたーと人事担当推薦文>
	If (G_USERTYPE = "company" Or G_USERTYPE = "dispatch" Or G_USERTYPE = "talent" Or G_USERTYPE = "staff") And Left(sStaffCode,1) = "T" Then
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">人事担当推薦文</th>"
		Response.Write "<td>"
		Response.Write rRS.Collect("RecommendationLetter")
		Response.Write "</td>"
		Response.Write "</tr>"
	End If
	'</りすたーと人事担当推薦文>

	'<学歴 />
	Call DspProfileEducateHistory(rDB, rRS, vUserType, vUserID)

	'<身につけたこと>
	If ChkStr(rRS.Collect("Learn")) <> "" Then
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">学生時代に身につけたこと</th>"
		Response.Write "<td>"
		Response.Write rRS.Collect("Learn")
		Response.Write "</td>"
		Response.Write "</tr>"
	End If
	'</身につけたこと>

	Response.Write "</tbody>"
	Response.Write "</table>"
	Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">▲ページTOPへ</a></p>" & vbCrLf
End Function

'******************************************************************************
'概　要：プロフィールページの最寄駅部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'履　歴：2007/04/12 LIS K.Kokubo 作成
'　　　：2008/10/28 LIS K.Kokubo up_GetDataNearbyStation廃止→up_LstP_NearbyStationに変更
'******************************************************************************
Function DspProfileNearbyStation(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sRailwayLine
	Dim sStation

	If GetRSState(rRS) = False Then Exit Function

	'最寄駅
	sSQL = "EXEC up_LstP_NearbyStation '" & rRS.Collect("StaffCode") & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	sStation = ""
	Do While GetRSState(oRS) = True
		If sStation <> "" Then sStation = sStation & "<br>"
		sStation = sStation & oRS.Collect("StationName")
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'最寄沿線
	sSQL = "EXEC up_LstP_NearbyRailwayLine '" & rRS.Collect("StaffCode") & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	sRailwayLine = ""
	Do While GetRSState(oRS) = True
		If sRailwayLine <> "" Then sRailwayLine = sRailwayLine & "<br>"
		sRailwayLine = sRailwayLine & oRS.Collect("RailwayCompanyName") & "　" &oRS.Collect("RailwayLineName")
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	Set oRS = Nothing

	Response.Write "<tr>"
	Response.Write "<th rowspan=""2"">最寄駅</th>"
	Response.Write "<th>沿線</th>"
	Response.Write "<td>" & sRailwayLine & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>駅</th>"
	Response.Write "<td>" & sStation & "</td>"
	Response.Write "</tr>"
End Function


'******************************************************************************
'概　要：プロフィールページの最寄駅部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'履　歴：2007/04/12 LIS K.Kokubo 作成
'　　　：2008/10/28 LIS K.Kokubo up_GetDataNearbyStation廃止→up_LstP_NearbyStationに変更
'******************************************************************************
Function DspProfileNearbyStationSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sRailwayLine
	Dim sStation

	If GetRSState(rRS) = False Then Exit Function

	'最寄駅
	sSQL = "EXEC up_LstP_NearbyStation '" & rRS.Collect("StaffCode") & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	sStation = ""
	Do While GetRSState(oRS) = True
		If sStation <> "" Then sStation = sStation & "<br>"
		sStation = sStation & oRS.Collect("StationName")
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	'最寄沿線
	sSQL = "EXEC up_LstP_NearbyRailwayLine '" & rRS.Collect("StaffCode") & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	sRailwayLine = ""
	Do While GetRSState(oRS) = True
		If sRailwayLine <> "" Then sRailwayLine = sRailwayLine & "<br>"
		sRailwayLine = sRailwayLine & oRS.Collect("RailwayCompanyName") & "　" &oRS.Collect("RailwayLineName")
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	Set oRS = Nothing

	Response.Write "<tr>"
	Response.Write "<th colspan=""2""  class=""promidasi"">最寄駅</th>"
	response.write "</tr>"
	response.write "<tr>"
	Response.Write "<th>沿線</th>"
	Response.Write "<td>" & sRailwayLine & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>駅</th>"
	Response.Write "<td>" & sStation & "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'概　要：プロフィールページの学歴情報部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileEducateHistory(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim idxEducateSchool

	sSQL = "sp_GetDataEducateHistory '" & rRS.Collect("StaffCode") & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = False Then
		'学歴なし
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">学歴</th>"
		Response.Write "<td></td>"
		Response.Write "</tr>"
	Else
		'学歴あり
		idxEducateSchool = 1
		Do While GetRSState(oRS) = True
			If idxEducateSchool = 1 Then
				Response.Write "<tr>"
				Response.Write "<th rowspan=""" & oRS.RecordCount & """>学歴</th>"
			Else
				Response.Write "<tr>"
			End If
			Response.Write "<th>" & oRS.Collect("SchoolType") & "</th>"
			Response.Write "<td>"
			Response.Write oRS.Collect("Speciality")
			Response.Write "(" & Year(oRS.Collect("GraduateDay")) & "年" & oRS.Collect("GraduateType") & ")"
			Response.Write "</td>"
			Response.Write "</tr>"

			oRS.MoveNext
			idxEducateSchool = idxEducateSchool + 1
		Loop
	End If
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：プロフィールページの職歴情報部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'履　歴：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileCareerHistory(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sTableClass
	Dim sBusinessCareerMemo
	Dim iYearPeriod
	Dim iMonthPeriod
	Dim sEntryDay
	Dim sEntryYear
	Dim sEntryMonth
	Dim sRetireDay
	Dim sRetireYear
	Dim sRetireMonth
	Dim idx:	idx = 1

	If GetRSState(rRS) = False Then Exit Function

	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	sSQL = "sp_GetDataNote '" & rRS.Collect("StaffCode") & "', 'BusinessCareerMemo'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sBusinessCareerMemo = oRS.Collect("Note")
	End If
	Call RSClose(oRS)

	sSQL = "sp_GetDataCareerHistory '" & rRS.Collect("StaffCode") & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = True Then
		Response.Write "<table class=""" & sTableClass & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
		Response.Write "<colgroup>"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:400px;"">"
		Response.Write "</colgroup>"
		Response.Write "<thead>"
		Response.Write "<tr>"
		Response.Write "<th colspan=""3"">職歴</th>"
		Response.Write "</tr>"
		Response.Write "</thead>"
		Response.Write "<tbody>"

		Do While GetRSState(oRS) = True
			sEntryDay = ""
			sEntryYear = ""
			sEntryMonth = ""
			sRetireDay = ""
			sRetireYear = ""
			sRetireMonth = ""

			'就業期間計算
			iYearPeriod = Int(DateDiff("m", oRS.Collect("EntryDay"), oRS.Collect("RetireDay")) / 12)
			iMonthPeriod = DateDiff("m", oRS.Collect("EntryDay"), oRS.Collect("RetireDay")) Mod 12 + 1
			If iMonthPeriod = 12 Then
				iMonthPeriod = 0
				iYearPeriod = iYearPeriod + 1
			End If

			If IsDate(oRS.Collect("EntryDay")) = True Then
				sEntryYear = Year(oRS.Collect("EntryDay"))
				sEntryMonth = Month(oRS.Collect("EntryDay"))
				If sEntryYear & sEntryMonth <> "" Then sEntryDay = sEntryYear & "年" & sEntryMonth & "月"
			End If
			If IsDate(oRS.Collect("RetireDay")) = True Then
				sRetireYear = Year(oRS.Collect("RetireDay"))
				sRetireMonth = Month(oRS.Collect("RetireDay"))
				If sRetireYear & sRetireMonth <> "" Then sRetireDay = sRetireYear & "年" & sRetireMonth & "月"
			End If

			Response.Write "<tr>"
			Response.Write "<th rowspan=""5"">職歴" & idx & "</th>"
			Response.Write "<th>勤続期間</th>"
			Response.Write "<td>"

			If sEntryDay & sRetireDay <> "" Then Response.Write sEntryDay & "～" & sRetireDay
			If ChkStr(iYearPeriod) & ChkStr(iMonthPeriod) <> "" Then
				Response.Write "&nbsp;("
				If iYearPeriod > 0 Then Response.Write iYearPeriod & "年"
				If iMonthPeriod > 0 Then Response.Write iMonthPeriod & "ヶ月"
				Response.Write ")"
			End If

			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>業種</th>"
			Response.Write "<td>"
			Response.Write oRS.Collect("IndustryType")
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>職種</th>"
			Response.Write "<td>"
			Response.Write oRS.Collect("JobType")
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>勤務内容</th>"
			Response.Write "<td>"
			Response.Write Replace(ChkStr(oRS.Collect("BusinessDetail")), vbCrLf, "<br>")
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>勤務形態</th>"
			Response.Write "<td>"
			Response.Write oRS.Collect("WorkingType")
			Response.Write "</td>"
			Response.Write "</tr>"

			idx = idx + 1
			oRS.MoveNext
		Loop

		Response.Write "</tbody>"
		Response.Write "</table>"
		Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">▲ページTOPへ</a></p>"
	End If
	Call RSClose(oRS)
	
	If sBusinessCareerMemo <> "" Then
	
		'その他職歴
		Response.Write "<table class=""" & sTableClass & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
		Response.Write "<colgroup>"
		Response.Write "<col style=""width:200px;"">"
		Response.Write "<col style=""width:400px;"">"
		Response.Write "</colgroup>"
		Response.Write "<thead>"
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">その他</th>"
		Response.Write "</tr>"
		Response.Write "</thead>"
		Response.Write "<tbody>"
		Response.Write "<tr>"
		Response.Write "<th>職歴メモ</th>"
		Response.Write "<td>"
		Response.Write Replace(sBusinessCareerMemo,vbCrLf,"<br>")
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</tbody>"
		Response.Write "</table>"
		Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">▲ページTOPへ</a></p>"
	End If

	Response.Write vbCrLf
End Function

'******************************************************************************
'概　要：プロフィールページの職歴情報部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'履　歴：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileCareerHistorySmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sTableClass
	Dim sBusinessCareerMemo
	Dim iYearPeriod
	Dim iMonthPeriod
	Dim sEntryDay
	Dim sEntryYear
	Dim sEntryMonth
	Dim sRetireDay
	Dim sRetireYear
	Dim sRetireMonth
	Dim idx:	idx = 1

	If GetRSState(rRS) = False Then Exit Function

	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	sSQL = "sp_GetDataNote '" & rRS.Collect("StaffCode") & "', 'BusinessCareerMemo'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		sBusinessCareerMemo = oRS.Collect("Note")
	End If
	Call RSClose(oRS)

	sSQL = "sp_GetDataCareerHistory '" & rRS.Collect("StaffCode") & "'"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)

	If GetRSState(oRS) = True Then
	
	%>
		<table class="profileSmart smartBlock" style="display:none;">
		<thead>
        	<tr>
    			<th colspan="2">職歴</th>
    		</tr>
    	</thead>
        <tbody>
        <%
		
		Do While GetRSState(oRS) = True
			sEntryDay = ""
			sEntryYear = ""
			sEntryMonth = ""
			sRetireDay = ""
			sRetireYear = ""
			sRetireMonth = ""

			'就業期間計算
			iYearPeriod = Int(DateDiff("m", oRS.Collect("EntryDay"), oRS.Collect("RetireDay")) / 12)
			iMonthPeriod = DateDiff("m", oRS.Collect("EntryDay"), oRS.Collect("RetireDay")) Mod 12 + 1
			If iMonthPeriod = 12 Then
				iMonthPeriod = 0
				iYearPeriod = iYearPeriod + 1
			End If

			If IsDate(oRS.Collect("EntryDay")) = True Then
				sEntryYear = Year(oRS.Collect("EntryDay"))
				sEntryMonth = Month(oRS.Collect("EntryDay"))
				If sEntryYear & sEntryMonth <> "" Then sEntryDay = sEntryYear & "年" & sEntryMonth & "月"
			End If
			If IsDate(oRS.Collect("RetireDay")) = True Then
				sRetireYear = Year(oRS.Collect("RetireDay"))
				sRetireMonth = Month(oRS.Collect("RetireDay"))
				If sRetireYear & sRetireMonth <> "" Then sRetireDay = sRetireYear & "年" & sRetireMonth & "月"
			End If
		
		%>
			<tr>
    			<th colspan="2" class="promidasi">職歴<%= idx %></th>
            </tr>
            <tr>
            	<th>勤続期間</th>
                <td>
                <%
				
					If sEntryDay & sRetireDay <> "" Then Response.Write sEntryDay & "～" & sRetireDay
					If ChkStr(iYearPeriod) & ChkStr(iMonthPeriod) <> "" Then
						Response.Write "&nbsp;("
						If iYearPeriod > 0 Then Response.Write iYearPeriod & "年"
						If iMonthPeriod > 0 Then Response.Write iMonthPeriod & "ヶ月"
						Response.Write ")"
					End If
				
				%>
                </td>
            </tr>
            <tr>
            	<th>業種</th>
                <td><%= oRS.Collect("IndustryType") %></td>
            </tr>
            <tr>
            	<th>職種</th>
                <td><%= oRS.Collect("JobType") %></td>
            </tr>
            <tr>
            	<th>勤務内容</th>
                <td>
                <%
					Response.Write Replace(ChkStr(oRS.Collect("BusinessDetail")), vbCrLf, "<br>")
				%>
                </td>
            </tr>
            <tr>
            	<th>勤務形態</th>
                <td><%= oRS.Collect("WorkingType") %></td>
            </tr>
            <%
			
				idx = idx + 1
				oRS.MoveNext
			Loop
			
  	        %>  
    	</tbody>
	</table>   
    <%

	End If
	Call RSClose(oRS)

	If sBusinessCareerMemo <> "" Then
	%>
		<table class="profileSmart smartBlock" style="display:none;">
            <thead>
                <tr>
                    <th colspan="2">その他</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <th colspan="2" class="promidasi">職歴メモ</th>
                </tr>
                    <tr>
                    <td>
                    <%
                        Response.Write Replace(sBusinessCareerMemo,vbCrLf,"<br>")
                    %>
                    </td>
                </tr>
            </tbody>
        </table>            
	<%
	End If

	Response.Write vbCrLf
End Function

'******************************************************************************
'概　要：プロフィールページのＩＴ職歴情報部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileCareerHistoryIT(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim idx
	Dim sTempDay_it
	Dim sRangeDay_it,sTableClass

	Dim sOSLanguage	: sOSLanguage = ""
	Dim sDBTool		: sDBTool = ""

	If GetRSState(rRS) = False Then Exit Function

	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	sSQL = "sp_GetDataCareerHistoryIT '" & rRS.Collect("StaffCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then

		Response.Write "<table class=""" & sTableClass & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
		Response.Write "<colgroup>"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:400px;"">"
		Response.Write "</colgroup>"
		Response.Write "<thead>"
		Response.Write "<tr>"
		Response.Write "<th colspan=""3"">IT系職務詳細</th>"
		Response.Write "</tr>"
		Response.Write "</thead>"
		Response.Write "<tbody>"

		sSQL = "sp_GetDataDevelopmentTool '" & rRS.Collect("StaffCode") & "', '', ''"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)

		idx = 1
		Do While GetRSState(oRS) = True
			sTempDay_it = ""
			sRangeDay_it = ""
			If oRS.Collect("StartDay") <> "" Then
				sTempDay_it = GetDateStr(oRS.Collect("StartDay"), "/")
				sTempDay_it = Year(sTempDay_it) & "/" & Month(sTempDay_it)
				sRangeDay_it = sTempDay_it & "～"
			End If
			If oRS.Collect("EndDay") <> "" Then
				sTempDay_it = GetDateStr(oRS.Collect("EndDay"), "/")
				sTempDay_it = Year(sTempDay_it) & "/" & Month(sTempDay_it)
				If sRangeDay_it = "" Then sRangeDay_it = "～"
				sRangeDay_it = sRangeDay_it & sTempDay_it
			End If

			Response.Write "<tr>"
			Response.Write "<th rowspan=""7"">IT職務" & idx & "<img src=""/img/spacer.gif"" width=""1"" height=""1""></th>"
			Response.Write "<th>期間</th>"
			Response.Write "<td>"
			Response.Write sRangeDay_it
			If oRS.Collect("StartDay") <> "" And oRS.Collect("EndDay") <> "" Then
				If Int(DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) / 12) = 0 And DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 < 11 Then
					Response.Write "(" & DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 + 1 & "ヶ月)"
				ElseIf DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 = 11 Then
					Response.Write "(1年)"
				Else
					Response.Write "(" & Int(DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) / 12) & "年" & DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 + 1 & "ヶ月)"
				End If
			End If
			Response.Write "</td>"
			Response.Write "</tr>"

			Response.Write "<tr>"
			Response.Write "<th>開発内容</th>"
			Response.Write "<td>"
			Response.Write Replace(ChkStr(oRS.Collect("DevelopmentDetail")),vbCrLf,"<br>")
			Response.Write "</td>"
			Response.Write "</tr>"

			Response.Write "<tr>"
			Response.Write "<th>役割</th>"
			Response.Write "<td>"
			Dim sType1 : sType1 = ""
			If oRS.Collect("PMFlag") = "1" Then sType1 = sType1 & "　PM"
			If oRS.Collect("PLFlag") = "1" Then sType1 = sType1 & "　PL"
			If oRS.Collect("SEFlag") = "1" Then sType1 = sType1 & "　SE"
			If oRS.Collect("PGFlag") = "1" Then sType1 = sType1 & "　PG"
			If oRS.Collect("TSFlag") = "1" Then sType1 = sType1 & "　TS"
			If sType1 <> "" Then sType1 = Mid(sType1, 2)
			Response.Write sType1
			Response.Write "</td>"
			Response.Write "</tr>"

			Response.Write "<tr>"
			Response.Write "<th>作業内容</th>"
			Response.Write "<td>"

			Dim sType2 : sType2 = ""
			If oRS.Collect("SystemAnalysisFlag") = "1" Then sType2 = sType2 & "　システム分析"
			If oRS.Collect("DesignFlag") = "1" Then sType2 = sType2 & "　設計"
			If oRS.Collect("DevelopmentFlag") = "1" Then sType2 = sType2 & "　開発"
			If oRS.Collect("TestFlag") = "1" Then sType2 = sType2 & "　テスト"
			If oRS.Collect("MaintenanceFlag") = "1" Then sType2 = sType2 & "　運用保守"
			If sType2 <> "" Then sType2 = Mid(sType2, 2)
			Response.Write sType2
			Response.Write "</td>"
			Response.Write "</tr>"

			Response.Write "<tr>"
			Response.Write "<th>プロジェクト人数</th>"
			Response.Write "<td>"
			If oRS.Collect("Number") <> "" Then
				Response.Write oRS.Collect("Number") & "人"
			End If
			Response.Write "</td>"
			Response.Write "</tr>"

			Response.Write "<tr>"
			Response.Write "<th>"
			sOSLanguage=""
			sDBTool = ""
			If oRS2.State <> 0 Then
				oRS2.Filter = "CareerHistoryITID = " & idx
				Do While GetRSState(oRS2) = True
					If oRS2.Collect("CategoryCode") = "OS" _
					Or oRS2.Collect("CategoryCode") = "DevelopmentLanguage" Then
						'使用OS、言語
						If sOSLanguage <> "" Then sOSLanguage = sOSLanguage & "<br>"
						sOSLanguage = sOSLanguage & oRS2.Collect("DevelopmentToolName")
					Else
						'DB、その他
						If sDBTool <> "" Then sDBTool = sDBTool & "<br>"
						sDBTool = sDBTool & oRS2.Collect("DevelopmentToolName")
					End If
					oRS2.MoveNext
				Loop
				oRS2.Filter = 0
			End If
			Response.Write "使用OS／言語"
			Response.Write "</th>"
			Response.Write "<td>"
			Response.Write sOSLanguage
			Response.Write "</td>"
			Response.Write "</tr>"

			Response.Write "<tr>"
			Response.Write "<th>使用ツール<br>／DB／その他</th>"
			Response.Write "<td>"
			If ChkStr(oRS.Collect("DevelopmentRemark")) <> "" Then
				If sDBTool <> "" Then sDBTool = sDBTool & "<br>"
				sDBTool = sDBTool & Replace(ChkStr(oRS.Collect("DevelopmentRemark")),vbCrLf,"<br>")
			End If
			Response.Write sDBTool
			Response.Write "</td>"
			Response.Write "</tr>"

			idx = idx + 1
			oRS.MoveNext
		Loop

		Call RSClose(oRS2)
		Call RSClose(oRS)

		Response.Write "</tbody>"
		Response.Write "</table>"
		Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">▲ページTOPへ</a></p>" & vbCrLf
	End If
End Function

'******************************************************************************
'概　要：プロフィールページのＩＴ職歴情報部分を出力(smart)
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileCareerHistoryITsmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim idx
	Dim sTempDay_it
	Dim sRangeDay_it,sTableClass

	Dim sOSLanguage	: sOSLanguage = ""
	Dim sDBTool		: sDBTool = ""

	If GetRSState(rRS) = False Then Exit Function

	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	sSQL = "sp_GetDataCareerHistoryIT '" & rRS.Collect("StaffCode") & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
	
	%>
		<table class="profileSmart smartBlock" style="display:none;">
            <thead>
                <tr>
                    <th colspan="2">IT系職務詳細</th>
                </tr>
            </thead>
            <tbody>
    		<%
    			sSQL = "sp_GetDataDevelopmentTool '" & rRS.Collect("StaffCode") & "', '', ''"
				flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		
				idx = 1
				Do While GetRSState(oRS) = True
					sTempDay_it = ""
					sRangeDay_it = ""
					If oRS.Collect("StartDay") <> "" Then
						sTempDay_it = GetDateStr(oRS.Collect("StartDay"), "/")
						sTempDay_it = Year(sTempDay_it) & "/" & Month(sTempDay_it)
						sRangeDay_it = sTempDay_it & "～"
					End If
					If oRS.Collect("EndDay") <> "" Then
						sTempDay_it = GetDateStr(oRS.Collect("EndDay"), "/")
						sTempDay_it = Year(sTempDay_it) & "/" & Month(sTempDay_it)
						If sRangeDay_it = "" Then sRangeDay_it = "～"
						sRangeDay_it = sRangeDay_it & sTempDay_it
					End If
    		%>
    			<tr>
    				<th colspan="2" class="promidasi">IT職務<%= idx %></th>
    			</tr>
                <tr>
                	<th>期間</th>
                    <td>
                    <%
						If oRS.Collect("StartDay") <> "" And oRS.Collect("EndDay") <> "" Then
							If Int(DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) / 12) = 0 And DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 < 11 Then
								Response.Write "(" & DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 + 1 & "ヶ月)"
							ElseIf DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 = 11 Then
								Response.Write "(1年)"
							Else
								Response.Write "(" & Int(DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) / 12) & "年" & DateDiff("m",oRS.Collect("StartDay"),oRS.Collect("EndDay")) mod 12 + 1 & "ヶ月)"
							End If
						End If
					%>
                    </td>
                </tr>
                <tr>
                	<th colspan="2" class="itNaiyo">開発内容</th>
                </tr>
                <tr>
                	<td colspan="2">
                    <%
						Response.Write Replace(ChkStr(oRS.Collect("DevelopmentDetail")),vbCrLf,"<br>")
					%>
                    </td>
                </tr>
                <tr>
                	<th>役割</th>
                    <td>
                    <%
						Dim sType1 : sType1 = ""
						If oRS.Collect("PMFlag") = "1" Then sType1 = sType1 & "　PM"
						If oRS.Collect("PLFlag") = "1" Then sType1 = sType1 & "　PL"
						If oRS.Collect("SEFlag") = "1" Then sType1 = sType1 & "　SE"
						If oRS.Collect("PGFlag") = "1" Then sType1 = sType1 & "　PG"
						If oRS.Collect("TSFlag") = "1" Then sType1 = sType1 & "　TS"
						If sType1 <> "" Then sType1 = Mid(sType1, 2)
						Response.Write sType1
					%>
                    </td>
                </tr>
                <tr>
                	<th colspan="2" class="itNaiyo">作業内容</th>
                </tr>
                <tr>
                    <td colspan="2">
                    <%
						Dim sType2 : sType2 = ""
						If oRS.Collect("SystemAnalysisFlag") = "1" Then sType2 = sType2 & "　システム分析"
						If oRS.Collect("DesignFlag") = "1" Then sType2 = sType2 & "　設計"
						If oRS.Collect("DevelopmentFlag") = "1" Then sType2 = sType2 & "　開発"
						If oRS.Collect("TestFlag") = "1" Then sType2 = sType2 & "　テスト"
						If oRS.Collect("MaintenanceFlag") = "1" Then sType2 = sType2 & "　運用保守"
						If sType2 <> "" Then sType2 = Mid(sType2, 2)
						Response.Write sType2
					%>
                    </td>
                </tr>
                <tr>
                	<th colspan="2" class="itNaiyo">プロジェクト人数</th>
               	</tr>
                <tr>
                	<td colspan="2">
                    <%
						If oRS.Collect("Number") <> "" Then
							Response.Write oRS.Collect("Number") & "人"
						End If
					%>
                    </td>
                </tr>
                <tr>
                	<th colspan="2" class="itNaiyo">
                    <%
						sOSLanguage=""
						sDBTool = ""
						If oRS2.State <> 0 Then
							oRS2.Filter = "CareerHistoryITID = " & idx
							Do While GetRSState(oRS2) = True
								If oRS2.Collect("CategoryCode") = "OS" _
								Or oRS2.Collect("CategoryCode") = "DevelopmentLanguage" Then
									'使用OS、言語
									If sOSLanguage <> "" Then sOSLanguage = sOSLanguage & "<br>"
									sOSLanguage = sOSLanguage & oRS2.Collect("DevelopmentToolName")
								Else
									'DB、その他
									If sDBTool <> "" Then sDBTool = sDBTool & "<br>"
									sDBTool = sDBTool & oRS2.Collect("DevelopmentToolName")
								End If
								oRS2.MoveNext
							Loop
							oRS2.Filter = 0
						End If
					%>
                    使用OS／言語
                    </th>
                </tr>
                <tr>
                    <td colspan="2">
                    	<%= sOSLanguage %>
                    </td>
                </tr>
                <tr>
                	<th colspan="2" class="itNaiyo">使用ツール/DB/その他</th>
                </tr>
                <tr>
                    <td colspan="2">
                    <%
						If ChkStr(oRS.Collect("DevelopmentRemark")) <> "" Then
							If sDBTool <> "" Then sDBTool = sDBTool & "<br>"
							sDBTool = sDBTool & Replace(ChkStr(oRS.Collect("DevelopmentRemark")),vbCrLf,"<br>")
						End If
						Response.Write sDBTool
					%>
                    </td>
                </tr>
                <%
				idx = idx + 1
					oRS.MoveNext
				Loop
		
				Call RSClose(oRS2)
				Call RSClose(oRS)
				
				%>                
            </tbody>
        </table>				
	<%
			
	End If
End Function


'******************************************************************************
'概　要：プロフィールページのスキル部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'　　　：2008/07/08 LIS K.Kokubo 表示用資格名対応
'　　　：2008/12/08 LIS K.Kokubo 語学スキル対応
'******************************************************************************
Function DspProfileSkill(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim dbStaffCode
	Dim dbLanguageSeq
	Dim dbLanguageCode
	Dim dbLanguageName
	Dim dbOtherLanguage
	Dim dbLanguageActionLevelName1
	Dim dbLanguageActionLevelName2
	Dim dbLanguageActionLevelName3

	Dim sTableClass
	Dim flgLicenseData		: flgLicenseData = False
	Dim flgLanguageData		: flgLanguageData = False
	Dim flgSkillData		: flgSkillData = False
	Dim sLicense			: sLicense = ""
	Dim sOtherLicense		: sOtherLicense = ""
	Dim sLanguage			: sLanguage = ""
	Dim sOA					: sOA = ""
	Dim sOS					: sOS = ""
	Dim sApplication		: sApplication = ""
	Dim sDatabase			: sDatabase = ""
	Dim sDevelopmentLanguage: sDevelopmentLanguage = ""
	Dim iRowSkill			: iRowSkill = 0
	Dim iCol				: iCol = 2

	If GetRSState(rRS) = False Then Exit Function

	dbStaffCode = rRS.Collect("StaffCode")
	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	'<資格取得>
	sSQL = "EXEC sp_GetDataLicense '" & rRS.Collect("StaffCode") & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then flgLicenseData = True
	Do While GetRSState(oRS) = True
		If sLicense <> "" Then sLicense = sLicense & "<br>"
		sLicense = sLicense & oRS.Collect("LicenseNameDsp")
		If oRS.Collect("LicenseNameDsp") <> oRS.Collect("LicenseName") Then sLicense = sLicense & "(" & oRS.Collect("LicenseName") & ")"
		If ChkStr(oRS.Collect("GetDay")) <> "" Then sLicense = sLicense & "[" & Year(GetDateStr(oRS.Collect("GetDay"), "/")) & "年取得]"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)
	'</資格取得>

	'<その他資格取得>
	sSQL = "EXEC sp_GetDataNote '" & dbStaffCode & "', 'OtherLicense';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		flgLicenseData = True
		If sLicense <> "" Then sOtherLicense = "<hr size=""1"">"
		sOtherLicense = sOtherLicense & vbCrLf & oRS.Collect("Note")
	End If
	Call RSClose(oRS)
	'</その他資格取得>

	'<語学スキル取得>
	sSQL = "EXEC up_LstP_Skill_Language '" & dbStaffCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing
		sSQL = "EXEC up_LstP_Skill_LanguageLevel '" & dbStaffCode & "','';"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then Set oRS2.ActiveConnection = Nothing
	End If

	Do While GetRSState(oRS)
		dbLanguageSeq = oRS.Collect("LanguageSeq")
		dbLanguageCode = oRS.Collect("LanguageCode")
		dbLanguageName = oRS.Collect("LanguageName")
		dbOtherLanguage = oRS.Collect("OtherLanguage")
		flgLanguageData = True

		dbLanguageActionLevelName1 = ""
		dbLanguageActionLevelName2 = ""
		dbLanguageActionLevelName3 = ""

		If GetRSState(oRS2) = True Then
			'会話レベル
			oRS2.Filter = "LanguageSeq = '" & dbLanguageSeq & "' And LanguageActionCode = '1'"
			If GetRSState(oRS2) = True Then
				dbLanguageActionLevelName1 = oRS2.Collect("LanguageActionLevelName")
			End If
			oRS2.Filter = 0
			'読解レベル
			oRS2.Filter = "LanguageSeq = '" & dbLanguageSeq & "' And LanguageActionCode = '2'"
			If GetRSState(oRS2) = True Then
				dbLanguageActionLevelName2 = oRS2.Collect("LanguageActionLevelName")
			End If
			oRS2.Filter = 0
			'作文レベル
			oRS2.Filter = "LanguageSeq = '" & dbLanguageSeq & "' And LanguageActionCode = '3'"
			If GetRSState(oRS2) = True Then
				dbLanguageActionLevelName3 = oRS2.Collect("LanguageActionLevelName")
			End If
			oRS2.Filter = 0
		End If

		sLanguage = sLanguage & "["
		If dbLanguageCode <> "999" Then
			sLanguage = sLanguage & dbLanguageName
		Else
			sLanguage = sLanguage & dbOtherLanguage
		End If
		sLanguage = sLanguage & "]"
		sLanguage = sLanguage & "<br>"
		If dbLanguageActionLevelName1 <> "" Then sLanguage = sLanguage & "会話：" & dbLanguageActionLevelName1 & "<br>"
		If dbLanguageActionLevelName2 <> "" Then sLanguage = sLanguage & "読解：" & dbLanguageActionLevelName2 & "<br>"
		If dbLanguageActionLevelName3 <> "" Then sLanguage = sLanguage & "作文：" & dbLanguageActionLevelName3 & "<br>"

		oRS.MoveNext
		If GetRSState(oRS) = True Then sLanguage = sLanguage & "<div class=""line1""></div>"
	Loop
	Call RSClose(oRS)
	Call RSClose(oRS2)
	'</語学スキル取得>

	'<スキル取得>
	sSQL = "EXEC sp_GetDataSkill '" & dbStaffCode & "', '';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		flgSkillData = True

		'OS
		oRS.Filter = "CategoryCode = 'OS'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sOS <> "" Then sOS = sOS & "<br>"
			sOS = sOS & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sOS = sOS & "(" & oRS.Collect("Period") & "年)"
			End If
			oRS.MoveNext
		Loop
		oRS.Filter = 0
		'If sOS = "" Then sOS = "ＯＳ経験なし"

		'Application
		oRS.Filter = "CategoryCode = 'Application'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sApplication <> "" Then sApplication = sApplication & "<br>"
			sApplication = sApplication & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sApplication = sApplication & "(" & oRS.Collect("Period") & "年)"
			End If
			oRS.MoveNext
		Loop
		oRS.Filter = 0
		'If sApplication = "" Then sApplication = "アプリケーション経験なし"

		'Database
		oRS.Filter = "CategoryCode = 'Database'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sDatabase <> "" Then sDatabase = sDatabase & "<br>"
			sDatabase = sDatabase & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sDatabase = sDatabase & "(" & oRS.Collect("Period") & "年)"
			End If
			oRS.MoveNext
		Loop
		oRS.Filter = 0
		'If sDatabase = "" Then sDatabase = "データベース経験なし"

		'DevelopmentLanguage
		oRS.Filter = "CategoryCode = 'DevelopmentLanguage'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sDevelopmentLanguage <> "" Then sDevelopmentLanguage = sDevelopmentLanguage & "<br>"
			sDevelopmentLanguage = sDevelopmentLanguage & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sDevelopmentLanguage = sDevelopmentLanguage & "(" & oRS.Collect("Period") & "年)"
			End If
			oRS.MoveNext
		Loop
		'If sDevelopmentLanguage = "" Then sDevelopmentLanguage = "開発言語経験なし"
	End If

	Call RSClose(oRS)
	'</スキル取得>

	If flgSkillData = True Then iCol = iCol + 1
	If flgLicenseData = True Or flgLanguageData = True Or flgSkillData = True Then
		
		%>
		<table class="profileSmart smartBlock" style="display:none;">
            <thead>
                <tr>
                    <th colspan="2">資格・スキル</th>
                </tr>
            </thead>
			<tbody>
            <% If flgLicenseData = True Then %>
            	<tr>
            		<th>資格</th>
            		<td>
                    <%= sLicense %>
                    <%= sOtherLicense %>
                    </td>
            	</tr>
            <% End If %>
            <% If flgLanguageData = True Then %>
            	<tr>
            		<th>語学</th>
            		<td>
                    <% sLanguage %>
                    </td>
            	</tr>
            <% End If %>
            <% If flgSkillData = True Then %>
            	<tr>
            		<th colspan="2" class="promidasi">全保有スキル</th>
                </tr>
                <tr>
                	<th>OS名</th>
                    <td><%= sOS %></td>
                </tr>
                <tr>
                	<th>アプリケーション名</th>
                	<td><%= sApplication %></td>
                </tr>
                <tr>
                	<th>開発言語名</th>
                	<td><%= sDevelopmentLanguage %></td>
                </tr>
                <tr>
                	<th>データベース名</th>
                	<td><%= sDatabase %></td>
                </tr>
            <% End If %>
            </tbody>
		</table>
        <%
		
		Response.Write "<table class=""" & sTableClass & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
		Response.Write "<colgroup>"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:400px;"">"
		Response.Write "</colgroup>"
		Response.Write "<thead>"
		Response.Write "<tr>"
		Response.Write "<th colspan=""" & iCol & """>資格・スキル</th>"
		Response.Write "</tr>"
		Response.Write "</thead>"
		Response.Write "<tbody>"

		If flgLicenseData = True Then
			Response.Write "<tr>"
			Response.Write "<th colspan=""" & iCol - 1 & """>資格</th>"
			Response.Write "<td>"
			Response.Write sLicense
			Response.Write sOtherLicense
			Response.Write "</td>"
			Response.Write "</tr>"
		End If

		If flgLanguageData = True Then
			Response.Write "<tr>"
			Response.Write "<th colspan=""" & iCol - 1 & """>語学</th>"
			Response.Write "<td>"
			Response.Write sLanguage
			Response.Write "</td>"
			Response.Write "</tr>"
		End If

		If flgSkillData = True Then
			Response.Write "<tr>"
			Response.Write "<th rowspan=""4"">全保有<br>スキル</th>"
			Response.Write "<th>OS名</th>"
			Response.Write "<td>"
			Response.Write sOS
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th style=""font-size:10px;"">アプリケーション名</th>"
			Response.Write "<td>"
			Response.Write sApplication
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>開発言語名</th>"
			Response.Write "<td>"
			Response.Write sDevelopmentLanguage
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>データベース名</th>"
			Response.Write "<td>"
			Response.Write sDatabase
			Response.Write "</td>"
			Response.Write "</tr>"
		End If

		Response.Write "</body>"
		Response.Write "</table>"
		Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">▲ページTOPへ</a></p>"
	End If
End Function

'******************************************************************************
'概　要：プロフィールページのスキル部分を出力(簡易登録)
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2012/06/07 タニザワ作成
'******************************************************************************
Function DspProfileSkillSimple(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim oRS2
	Dim flgQE
	Dim sError

	Dim dbStaffCode
	Dim dbLanguageSeq
	Dim dbLanguageCode
	Dim dbLanguageName
	Dim dbOtherLanguage
	Dim dbLanguageActionLevelName1
	Dim dbLanguageActionLevelName2
	Dim dbLanguageActionLevelName3

	Dim sTableClass
	Dim flgLicenseData		: flgLicenseData = False
	Dim flgLanguageData		: flgLanguageData = False
	Dim flgSkillData		: flgSkillData = False
	Dim sLicense			: sLicense = ""
	Dim sOtherLicense		: sOtherLicense = ""
	Dim sLanguage			: sLanguage = ""
	Dim sOA					: sOA = ""
	Dim sOS					: sOS = ""
	Dim sApplication		: sApplication = ""
	Dim sDatabase			: sDatabase = ""
	Dim sDevelopmentLanguage: sDevelopmentLanguage = ""
	Dim iRowSkill			: iRowSkill = 0
	Dim iCol				: iCol = 2

	If GetRSState(rRS) = False Then Exit Function

	dbStaffCode = rRS.Collect("StaffCode")
	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	'<資格取得>
	sSQL = "EXEC sp_GetDataLicense_Simple '" & rRS.Collect("StaffCode") & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then flgLicenseData = True
	Do While GetRSState(oRS) = True
		If sLicense <> "" Then sLicense = sLicense & "<br>"
		sLicense = sLicense & oRS.Collect("LicenseNameDsp")
		If oRS.Collect("LicenseNameDsp") <> oRS.Collect("LicenseName") Then sLicense = sLicense & "(" & oRS.Collect("LicenseName") & ")"
		If ChkStr(oRS.Collect("GetDay")) <> "" Then sLicense = sLicense & "[" & Year(GetDateStr(oRS.Collect("GetDay"), "/")) & "年取得]"

		oRS.MoveNext
	Loop
	Call RSClose(oRS)
	'</資格取得>

	'<その他資格取得>
	sSQL = "EXEC sp_GetDataNote '" & dbStaffCode & "', 'OtherLicense';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		flgLicenseData = True
		If sLicense <> "" Then sOtherLicense = "<hr size=""1"">"
		sOtherLicense = sOtherLicense & vbCrLf & oRS.Collect("Note")
	End If
	Call RSClose(oRS)
	'</その他資格取得>

	'<語学スキル取得>
	sSQL = "EXEC up_LstP_Skill_Language '" & dbStaffCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		Set oRS.ActiveConnection = Nothing
		sSQL = "EXEC up_LstP_Skill_LanguageLevel '" & dbStaffCode & "','';"
		flgQE = QUERYEXE(rDB, oRS2, sSQL, sError)
		If GetRSState(oRS2) = True Then Set oRS2.ActiveConnection = Nothing
	End If

	Do While GetRSState(oRS)
		dbLanguageSeq = oRS.Collect("LanguageSeq")
		dbLanguageCode = oRS.Collect("LanguageCode")
		dbLanguageName = oRS.Collect("LanguageName")
		dbOtherLanguage = oRS.Collect("OtherLanguage")
		flgLanguageData = True

		dbLanguageActionLevelName1 = ""
		dbLanguageActionLevelName2 = ""
		dbLanguageActionLevelName3 = ""

		If GetRSState(oRS2) = True Then
			'会話レベル
			oRS2.Filter = "LanguageSeq = '" & dbLanguageSeq & "' And LanguageActionCode = '1'"
			If GetRSState(oRS2) = True Then
				dbLanguageActionLevelName1 = oRS2.Collect("LanguageActionLevelName")
			End If
			oRS2.Filter = 0
			'読解レベル
			oRS2.Filter = "LanguageSeq = '" & dbLanguageSeq & "' And LanguageActionCode = '2'"
			If GetRSState(oRS2) = True Then
				dbLanguageActionLevelName2 = oRS2.Collect("LanguageActionLevelName")
			End If
			oRS2.Filter = 0
			'作文レベル
			oRS2.Filter = "LanguageSeq = '" & dbLanguageSeq & "' And LanguageActionCode = '3'"
			If GetRSState(oRS2) = True Then
				dbLanguageActionLevelName3 = oRS2.Collect("LanguageActionLevelName")
			End If
			oRS2.Filter = 0
		End If

		sLanguage = sLanguage & "["
		If dbLanguageCode <> "999" Then
			sLanguage = sLanguage & dbLanguageName
		Else
			sLanguage = sLanguage & dbOtherLanguage
		End If
		sLanguage = sLanguage & "]"
		sLanguage = sLanguage & "<br>"
		If dbLanguageActionLevelName1 <> "" Then sLanguage = sLanguage & "会話：" & dbLanguageActionLevelName1 & "<br>"
		If dbLanguageActionLevelName2 <> "" Then sLanguage = sLanguage & "読解：" & dbLanguageActionLevelName2 & "<br>"
		If dbLanguageActionLevelName3 <> "" Then sLanguage = sLanguage & "作文：" & dbLanguageActionLevelName3 & "<br>"

		oRS.MoveNext
		If GetRSState(oRS) = True Then sLanguage = sLanguage & "<div class=""line1""></div>"
	Loop
	Call RSClose(oRS)
	Call RSClose(oRS2)
	'</語学スキル取得>

	'<スキル取得>
	sSQL = "EXEC sp_GetDataSkill_Simple '" & dbStaffCode & "', '';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		flgSkillData = True

		'OS
		oRS.Filter = "CategoryCode = 'OS'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sOS <> "" Then sOS = sOS & "<br>"
			sOS = sOS & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sOS = sOS & "(" & oRS.Collect("Period") & "年)"
			End If
			oRS.MoveNext
		Loop
		oRS.Filter = 0
		'If sOS = "" Then sOS = "ＯＳ経験なし"

		'Application
		oRS.Filter = "CategoryCode = 'Application'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sApplication <> "" Then sApplication = sApplication & "<br>"
			sApplication = sApplication & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sApplication = sApplication & "(" & oRS.Collect("Period") & "年)"
			End If
			oRS.MoveNext
		Loop
		oRS.Filter = 0
		'If sApplication = "" Then sApplication = "アプリケーション経験なし"

		'Database
		oRS.Filter = "CategoryCode = 'Database'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sDatabase <> "" Then sDatabase = sDatabase & "<br>"
			sDatabase = sDatabase & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sDatabase = sDatabase & "(" & oRS.Collect("Period") & "年)"
			End If
			oRS.MoveNext
		Loop
		oRS.Filter = 0
		'If sDatabase = "" Then sDatabase = "データベース経験なし"

		'DevelopmentLanguage
		oRS.Filter = "CategoryCode = 'DevelopmentLanguage'"
		If GetRSState(oRS) = True Then iRowSkill = iRowSkill + 1
		Do While GetRSState(oRS) = True
			If sDevelopmentLanguage <> "" Then sDevelopmentLanguage = sDevelopmentLanguage & "<br>"
			sDevelopmentLanguage = sDevelopmentLanguage & oRS.Collect("SkillName")
			If oRS.Collect("Period") <> "" Then
				sDevelopmentLanguage = sDevelopmentLanguage & "(" & oRS.Collect("Period") & "年)"
			End If
			oRS.MoveNext
		Loop
		'If sDevelopmentLanguage = "" Then sDevelopmentLanguage = "開発言語経験なし"
	End If

	Call RSClose(oRS)
	'</スキル取得>

	If flgSkillData = True Then iCol = iCol + 1
	If flgLicenseData = True Or flgLanguageData = True Or flgSkillData = True Then
		%>
		<table class="profileSmart smartBlock" style="display:none;">
            <thead>
                <tr>
                    <th colspan="2">資格・スキル</th>
                </tr>
            </thead>
			<tbody>
            <% If flgLicenseData = True Then %>
            	<tr>
            		<th>資格</th>
            		<td>
                    <%= sLicense %>
                    <%= sOtherLicense %>
                    </td>
            	</tr>
            <% End If %>
            <% If flgLanguageData = True Then %>
            	<tr>
            		<th>語学</th>
            		<td>
                    <%= sLanguage %>
                    </td>
            	</tr>
            <% End If %>
            <% If flgSkillData = True Then %>
            	<tr>
            		<th colspan="2" class="promidasi">全保有スキル</th>
                </tr>
                <tr>
                	<th>OS名</th>
                    <td><%= sOS %></td>
                </tr>
                <tr>
                	<th>アプリケーション名</th>
                	<td><%= sApplication %></td>
                </tr>
                <tr>
                	<th>開発言語名</th>
                	<td><%= sDevelopmentLanguage %></td>
                </tr>
                <tr>
                	<th>データベース名</th>
                	<td><%= sDatabase %></td>
                </tr>
            <% End If %>
            </tbody>
		</table>
        <%
	
	
		Response.Write "<table class=""" & sTableClass & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
		Response.Write "<colgroup>"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:100px;"">"
		Response.Write "<col style=""width:400px;"">"
		Response.Write "</colgroup>"
		Response.Write "<thead>"
		Response.Write "<tr>"
		Response.Write "<th colspan=""" & iCol & """>資格・スキル</th>"
		Response.Write "</tr>"
		Response.Write "</thead>"
		Response.Write "<tbody>"

		If flgLicenseData = True Then
			Response.Write "<tr>"
			Response.Write "<th colspan=""" & iCol - 1 & """>資格</th>"
			Response.Write "<td>"
			Response.Write sLicense
			Response.Write sOtherLicense
			Response.Write "</td>"
			Response.Write "</tr>"
		End If

		If flgLanguageData = True Then
			Response.Write "<tr>"
			Response.Write "<th colspan=""" & iCol - 1 & """>語学</th>"
			Response.Write "<td>"
			Response.Write sLanguage
			Response.Write "</td>"
			Response.Write "</tr>"
		End If

		If flgSkillData = True Then
			Response.Write "<tr>"
			Response.Write "<th rowspan=""4"">全保有<br>スキル</th>"
			Response.Write "<th>OS名</th>"
			Response.Write "<td>"
			Response.Write sOS
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th style=""font-size:10px;"">アプリケーション名</th>"
			Response.Write "<td>"
			Response.Write sApplication
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>開発言語名</th>"
			Response.Write "<td>"
			Response.Write sDevelopmentLanguage
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<th>データベース名</th>"
			Response.Write "<td>"
			Response.Write sDatabase
			Response.Write "</td>"
			Response.Write "</tr>"
		End If

		Response.Write "</body>"
		Response.Write "</table>"
		Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">▲ページTOPへ</a></p>"
	End If
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHope(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sPattern

	If GetRSState(rRS) = False Then Exit Function

	If vUserType = "staff" Then
		sPattern = "pattern9"
	ElseIf vUserType = "company" Then
		sPattern = "pattern8"
	End If

	Response.Write "<table class=""" & sPattern & " cw smartNone"" border=""0"" style=""margin-bottom:15px;"">"
	Response.Write "<colgroup>"
	Response.Write "<col style=""width:100px;"">"
	Response.Write "<col style=""width:100px;"">"
	Response.Write "<col style=""width:400px;"">"
	Response.Write "</colgroup>"
	Response.Write "<thead>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""3"">希望条件</th>"
	Response.Write "</tr>"
	Response.Write "</thead>"
	Response.Write "<tbody>"

	'希望勤務形態
	Call DspProfileHopeWorkingType(rDB, rRS, vUserType, vUserID)
    '希望勤務形態
	Call DspProfileHopeContactableTime(rDB, rRS, vUserType, vUserID)
	'希望業種
	Call DspProfileHopeIndustry(rDB, rRS, vUserType, vUserID)
	'希望職種
	Call DspProfileHopeJobType(rDB, rRS, vUserType, vUserID)
    '副業
    Call DspProfileSideJob(rDB, rRS, vUserType, vUserID)
	'希望勤務地
	Call DspProfileHopeWorkingPlace(rDB, rRS, vUserType, vUserID)
	'給与条件
	Call DspProfileHopeSalary(rDB, rRS, vUserType, vUserID)
	'期間・時間
	Call DspProfileHopeSpan(rDB, rRS, vUserType, vUserID)
	'福利厚生
	Call DspProfileHopeWelfare(rDB, rRS, vUserType, vUserID)

	Response.Write "</tbody>"
	Response.Write "</table>"
	Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">▲ページTOPへ</a></p>" & vbCrLf
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件勤務形態部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeWorkingType(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim sStaffCode
	Dim sWorkingType	: sWorkingType = ""

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sSQL = "sp_GetDataWorkingType '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sWorkingType = sWorkingType & oRS.Collect("WorkingTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then sWorkingType = sWorkingType & "　"
	Loop
	Call RSClose(oRS)

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">勤務形態</th>"
	Response.Write "<td>"
	Response.Write sWorkingType
	Response.Write "</td>"
	Response.Write "</tr>"

	If G_USERTYPE = "dispatch" Then
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">正社員紹介予定派遣の希望</th>"
		Response.Write "<td>"

		If rRS.Collect("TempToPermFlag") = "1" Then
			Response.Write "希望する"
		Else
			Response.Write "希望しない"
		End If

		Response.Write "</td>"
		Response.Write "</tr>"
	End If
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件業種部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeIndustry(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">業種</th>"
	Response.Write "<td>"

	sSQL = "sp_GetDataIndustryType '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		Response.Write oRS.Collect("IndustryTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then Response.Write "<br>"
	Loop
	Call RSClose(oRS)

	Response.Write "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件業種部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeJobType(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">職種</th>"
	Response.Write "<td>"

	sSQL = "sp_GetDataHopeJobType '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		Response.Write oRS.Collect("JobTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then Response.Write "<br>"
	Loop
	Call RSClose(oRS)

	Response.Write "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件勤務地部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeWorkingPlace(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode

	Dim flgCommuteTime	: flgCommuteTime = False
	Dim flgStation		: flgStation = False
	Dim sArea			: sArea = ""
	Dim sPlace			: sPlace = ""
	Dim sHopeCommuteTime: sHopeCommuteTime = ""
	Dim sStation		: sStation = ""
	Dim sRailwayLine	: sRailwayLine = ""
	Dim iRow			: iRow = 1

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sSQL = "sp_GetDataHopeWorkingPlace '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then iRow = iRow + 1
	Do While GetRSState(oRS) = True
		If InStr(sArea, oRS.Collect("Area")) = 0 Then
			'重複しないエリア名のみ
			sArea = sArea & "　" & oRS.Collect("Area")
		End If
		sPlace = sPlace & oRS.Collect("WorkingPlace")
		oRS.MoveNext
		If GetRSState(oRS) = True Then sPlace = sPlace & "<br>"
	Loop
	Call RSClose(oRS)

	'希望駅
	sSQL = "sp_GetDataHopeCommuting '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		If sHopeCommuteTime <> "" Then sHopeCommuteTime = sHopeCommuteTime & "<br>"
		If ChkStr(oRS.Collect("HopeCommuteTime")) <> "" Then
			flgCommuteTime = True
			sHopeCommuteTime = oRS.Collect("HopeCommuteTime") & "分"
		End If

		If ChkStr(oRS.Collect("StationName")) <> "" Then
			flgStation = True
			If sStation <> "" Then sStation = sStation & "<br>"
			sStation = sStation & oRS.Collect("StationName") & "駅"

			If oRS.Collect("MinuteToStation") <> "" Then
				sStation = sStation & "(駅から" & oRS.Collect("MinuteToStation") & "分以内)"
			End If
		End If
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	If flgCommuteTime = True Then iRow = iRow + 1
	If flgStation = True Then iRow = iRow + 1

	'希望沿線
'	sSQL = "sp_GetDataHopeRailwayLine '" & sStaffCode & "'"
'	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
'	If GetRSState(oRS) = True Then iRow = iRow + 1
'	Do While GetRSState(oRS) = True
'		sRailwayLine = sRailwayLine & _
'			oRS.Collect("RailwayCompanyName") & _
'			"　" & oRS.Collect("RailwayLineName") & "<br>"
'		oRS.MoveNext
'		If GetRSState(oRS) = True Then sRailwayLine = sRailwayLine & "<br>"
'	Loop
'	Call RSClose(oRS)

	Response.Write "<tr>"
	Response.Write "<th rowspan=""" & iRow & """>勤務地</th>"
	Response.Write "<th>希望国</th>"
	Response.Write "<td>" & sArea & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>希望勤務地</th>"
	Response.Write "<td>" & sPlace & "</td>"
	Response.Write "</tr>"

	If sHopeCommuteTime <> "" Then
		Response.Write "<tr>"
		Response.Write "<th>希望通勤時間</th>"
		Response.Write "<td>" & sHopeCommuteTime & "</td>"
		Response.Write "</tr>"
	End If

	If sStation <> "" Then
		Response.Write "<tr>"
		Response.Write "<th>希望駅</th>"
		Response.Write "<td>" & sStation & "</td>"
		Response.Write "</tr>"
	End If
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件給与部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeSalary(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sStaffCode
	Dim YMin
	Dim YMax
	Dim MMin
	Dim MMax
	Dim DMin
	Dim DMax
	Dim HMin
	Dim HMax
	Dim PercentagePay
	Dim Remark
	Dim iRow

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	If ChkStr(rRS.Collect("YearlyIncomeMin")) <> "" Then YMin = GetJapaneseYen(rRS.Collect("YearlyIncomeMin"))
	If ChkStr(rRS.Collect("YearlyIncomeMax")) <> "" Then YMax = GetJapaneseYen(rRS.Collect("YearlyIncomeMax"))
	If ChkStr(rRS.Collect("MonthlyIncomeMin")) <> "" Then MMin = GetJapaneseYen(rRS.Collect("MonthlyIncomeMin"))
	If ChkStr(rRS.Collect("MonthlyIncomeMax")) <> "" Then MMax = GetJapaneseYen(rRS.Collect("MonthlyIncomeMax"))
	If ChkStr(rRS.Collect("DailyIncomeMin")) <> "" Then DMin = GetJapaneseYen(rRS.Collect("DailyIncomeMin"))
	If ChkStr(rRS.Collect("DailyIncomeMax")) <> "" Then DMax = GetJapaneseYen(rRS.Collect("DailyIncomeMax"))
	If ChkStr(rRS.Collect("HourlyIncomeMin")) <> "" Then HMin = GetJapaneseYen(rRS.Collect("HourlyIncomeMin"))
	If ChkStr(rRS.Collect("HourlyIncomeMax")) <> "" Then HMax = GetJapaneseYen(rRS.Collect("HourlyIncomeMax"))
	PercentagePay = ChkStr(rRS.Collect("PercentagePayFlag"))
	Remark = ChkStr(rRS.Collect("IncomeRemark"))

	iRow = 0
	If YMin & YMax <> "" Then iRow = iRow + 1
	If MMin & MMax <> "" Then iRow = iRow + 1
	If DMin & DMax <> "" Then iRow = iRow + 1
	If HMin & HMax <> "" Then iRow = iRow + 1
	If PercentagePay <> "" Then iRow = iRow + 1
	If Remark <> "" Then iRow = iRow + 1

	If iRow > 0 Then
		Response.Write "<tr>"
		Response.Write "<th rowspan=""" & iRow & """>希望給与</th>"
		If YMin & YMax <> "" Then
			Response.Write "<th>年収</th>"
			Response.Write "<td>" & YMin & "～" & YMax &"</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		If MMin & MMax <> "" Then
			Response.Write "<th>月収</th>"
			Response.Write "<td>" & MMin & "～" & MMax & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		If DMin & DMax <> "" Then
			Response.Write "<th>日給</th>"
			Response.Write "<td>" & DMin & "～" & DMax & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		If HMin & HMax <> "" Then
			Response.Write "<th>時給</th>"
			Response.Write "<td>" & HMin & "～" & HMax & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		Response.Write "<th>歩合</th>"
		Response.Write "<td>"
		If PercentagePay = "1" Then
			Response.Write "希望する"
		ElseIf PercentagePay = "0" Then
			Response.Write "希望しない"
		Else
			Response.Write "こだわらない"
		End If
		Response.Write "</td>"
		Response.Write "</tr>"

		If Remark <> "" Then
			Response.Write "<tr>"
			Response.Write "<th>備考</th>"
			Response.Write "<td>" & Remark & "</td>"
			Response.Write "</tr>"
		End If
	Else
		'希望給与条件が無い場合
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">希望給与</th>"
		Response.Write "<td></td>"
		Response.Write "</tr>"
	End If
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件期間・時間部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeSpan(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sStaffCode
	Dim sTransfer
	Dim sWorkPeriod
	Dim sWorkPeriodFlag
	Dim sWorkMonthPeriod
	Dim sWorkTime
	Dim sWSTime
	Dim sWETime
	Dim sOverWork
	Dim sOverWorkTimeMax
	Dim sWorkShift
	Dim sHoliday

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sTransfer = ChkStr(rRS.Collect("TransferFlag"))
	Select Case sTransfer
		Case "1":	sTransfer = "可"
		Case "0":	sTransfer = "不可"
		Case Else:	sTransfer = "こだわらない"
	End Select

	sWorkPeriod = ChkStr(oRS.Collect("WorkPeriod"))
	sWorkPeriodFlag = ChkStr(oRS.Collect("WorkPeriodFlag"))
	sWorkMonthPeriod = ChkStr(oRS.Collect("WorkMonthPeriod"))
	sWSTime = ChkStr(oRS.Collect("WorkStartTime"))
	sWETime = ChkStr(oRS.Collect("WorkEndTime"))
	If sWSTime & sWETime <> "" Then
		If sWSTime <> "" Then sWorkTime = sWorkTime & Left(sWSTime, 2) & ":" & Right(sWSTime, 2)
		sWorkTime = sWorkTime & "～"
		If sWETime <> "" Then sWorkTime = sWorkTime & Left(sWETime, 2) & ":" & Right(sWETime, 2)
	End If

	sOverWork = ChkStr(rRS.Collect("OverWorkFlag"))
	Select Case sOverWork
		Case "1":	sOverWork = "可"
		Case "0":	sOverWork = "不可"
		Case Else:	sOverWork = "こだわらない"
	End Select
	sOverWorkTimeMax = ChkStr(rRS.Collect("OverWorkTimeMax"))
	If sOverWorkTimeMax <> "" Then sOverWorkTimeMax = Left(sOverWorkTimeMax, 2) & ":" & Right(sOverWorkTimeMax, 2)

	sWorkShift = ChkStr(rRS.Collect("WorkShiftFlag"))
	Select Case sWorkShift
		Case "1":	sWorkShift = "可"
		Case "0":	sWorkShift = "不可"
		Case Else:	sWorkShift = "こだわらない"
	End Select

	sHoliday = ""
	If oRS.Collect("MonHolidayFlag") = "1" Then sHoliday = sHoliday & "月"
	If oRS.Collect("TueHolidayFlag") = "1" Then sHoliday = sHoliday & "火"
	If oRS.Collect("WedHolidayFlag") = "1" Then sHoliday = sHoliday & "水"
	If oRS.Collect("ThuHolidayFlag") = "1" Then sHoliday = sHoliday & "木"
	If oRS.Collect("FriHolidayFlag") = "1" Then sHoliday = sHoliday & "金"
	If oRS.Collect("SatHolidayFlag") = "1" Then sHoliday = sHoliday & "土"
	If oRS.Collect("SunHolidayFlag") = "1" Then sHoliday = sHoliday & "日"
	If oRS.Collect("PublicHolidayFlag") = "1" Then sHoliday = sHoliday & "祝"
	If ChkStr(oRS.Collect("WeeklyHolidayType")) <> "" Then
		If sHoliday <> "" Then sHoliday = sHoliday & "<br>"
		sHoliday = sHoliday & oRS.Collect("WeeklyHolidayType")
	End If
	If ChkStr(oRS.Collect("HolidayRemark")) <> "" Then
		If sHoliday <> "" Then sHoliday = sHoliday & "<br>"
		sHoliday = sHoliday & oRS.Collect("HolidayRemark")
	End If

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">転勤可／不可</th>"
	Response.Write "<td>" & sTransfer & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">勤務期間</th>"
	Response.Write "<td>"
	Response.Write sWorkPeriod

	If sWorkPeriodFlag = "3" And sWorkMonthPeriod <> "" Then
		'短期
		Response.Write " ( " & sWorkMonthPeriod & "ヶ月 )"
	End If

	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">就業時間</th>"
	Response.Write "<td>" & sWorkTime & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">残業</th>"
	Response.Write "<td>"
	Response.Write sOverWork

	If sOverWorkTimeMax <> "" Then
		Response.Write "(" & sOverWorkTimeMax & "まで)"
	End If

	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">シフト(交代)稼働</th>"
	Response.Write "<td>" & sWorkShift & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">休日</th>"
	Response.Write "<td>" & sHoliday & "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件福利厚生部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeWelfare(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sStaffCode
	Dim sWelfareProgram

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sWelfareProgram = ""
	If rRS.Collect("TrafficFeeFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　交通費支給"
	If rRS.Collect("SocietyInsuranceFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　社会保険完備"
	If rRS.Collect("SanatoriumFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　保養所"
	If rRS.Collect("EnterprisePensionFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　企業年金"
	If rRS.Collect("WealthShapeFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　財形貯蓄"
	If rRS.Collect("StockOptionFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　持株制度(ストックオプション)"
	If rRS.Collect("RetirementPayFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　退職金制度"
	If rRS.Collect("ResidencePayFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　住宅手当"
	If rRS.Collect("FamilyPayFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　家族手当"
	If rRS.Collect("EmployeeDormitoryFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　社員寮"
	If rRS.Collect("CompanyHouseFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　社宅"
	If rRS.Collect("NewEmployeeTrainingFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　新入社員研修"
	If rRS.Collect("OverseasTrainingFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　海外研修"
	If rRS.Collect("OtherTrainingFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　各種研修"
	If rRS.Collect("FlexTimeFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　フレックスタイム"

	If sWelfareProgram <> "" Then
		sWelfareProgram = Mid(sWelfareProgram, 2)

		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">福利厚生</th>"
		Response.Write "<td>" & sWelfareProgram & "</td>"
		Response.Write "</tr>"
	End If
End Function



'******************************************************************************
'概　要：プロフィールページの希望条件部分を出力(smart)
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sPattern

	If GetRSState(rRS) = False Then Exit Function

	If vUserType = "staff" Then
		sPattern = "pattern9"
	ElseIf vUserType = "company" Then
		sPattern = "pattern8"
	End If

	%>
	<table class="profileSmart smartBlock" style="display:none;">
        <thead>
            <tr>
                <th colspan="2">希望条件</th>
            </tr>
        </thead>
        <tbody>
	
	<%

	'希望勤務形態
	Call DspProfileHopeWorkingTypeSmart(rDB, rRS, vUserType, vUserID)
	'希望業種
	Call DspProfileHopeIndustrySmart(rDB, rRS, vUserType, vUserID)
	'希望職種
	Call DspProfileHopeJobTypeSmart(rDB, rRS, vUserType, vUserID)
	'希望勤務地
	Call DspProfileHopeWorkingPlaceSmart(rDB, rRS, vUserType, vUserID)
	'給与条件
	Call DspProfileHopeSalarySmart(rDB, rRS, vUserType, vUserID)
	'期間・時間
	Call DspProfileHopeSpanSmart(rDB, rRS, vUserType, vUserID)
	'福利厚生
	Call DspProfileHopeWelfareSmart(rDB, rRS, vUserType, vUserID)

	Response.Write "</tbody>"
	Response.Write "</table>" & vbCrLf
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件勤務形態部分を出力(Smart)
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeWorkingTypeSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim sStaffCode
	Dim sWorkingType	: sWorkingType = ""

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sSQL = "sp_GetDataWorkingType '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sWorkingType = sWorkingType & oRS.Collect("WorkingTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then sWorkingType = sWorkingType & "　"
	Loop
	Call RSClose(oRS)

	Response.Write "<tr>"
	Response.Write "<th>勤務形態</th>"
	Response.Write "<td>"
	Response.Write sWorkingType
	Response.Write "</td>"
	Response.Write "</tr>"

	If G_USERTYPE = "dispatch" Then
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"" class=""itNaiyo"">正社員紹介予定派遣の希望</th>"
		response.write "</tr><tr>"
		Response.Write "<td colspan=""2"">"

		If rRS.Collect("TempToPermFlag") = "1" Then
			Response.Write "希望する"
		Else
			Response.Write "希望しない"
		End If

		Response.Write "</td>"
		Response.Write "</tr>"
	End If
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件業種部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeIndustrySmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	Response.Write "<tr>"
	Response.Write "<th>業種</th>"
	Response.Write "<td>"

	sSQL = "sp_GetDataIndustryType '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		Response.Write oRS.Collect("IndustryTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then Response.Write "<br>"
	Loop
	Call RSClose(oRS)

	Response.Write "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件業種部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeJobTypeSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	Response.Write "<tr>"
	Response.Write "<th>職種</th>"
	Response.Write "<td>"

	sSQL = "sp_GetDataHopeJobType '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		Response.Write oRS.Collect("JobTypeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then Response.Write "<br>"
	Loop
	Call RSClose(oRS)

	Response.Write "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件勤務地部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeWorkingPlaceSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode

	Dim flgCommuteTime	: flgCommuteTime = False
	Dim flgStation		: flgStation = False
	Dim sArea			: sArea = ""
	Dim sPlace			: sPlace = ""
	Dim sHopeCommuteTime: sHopeCommuteTime = ""
	Dim sStation		: sStation = ""
	Dim sRailwayLine	: sRailwayLine = ""
	Dim iRow			: iRow = 1

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sSQL = "sp_GetDataHopeWorkingPlace '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then iRow = iRow + 1
	Do While GetRSState(oRS) = True
		If InStr(sArea, oRS.Collect("Area")) = 0 Then
			'重複しないエリア名のみ
			sArea = sArea & "　" & oRS.Collect("Area")
		End If
		sPlace = sPlace & oRS.Collect("WorkingPlace")
		oRS.MoveNext
		If GetRSState(oRS) = True Then sPlace = sPlace & "<br>"
	Loop
	Call RSClose(oRS)

	'希望駅
	sSQL = "sp_GetDataHopeCommuting '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		If sHopeCommuteTime <> "" Then sHopeCommuteTime = sHopeCommuteTime & "<br>"
		If ChkStr(oRS.Collect("HopeCommuteTime")) <> "" Then
			flgCommuteTime = True
			sHopeCommuteTime = oRS.Collect("HopeCommuteTime") & "分"
		End If

		If ChkStr(oRS.Collect("StationName")) <> "" Then
			flgStation = True
			If sStation <> "" Then sStation = sStation & "<br>"
			sStation = sStation & oRS.Collect("StationName") & "駅"

			If oRS.Collect("MinuteToStation") <> "" Then
				sStation = sStation & "(駅から" & oRS.Collect("MinuteToStation") & "分以内)"
			End If
		End If
		oRS.MoveNext
	Loop
	Call RSClose(oRS)

	If flgCommuteTime = True Then iRow = iRow + 1
	If flgStation = True Then iRow = iRow + 1

	'希望沿線
'	sSQL = "sp_GetDataHopeRailwayLine '" & sStaffCode & "'"
'	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
'	If GetRSState(oRS) = True Then iRow = iRow + 1
'	Do While GetRSState(oRS) = True
'		sRailwayLine = sRailwayLine & _
'			oRS.Collect("RailwayCompanyName") & _
'			"　" & oRS.Collect("RailwayLineName") & "<br>"
'		oRS.MoveNext
'		If GetRSState(oRS) = True Then sRailwayLine = sRailwayLine & "<br>"
'	Loop
'	Call RSClose(oRS)

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"" class=""itNaiyo"">勤務地</th>"
	response.write "</tr><tr>"
	Response.Write "<th>希望国</th>"
	Response.Write "<td>" & sArea & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>希望勤務地</th>"
	Response.Write "<td>" & sPlace & "</td>"
	Response.Write "</tr>"

	If sHopeCommuteTime <> "" Then
		Response.Write "<tr>"
		Response.Write "<th>希望通勤時間</th>"
		Response.Write "<td>" & sHopeCommuteTime & "</td>"
		Response.Write "</tr>"
	End If

	If sStation <> "" Then
		Response.Write "<tr>"
		Response.Write "<th>希望駅</th>"
		Response.Write "<td>" & sStation & "</td>"
		Response.Write "</tr>"
	End If
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件給与部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeSalarySmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sStaffCode
	Dim YMin
	Dim YMax
	Dim MMin
	Dim MMax
	Dim DMin
	Dim DMax
	Dim HMin
	Dim HMax
	Dim PercentagePay
	Dim Remark
	Dim iRow

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	If ChkStr(rRS.Collect("YearlyIncomeMin")) <> "" Then YMin = GetJapaneseYen(rRS.Collect("YearlyIncomeMin"))
	If ChkStr(rRS.Collect("YearlyIncomeMax")) <> "" Then YMax = GetJapaneseYen(rRS.Collect("YearlyIncomeMax"))
	If ChkStr(rRS.Collect("MonthlyIncomeMin")) <> "" Then MMin = GetJapaneseYen(rRS.Collect("MonthlyIncomeMin"))
	If ChkStr(rRS.Collect("MonthlyIncomeMax")) <> "" Then MMax = GetJapaneseYen(rRS.Collect("MonthlyIncomeMax"))
	If ChkStr(rRS.Collect("DailyIncomeMin")) <> "" Then DMin = GetJapaneseYen(rRS.Collect("DailyIncomeMin"))
	If ChkStr(rRS.Collect("DailyIncomeMax")) <> "" Then DMax = GetJapaneseYen(rRS.Collect("DailyIncomeMax"))
	If ChkStr(rRS.Collect("HourlyIncomeMin")) <> "" Then HMin = GetJapaneseYen(rRS.Collect("HourlyIncomeMin"))
	If ChkStr(rRS.Collect("HourlyIncomeMax")) <> "" Then HMax = GetJapaneseYen(rRS.Collect("HourlyIncomeMax"))
	PercentagePay = ChkStr(rRS.Collect("PercentagePayFlag"))
	Remark = ChkStr(rRS.Collect("IncomeRemark"))

	iRow = 0
	If YMin & YMax <> "" Then iRow = iRow + 1
	If MMin & MMax <> "" Then iRow = iRow + 1
	If DMin & DMax <> "" Then iRow = iRow + 1
	If HMin & HMax <> "" Then iRow = iRow + 1
	If PercentagePay <> "" Then iRow = iRow + 1
	If Remark <> "" Then iRow = iRow + 1

	If iRow > 0 Then
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"" class=""itNaiyo"">希望給与</th>"
		response.write "</tr><tr>"
		If YMin & YMax <> "" Then
			Response.Write "<th>年収</th>"
			Response.Write "<td>" & YMin & "～" & YMax &"</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		If MMin & MMax <> "" Then
			Response.Write "<th>月収</th>"
			Response.Write "<td>" & MMin & "～" & MMax & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		If DMin & DMax <> "" Then
			Response.Write "<th>日給</th>"
			Response.Write "<td>" & DMin & "～" & DMax & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		If HMin & HMax <> "" Then
			Response.Write "<th>時給</th>"
			Response.Write "<td>" & HMin & "～" & HMax & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
		End If

		Response.Write "<th>歩合</th>"
		Response.Write "<td>"
		If PercentagePay = "1" Then
			Response.Write "希望する"
		ElseIf PercentagePay = "0" Then
			Response.Write "希望しない"
		Else
			Response.Write "こだわらない"
		End If
		Response.Write "</td>"
		Response.Write "</tr>"

		If Remark <> "" Then
			Response.Write "<tr>"
			Response.Write "<th>備考</th>"
			Response.Write "<td>" & Remark & "</td>"
			Response.Write "</tr>"
		End If
	Else
		'希望給与条件が無い場合
		Response.Write "<tr>"
		Response.Write "<th colspan=""2"">希望給与</th>"
		response.write "</tr><tr>"
		Response.Write "<td colspan=""2""></td>"
		Response.Write "</tr>"
	End If
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件期間・時間部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeSpanSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sStaffCode
	Dim sTransfer
	Dim sWorkPeriod
	Dim sWorkPeriodFlag
	Dim sWorkMonthPeriod
	Dim sWorkTime
	Dim sWSTime
	Dim sWETime
	Dim sOverWork
	Dim sOverWorkTimeMax
	Dim sWorkShift
	Dim sHoliday

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sTransfer = ChkStr(rRS.Collect("TransferFlag"))
	Select Case sTransfer
		Case "1":	sTransfer = "可"
		Case "0":	sTransfer = "不可"
		Case Else:	sTransfer = "こだわらない"
	End Select

	sWorkPeriod = ChkStr(oRS.Collect("WorkPeriod"))
	sWorkPeriodFlag = ChkStr(oRS.Collect("WorkPeriodFlag"))
	sWorkMonthPeriod = ChkStr(oRS.Collect("WorkMonthPeriod"))
	sWSTime = ChkStr(oRS.Collect("WorkStartTime"))
	sWETime = ChkStr(oRS.Collect("WorkEndTime"))
	If sWSTime & sWETime <> "" Then
		If sWSTime <> "" Then sWorkTime = sWorkTime & Left(sWSTime, 2) & ":" & Right(sWSTime, 2)
		sWorkTime = sWorkTime & "～"
		If sWETime <> "" Then sWorkTime = sWorkTime & Left(sWETime, 2) & ":" & Right(sWETime, 2)
	End If

	sOverWork = ChkStr(rRS.Collect("OverWorkFlag"))
	Select Case sOverWork
		Case "1":	sOverWork = "可"
		Case "0":	sOverWork = "不可"
		Case Else:	sOverWork = "こだわらない"
	End Select
	sOverWorkTimeMax = ChkStr(rRS.Collect("OverWorkTimeMax"))
	If sOverWorkTimeMax <> "" Then sOverWorkTimeMax = Left(sOverWorkTimeMax, 2) & ":" & Right(sOverWorkTimeMax, 2)

	sWorkShift = ChkStr(rRS.Collect("WorkShiftFlag"))
	Select Case sWorkShift
		Case "1":	sWorkShift = "可"
		Case "0":	sWorkShift = "不可"
		Case Else:	sWorkShift = "こだわらない"
	End Select

	sHoliday = ""
	If oRS.Collect("MonHolidayFlag") = "1" Then sHoliday = sHoliday & "月"
	If oRS.Collect("TueHolidayFlag") = "1" Then sHoliday = sHoliday & "火"
	If oRS.Collect("WedHolidayFlag") = "1" Then sHoliday = sHoliday & "水"
	If oRS.Collect("ThuHolidayFlag") = "1" Then sHoliday = sHoliday & "木"
	If oRS.Collect("FriHolidayFlag") = "1" Then sHoliday = sHoliday & "金"
	If oRS.Collect("SatHolidayFlag") = "1" Then sHoliday = sHoliday & "土"
	If oRS.Collect("SunHolidayFlag") = "1" Then sHoliday = sHoliday & "日"
	If oRS.Collect("PublicHolidayFlag") = "1" Then sHoliday = sHoliday & "祝"
	If ChkStr(oRS.Collect("WeeklyHolidayType")) <> "" Then
		If sHoliday <> "" Then sHoliday = sHoliday & "<br>"
		sHoliday = sHoliday & oRS.Collect("WeeklyHolidayType")
	End If
	If ChkStr(oRS.Collect("HolidayRemark")) <> "" Then
		If sHoliday <> "" Then sHoliday = sHoliday & "<br>"
		sHoliday = sHoliday & oRS.Collect("HolidayRemark")
	End If

	Response.Write "<tr>"
	Response.Write "<th>転勤可／不可</th>"
	Response.Write "<td>" & sTransfer & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>勤務期間</th>"
	Response.Write "<td>"
	Response.Write sWorkPeriod

	If sWorkPeriodFlag = "3" And sWorkMonthPeriod <> "" Then
		'短期
		Response.Write " ( " & sWorkMonthPeriod & "ヶ月 )"
	End If

	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>就業時間</th>"
	Response.Write "<td>" & sWorkTime & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>残業</th>"
	Response.Write "<td>"
	Response.Write sOverWork

	If sOverWorkTimeMax <> "" Then
		Response.Write "(" & sOverWorkTimeMax & "まで)"
	End If

	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>シフト(交代)稼働</th>"
	Response.Write "<td>" & sWorkShift & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>休日</th>"
	Response.Write "<td>" & sHoliday & "</td>"
	Response.Write "</tr>"
End Function

'******************************************************************************
'概　要：プロフィールページの希望条件福利厚生部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileHopeWelfareSmart(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sStaffCode
	Dim sWelfareProgram

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sWelfareProgram = ""
	If rRS.Collect("TrafficFeeFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　交通費支給"
	If rRS.Collect("SocietyInsuranceFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　社会保険完備"
	If rRS.Collect("SanatoriumFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　保養所"
	If rRS.Collect("EnterprisePensionFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　企業年金"
	If rRS.Collect("WealthShapeFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　財形貯蓄"
	If rRS.Collect("StockOptionFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　持株制度(ストックオプション)"
	If rRS.Collect("RetirementPayFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　退職金制度"
	If rRS.Collect("ResidencePayFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　住宅手当"
	If rRS.Collect("FamilyPayFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　家族手当"
	If rRS.Collect("EmployeeDormitoryFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　社員寮"
	If rRS.Collect("CompanyHouseFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　社宅"
	If rRS.Collect("NewEmployeeTrainingFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　新入社員研修"
	If rRS.Collect("OverseasTrainingFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　海外研修"
	If rRS.Collect("OtherTrainingFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　各種研修"
	If rRS.Collect("FlexTimeFlag") = "1" Then sWelfareProgram = sWelfareProgram & "　フレックスタイム"

	If sWelfareProgram <> "" Then
		sWelfareProgram = Mid(sWelfareProgram, 2)

		Response.Write "<tr>"
		Response.Write "<th>福利厚生</th>"
		Response.Write "<td>" & sWelfareProgram & "</td>"
		Response.Write "</tr>"
	End If
End Function




'******************************************************************************
'概　要：プロフィールページの転職の将来像（キャリア棚卸しアナライザーより）を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2009/09/24 LIS T.Ezaki 作成
'******************************************************************************
Function DspCareerAnalyzer(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode
	Dim sPattern

	If vUserType = "staff" Then
		sPattern = "pattern9"
	ElseIf vUserType = "company" Then
		sPattern = "pattern8"
	End If
	
	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")
	
	sSQL = "select "
	sSQL = sSQL & " IdealIndustryText"
	sSQL = sSQL & ",IdealIndustryPriority"
	sSQL = sSQL & ",IdealIndustryDistance"
	sSQL = sSQL & ",IdealPositionText"
	sSQL = sSQL & ",IdealPositionPriority"
	sSQL = sSQL & ",IdealPositionDistance"
	sSQL = sSQL & ",IdealJobText"
	sSQL = sSQL & ",IdealJobPriority"
	sSQL = sSQL & ",IdealJobDistance"
	sSQL = sSQL & ",IdealCustomText"
	sSQL = sSQL & ",IdealCustomPriority"
	sSQL = sSQL & ",IdealCustomDistance"
	sSQL = sSQL & ",IdealServiceText"
	sSQL = sSQL & ",IdealServicePriority"
	sSQL = sSQL & ",IdealServiceDistance"
	sSQL = sSQL & ",IdealRelationsText"
	sSQL = sSQL & ",IdealRelationsPriority"
	sSQL = sSQL & ",IdealRelationsDistance"
	sSQL = sSQL & ",IdealFutureText"
	sSQL = sSQL & ",IdealFuturePriority"
	sSQL = sSQL & ",IdealFutureDistance"
	sSQL = sSQL & ",IdealWorkareaText"
	sSQL = sSQL & ",IdealWorkareaPriority"
	sSQL = sSQL & ",IdealWorkareaDistance"
	sSQL = sSQL & ",IdealTrainingText"
	sSQL = sSQL & ",IdealTrainingPriority"
	sSQL = sSQL & ",IdealTrainingDistance"
	sSQL = sSQL & ",IsNull(CareerGoolType,'') as CareerGoolType"
	sSQL = sSQL & ",IsNull(CareerGoolEtc,'') as CareerGoolEtc"
	sSQL = sSQL & ",IsNull(CareerGoolDetail,'') as CareerGoolDetail"
	sSQL = sSQL & ",Publicflag"
	sSQL = sSQL & " from P_CareerAnalyzer"
	sSQL = sSQL & " Where StaffCode = '" & sStaffCode & "' and Publicflag = 1"
	
	flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
	If GetRSState(oRS) = true then
		response.write "<table class=""" & sPattern & " cw"" style=""margin-bottom:15px;"">"
		response.write "<thead><tr><th colspan=""4"" style=""width:588px;"">将来な理想像</th></tr></thead>"
		response.write "<tbody>"
		response.write "<tr>"
		response.write "<th style=""width:188px;"">種類</th>"
		response.write "<th style=""text-align:left"">理想像</th>"
		response.write "<th style=""text-align:center"">優先度</th>"
		response.write "<th style=""text-align:center"">距離感</th>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>業界・業種の理想像</th>"
		response.write "<td>" & oRS.Collect("IdealIndustryText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealIndustryPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealIndustryDistance")
			case "1"	:	response.write "遠い"
			case "2"	:	response.write "やや遠い"
			case "3"	:	response.write "ふつう"
			case "4"	:	response.write "やや近い"
			case "5"	:	response.write "近い" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>ポジションの理想像</th>"
		response.write "<td>" & oRS.Collect("IdealPositionText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealPositionPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealPositionDistance")
			case "1"	:	response.write "遠い"
			case "2"	:	response.write "やや遠い"
			case "3"	:	response.write "ふつう"
			case "4"	:	response.write "やや近い"
			case "5"	:	response.write "近い" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>職種・仕事内容の理想像</th>"
		response.write "<td>" & oRS.Collect("IdealJobText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealJobPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealJobDistance")
			case "1"	:	response.write "遠い"
			case "2"	:	response.write "やや遠い"
			case "3"	:	response.write "ふつう"
			case "4"	:	response.write "やや近い"
			case "5"	:	response.write "近い" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>社風の理想像</th>"
		response.write "<td>" & oRS.Collect("IdealCustomText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealCustomPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealCustomDistance")
			case "1"	:	response.write "遠い"
			case "2"	:	response.write "やや遠い"
			case "3"	:	response.write "ふつう"
			case "4"	:	response.write "やや近い"
			case "5"	:	response.write "近い" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>給与・待遇の理想像</th>"
		response.write "<td>" & oRS.Collect("IdealServiceText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealServicePriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealServiceDistance")
			case "1"	:	response.write "遠い"
			case "2"	:	response.write "やや遠い"
			case "3"	:	response.write "ふつう"
			case "4"	:	response.write "やや近い"
			case "5"	:	response.write "近い" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>人間関係の理想像</th>"
		response.write "<td>" & oRS.Collect("IdealRelationsText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealRelationsPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealRelationsDistance")
			case "1"	:	response.write "遠い"
			case "2"	:	response.write "やや遠い"
			case "3"	:	response.write "ふつう"
			case "4"	:	response.write "やや近い"
			case "5"	:	response.write "近い" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>会社将来性の理想像</th>"
		response.write "<td>" & oRS.Collect("IdealFutureText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealFuturePriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealFutureDistance")
			case "1"	:	response.write "遠い"
			case "2"	:	response.write "やや遠い"
			case "3"	:	response.write "ふつう"
			case "4"	:	response.write "やや近い"
			case "5"	:	response.write "近い" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>勤務地の理想像</th>"
		response.write "<td>" & oRS.Collect("IdealWorkareaText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealWorkareaPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealWorkareaDistance")
			case "1"	:	response.write "遠い"
			case "2"	:	response.write "やや遠い"
			case "3"	:	response.write "ふつう"
			case "4"	:	response.write "やや近い"
			case "5"	:	response.write "近い" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>教育研修の理想像</th>"
		response.write "<td>" & oRS.Collect("IdealTrainingText") & "</td>"
		response.write "<td style=""text-align:center"">" & oRS.Collect("IdealTrainingPriority") & "</td>"
		response.write "<td style=""text-align:center"">"
		select case oRS.Collect("IdealTrainingDistance")
			case "1"	:	response.write "遠い"
			case "2"	:	response.write "やや遠い"
			case "3"	:	response.write "ふつう"
			case "4"	:	response.write "やや近い"
			case "5"	:	response.write "近い" 
		end select
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr>"
		response.write "<th>キャリアのゴールイメージ"
		response.write "</th>"
		response.write "<td colspan=""3"">"
		response.write "<strong style=""font-size:16px;"">"
		Select Case oRS.Collect("CareerGoolType")
			Case "1"	:	Response.Write "スキル探求志向"
			Case "2"	:	Response.Write "マネージメント志向"
			Case "3"	:	Response.Write "独立志向"
			Case "4"	:	Response.Write oRS.Collect("CareerGoolEtc")
		End Select
		response.write "</strong><br>"
		Response.Write Replace(oRS.Collect("CareerGoolDetail"),vbCrLf,"<br>")
		response.write "</td>"
		response.write "</tr>"
		response.write "</tbody>"
		response.write "</table>"
		Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">▲ページTOPへ</a></p>" & vbCrLf
	end if
	Call RSClose(oRS)
End Function

'******************************************************************************
'概　要：プロフィールページの最新送信メール状況部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileMail(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	'DB
	Dim dbStaffCode
	Dim dbSendDay1
	Dim dbSubject1
	Dim dbBody1
	Dim dbSendDay2
	Dim dbSubject2
	Dim dbBody2

	If GetRSState(rRS) = False Then Exit Function

	dbStaffCode = rRS.Collect("StaffCode")

	If vUserType = "company" Or vUserType = "dispatch" Then
		sSQL = "EXEC up_DtlMailHistory_Staff '" & vUserID & "', '" & dbStaffCode & "', '" & vOrderCode & "';"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			dbSendDay1 = oRS.Collect("SendDay1")
			dbSubject1 = oRS.Collect("Subject1")
			dbBody1 = ChkStr(oRS.Collect("Body1"))
			dbSendDay2 = oRS.Collect("SendDay2")
			dbSubject2 = oRS.Collect("Subject2")
			dbBody2 = ChkStr(oRS.Collect("Body2"))

			Response.Write "<table class=""pattern2 cw"" border=""0"" style=""margin-bottom:15px;"">"
			Response.Write "<colgroup>"
			Response.Write "<col style=""width:300px;"">"
			Response.Write "<col style=""width:300px;"">"
			Response.Write "<colgroup>"
			Response.Write "<thead>"
			Response.Write "<tr>"
			Response.Write "<th colspan=""2"" style=""text-align:center;"">メール最新状況</th>"
			Response.Write "</tr>"
			Response.Write "</thead>"
			Response.Write "<tbody>"
			Response.Write "<tr>"
			Response.Write "<th style=""text-align:center;"">貴社が送信したもの</th>"
			Response.Write "<th style=""text-align:center;"">求職者が送信したもの</th>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td style=""vertical-align:top;"">"
			If ChkStr(dbSendDay1) <> "" Then
				Response.Write "【送信日時】<br>" & dbSendDay1
				Response.Write "<div class=""line1""></div>"
				Response.Write "【件名】<br>" & dbSubject1
				Response.Write "<div class=""line1""></div>"
				Response.Write "【本文】<br>" & Replace(dbBody1,vbCrLf,"<br>")
			End If
			Response.Write "</td>"
			Response.Write "<td style=""vertical-align:top;"">"
			If ChkStr(dbSendDay2) <> "" Then
				Response.Write "【送信日時】<br>" & dbSendDay2
				Response.Write "<div class=""line1""></div>"
				Response.Write "【件名】<br>" & dbSubject2
				Response.Write "<div class=""line1""></div>"
				Response.Write "【本文】<br>" & Replace(dbBody2,vbCrLf,"<br>")
			End If
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "</tbody>"
			Response.Write "</table>"
			Response.Write "<p style=""text-align:right;"" class=""m0""><a class=""stext"" href=""#pagetop"">▲ページTOPへ</a></p>" & vbCrLf
		End If
		Call RSClose(oRS)
	End If
End Function

'******************************************************************************
'概　要：プロフィールページの求職者コードを出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/12 LIS K.Kokubo 作成
'　　　：2008/09/09 LIS K.Kokubo 求職者名入力対応
'******************************************************************************
Function DspProfileStaffCode(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim oRS
	Dim sSQL
	Dim flgQE
	Dim sError

	Dim dbStaffName
	Dim sStaffCode
	Dim sTableClass

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sSQL = "EXEC up_DtlCMPStaffName '" & vUserID & "', '" & sStaffCode & "';"
	flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		dbStaffName = ChkStr(oRS.Collect("StaffName"))
	End If
	Call RSClose(oRS)

	sTableClass = "pattern9"
	If vUserType = "company" Or vUserType = "dispatch" Then sTableClass = "pattern8"

	Response.Write "<div class=""center"">"
	If dbStaffName <> "" Then
		Response.Write dbStaffName & "&nbsp;(求職者コード：" & sStaffCode & ")"
	Else
		Response.Write sStaffCode & "(会員ID)"
	End If

	Response.Write "</div>" & vbCrLf
End Function

'******************************************************************************
'概　要：プロフィールページの各編集ボタンを出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/16 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileEditButton(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Response.Write "<form id=""frmEdit"" action="""" method=""post"">"
	Response.Write "<input type=""hidden"" name=""CONF_StaffCode"" value=""" & vUserID & """>"
	Response.Write "<h2>各種編集ボタン</h2>"
	Response.Write "※編集する場合は以下のボタンを押して下さい。<br>"
	Response.Write "<table class=""cw profileEdit"" border=""1"" bordercolor=""#ff9999"" style=""font-size:10px; width:745px;margin-left:-5px;"">"
	Response.Write "<tr>"
	Response.Write "<td align=""center"" valign=""top"">"
	Response.Write "履歴書・職務経歴書(一般）・スカウト情報<br>"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit1.jpg"" id=""button1"" value=""個人データ"" onclick=""document.forms.frmEdit.action = './person_edit1.asp'; document.forms.frmEdit.submit();"">"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit2.jpg"" id=""button2"" value=""学歴職歴"" onclick=""document.forms.frmEdit.action = './person_edit2.asp'; document.forms.frmEdit.submit();"" style=""padding-left:3px;"">"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit3.jpg"" id=""button3"" value=""資格・スキル"" onclick=""document.forms.frmEdit.action = './person_edit3.asp'; document.forms.frmEdit.submit();"" style=""padding-left:3px;"">"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit4.jpg"" id=""button4"" value=""IT系&#10;&#13;開発スキル"" onclick=""document.forms.frmEdit.action = './person_edit4.asp'; document.forms.frmEdit.submit();"" style=""padding-left:3px;"">"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit9.jpg"" id=""button9"" value=""語学スキル"" onclick=""document.forms.frmEdit.action = './person_edit9.asp'; document.forms.frmEdit.submit();"" style=""padding-left:3px;"">"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit7.jpg"" id=""button7"" value=""自己ＰＲ&#10;&#13;志望動機等"" onclick=""document.forms.frmEdit.action = './person_edit7.asp'; document.forms.frmEdit.submit();"" style=""padding-left:3px;"">"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit8.jpg"" id=""button8"" value=""得意分野等"" onclick=""document.forms.frmEdit.action = './person_edit8.asp'; document.forms.frmEdit.submit();"" style=""padding-left:3px;"">"
	Response.Write "</td>"
	Response.Write "<td align=""center"" valign=""top"">"
	Response.Write "職務経歴書(IT)<br>"
	Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit5.jpg"" id=""button5"" value=""IT系&#10;&#13;職務詳細"" onclick=""document.forms.frmEdit.action = './person_edit5.asp'; document.forms.frmEdit.submit();"">"
	Response.Write "</td>"
	Response.Write "<td align=""center"" valign=""top"">"
	Response.Write "スカウト情報<br>"
    '2015/09/01 リンク変更
	'Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit6.jpg"" id=""button6"" value=""本人希望"" onclick=""document.forms.frmEdit.action = './person_edit6.asp'; document.forms.frmEdit.submit();"">"
    Response.Write "<input type=""image"" src=""/img/staff/person_detail_edit6.jpg"" id=""button6"" value=""本人希望"" onclick=""document.forms.frmEdit.action = './step2a.asp'; document.forms.frmEdit.submit();"">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</form>" & vbCrLf

	Response.Write "<p>"
	Response.Write "履歴書情報を多く入力して他人と差をつけましょう！<br>"
	Response.Write "<span style=""color:red;font-weight:bold;font-size:14px;line-height:font-weight:14px;"">★下記内容は企業側がスカウトする際に表示される内容と同一です。</span><br>"
	Response.Write "<span style=""color:#ff0000;"">※魅力ある内容を入力すればより採用されやすくなります。なお個人を特定できる情報が出ていないかご確認ください。</span>"
	Response.Write "</p>" & vbCrLf
End Function

'******************************************************************************
'概　要：プロフィールページの最終更新日を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/16 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileUpdateDay(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	If GetRSState(rRS) = False Then Exit Function

	Response.Write "<table class=""cw"" border=""0"">"
	Response.Write "<tr>"
	Response.Write "<td style=""text-align:right;"">最終更新日：" & GetDateStr(rRS.Collect("UpdateDay"), "/") & "</td>"
	Response.Write "</tr>"
	Response.Write "</table>" & vbCrLf
End Function

'******************************************************************************
'概　要：プロフィールページの職歴入力無しに対する文言を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/16 LIS K.Kokubo 作成
'******************************************************************************
Function DspProfileAttentionCareerHistory(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	If GetRSState(rRS) = False Then Exit Function
%>
<div style="margin-top:0px;">
	<table class="cw" border="0" style="width:270px; float:left; margin-bottom:10px;">
		<thead>
		<tr>
			<th style="padding:5px;border:1px solid #ff0000; border-width:1px 1px 0px 1px;background-color:#ffdddd;">職歴を入力しましょう！</th>
		</tr>
		</thead>
		<tbody>
		<tr>
			<td style="padding:5px;border:1px solid #ff0000;background-color:#ffeeee;">
				<p>
					職歴は求人企業がとても注目している項目です。<br>
					職歴の入力が無い方は、<br>
					<b>(1)スカウトされにくい</b><br>
					<b>(2)応募しても書類選考で落ちてしまう</b><br>
					等の可能性が高くなります。<br>
				</p>
			</td>
		</tr>
		</tbody>
	</table>

	<table class="cw" border="0" style="width:320px; float:right;">
		<thead>
		<tr>
			<th style="padding:5px;border:1px solid #ff0000; border-width:1px 1px 0px 1px;background-color:#ffdddd;">登録完了により、以下機能がご利用可能です</th>
		</tr>
		</thead>
		<tbody>
		<tr>
			<td style="padding:5px;border:1px solid #ff0000;background-color:#ffeeee;">
				<p>
					◆<a href="/order/order_search_detail.asp" title="お仕事詳細検索">好きな仕事に応募</a>できる。<br>
					◆<a href="/staff/s_resume.asp" title="履歴書">履歴書</a>・<a href="/staff/s_careersheet.asp" title="職務経歴書">職務経歴書</a>の印刷ができる。<br>
					◆<a href="/s_contents/motive_index.asp" title="志望動機">志望動機</a>・<a href="/s_contents/s_jikopr.asp" title="自己PR">自己PR</a>の作成支援ツールが使用できる。<br>
					◆心理学に基づいた<a href="/s_contents/s_mynavi.asp" title="適職診断">適職診断</a>が受けられる。<br>
					◆<a href="/s_contents/enquete.asp" title="おしごとアンケート">おしごとアンケート</a>に参加できる。<br>
				</p>
			</td>
		</tr>
		</tbody>
	</table>
    <p style="text-align:right;" class="m0"><a class="stext" href="#pagetop">▲ページTOPへ</a></p>
    <br clear="all">
</div>
<%
End Function

'******************************************************************************
'概　要：メールマガジンからのプロフィールページへのアクセスをログに記録
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2007/04/16 LIS K.Kokubo 作成
'******************************************************************************
Function RegMailMagazineAccess(ByRef rDB, ByVal vUserID, ByVal vMailMagazineID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sSuspensionFlag

	RegMailMagazineAccess = False
	sSuspensionFlag = "0"
	If vMailMagazineID <> "" Then
		RegMailMagazineAccess = True
		sSQL = "up_MailMagazineAccess '" & vMailMagazineID & "', '" & vUserID & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then
			If oRS.Collect("SuspensionFlag") = "1" Then RegMailMagazineAccess = False
		Else
			RegMailMagazineAccess = False
		End If
		Call RSClose(oRS)
	End If
End Function

'******************************************************************************
'概　要：情報更新の通知をリスの社員にメールする
'引　数：rDB			：接続中のDBConnection
'　　　：vUserID		：ログイン中ユーザＩＤ
'　　　：vMailServer	：メールサーバー
'　　　：vFrom			：送信元
'　　　：vSubject		：件名
'　　　：vBody			：内容
'使用元：
'備　考：
'履　歴：2007/04/16 LIS K.Kokubo 作成
'　　　：2010/10/22 LIS K.Kokubo ＬＩＳ現存社員にのみメールを送るように修正
'******************************************************************************
Function SendMailStaffEdit(ByRef rDB, ByVal vUserID, ByVal vMailServer, ByVal vFrom, ByVal vSubject, ByVal vBody)
	Dim sSQL,oRS,flgQE,sSQLErr
	Dim sRet,sBody,sLisEmployeeMail

	sLisEmployeeMail = ""
	'<対象求職者をウォッチリストに登録しているリス現存社員のメールアドレス一覧を取得>
	sSQL = "EXEC up_LstLISWatchList '" & G_USERID & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sSQLErr)
	Do While GetRSState(oRS) = True
		If sLisEmployeeMail <> "" Then sLisEmployeeMail = sLisEmployeeMail & vbTab
		sLisEmployeeMail = sLisEmployeeMail & ChkStr(oRS.Collect("MailAddress"))
		oRS.MoveNext
	Loop
	'</対象求職者をウォッチリストに登録しているリス現存社員のメールアドレス一覧を取得>

	sBody = ""
	sBody = sBody & "■求職者コード　[" & G_USERID & "]" & vbCrLf
	sBody = sBody & HTTP_BI_CURRENTURL & "staff/Staff_detail.asp?staffcode=" & G_USERID & vbCrLf & vbCrLf 
	sBody = sBody & vBody

	sRet = SndMail(vMailServer, sLisEmployeeMail, vFrom, vSubject, sBody, "")
End Function

'******************************************************************************
'概　要：企業が求職者詳細を閲覧したらログに書き込む
'引　数：rDB			：接続中のDBConnection
'　　　：vCompanyCode	：ログイン中企業コード
'　　　：vOrderCode		：ログイン中企業が求職者を検索しているお仕事の情報コード
'　　　：vStaffCode		：求職者コード
'使用元：
'備　考：
'更　新：2007/10/22 LIS K.Kokubo 作成
'******************************************************************************
Function RegAccessHistoryStaff(ByRef rDB, ByVal vCompanyCode, ByVal vOrderCode, ByVal vStaffCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	RegAccessHistoryStaff = False

	If vCompanyCode <> "" And vOrderCode <> "" And vStaffCode <> "" And vOrderCode <> vStaffCode Then
		sSQL = "up_RegLOG_AccessHistoryStaff '" & vOrderCode & "', '" & vStaffCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)

		RegAccessHistoryStaff = flgQE
	End If
End Function

'******************************************************************************
'概　要：企業が求職者詳細を閲覧したらログに書き込む
'引　数：rDB				：接続中のDBConnection
'　　　：vCompanyCode		：ログイン中企業コード
'　　　：vOrderCode			：ログイン中企業が求職者を検索しているお仕事の情報コード
'　　　：vStaffCodeArray	：求職者コードの配列
'　　　：vPage				：表示中の一覧のページ数
'使用元：
'備　考：
'更　新：2007/11/12 LIS K.Kokubo 作成
'******************************************************************************
Function RegAccessHistoryStaffList(ByRef rDB, ByVal vCompanyCode, ByVal vOrderCode, ByVal vStaffCodeArray)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sWhere

	Dim flgResult
	Dim idx
	Dim sDeclare
	Dim sParams

	flgResult = False

	'対象求職者の有無をチェックして, INSERT文, UPDATE文 の作成
	If vCompanyCode <> "" And vOrderCode <> "" And UBound(vStaffCodeArray) >= 0 Then
		'**************************************************************************
		'①UPDATE文作成 start
		'--------------------------------------------------------------------------
		sWhere = ""
		sDeclare = "@vOrderCode CHAR(8),@vCompanyCode CHAR(8)"
		sParams = ",@vOrderCode = N'" & vOrderCode & "',@vCompanyCode = N'" & vCompanyCode & "'"

		For idx = LBound(vStaffCodeArray) To UBound(vStaffCodeArray)
			sDeclare = sDeclare & ",@vStaffCode" & idx & " CHAR(8)"
			sParams = sParams & ",@vStaffCode" & idx & " = N'" & vStaffCodeArray(idx) & "'"

			If sWhere <> "" Then sWhere = sWhere & ","
			sWhere = sWhere & "@vStaffCode" & idx
		Next

		sSQL = "" & _
			"UPDATE LOG_AccessHistoryStaffList " & _
			"SET YYYYMM = CONVERT(VARCHAR(4), YEAR(GETDATE())) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(GETDATE())), 2) " & _
				",UpdateDay = GETDATE() " & _
			"WHERE YYYYMM = CONVERT(VARCHAR(4), YEAR(GETDATE())) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(GETDATE())), 2) " & _
				"AND OrderCode = @vOrderCode " & _
				"AND StaffCode IN (" & sWhere & ") "

		sSQL = Replace(sSQL, "'", "''")
		sSQL = "sp_executesql N'" & sSQL & "', N'" & sDeclare & "'" & sParams
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		flgResult = flgQE
		'--------------------------------------------------------------------------
		'①UPDATE文作成 end
		'**************************************************************************

		'**************************************************************************
		'②INSERT文作成 start
		'--------------------------------------------------------------------------
		sSQL = ""
		sWhere = ""
		sDeclare = "@vOrderCode CHAR(8),@vCompanyCode CHAR(8)"
		sParams = ",@vOrderCode = N'" & vOrderCode & "',@vCompanyCode = N'" & vCompanyCode & "'"

		For idx = LBound(vStaffCodeArray) To UBound(vStaffCodeArray)
			sDeclare = sDeclare & ",@vStaffCode" & idx & " CHAR(8)"
			sParams = sParams & ",@vStaffCode" & idx & " = N'" & vStaffCodeArray(idx) & "'"

			If sSQL <> "" Then sSQL = sSQL & "UNION "
			sSQL = sSQL & _
				"SELECT CONVERT(VARCHAR(4), YEAR(GETDATE())) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(GETDATE())), 2) AS YYYYMM " & _
					",@vStaffCode" & idx & " AS StaffCode " & _
					",@vOrderCode AS OrderCode " & _
					",@vCompanyCode AS CompanyCode " & _
					",GETDATE() AS RegistDay "

			If sWhere <> "" Then sWhere = sWhere & ","
			sWhere = sWhere & "@vStaffCode" & idx
		Next

		sSQL = "INSERT INTO LOG_AccessHistoryStaffList " & _
			"SELECT INS.YYYYMM, INS.StaffCode, INS.OrderCode, INS.CompanyCode, INS.RegistDay, INS.RegistDay " & _
			"FROM (" & sSQL & ") AS INS " & _
			"WHERE NOT EXISTS( " & _
					"SELECT * " & _
					"FROM LOG_AccessHistoryStaffList AS NEX " & _
					"WHERE INS.YYYYMM = NEX.YYYYMM " & _
						"AND INS.OrderCode = NEX.OrderCode " & _
						"AND NEX.StaffCode IN (" & sWhere & ") " & _
				")"

		sSQL = Replace(sSQL, "'", "''")
		sSQL = "sp_executesql N'" & sSQL & "', N'" & sDeclare & "'" & sParams
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		flgResult = flgQE
		'--------------------------------------------------------------------------
		'②INSERT文作成 end
		'**************************************************************************
	End If

	RegAccessHistoryStaffList = flgResult
End Function

'******************************************************************************
'概　要：プロフィールページの希望連絡時間帯部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2015/07/17 LIS K.Kimura 作成
'******************************************************************************
Function DspProfileHopeContactableTime(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim sStaffCode
	Dim sContactableTime	: sContactableTime = ""

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	sSQL = "sp_GetDataContactableTime '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	Do While GetRSState(oRS) = True
		sContactableTime = sContactableTime & oRS.Collect("ContactableTimeName")
		oRS.MoveNext
		If GetRSState(oRS) = True Then sContactableTime = sContactableTime & "　"
	Loop
	Call RSClose(oRS)

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">希望連絡時間帯</th>"
	Response.Write "<td>"
	Response.Write sContactableTime
	Response.Write "</td>"
	Response.Write "</tr>"

End Function

'******************************************************************************
'概　要：プロフィールページの希望条件業種部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：sp_GetDetailStaff で生成されたレコードセットオブジェクト
'　　　：vUserType		：ログイン中ユーザ種類
'　　　：vUserID		：ログイン中ユーザＩＤ
'使用元：
'備　考：
'更　新：2018/07/10 LIS takaya 作成
'******************************************************************************
Function DspProfileSideJob(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError
	Dim sStaffCode
    Dim sMsg

	If GetRSState(rRS) = False Then Exit Function

	sStaffCode = rRS.Collect("StaffCode")

	Response.Write "<tr>"
	Response.Write "<th colspan=""2"">副業</th>"
	Response.Write "<td>"

	sSQL = "select doit,AtHome,Attendance from p_SideJob where staffcode = '" & sStaffCode & "'"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	if GetRSState(oRS) = True then
        if oRS.Collect("Doit") = "2" then
            sMsg = "希望する"

            if oRS.Collect("AtHome") = "1" or oRS.Collect("Attendance") = "1"  then
                sMsg = sMsg + "　( " 

                if oRS.Collect("AtHome") = "1" then
                    sMsg = sMsg + " 在宅 "
                end if

                if oRS.Collect("Attendance") = "1" then
                    sMsg = sMsg + " 出勤 "
                end if

                sMsg = sMsg + " )" 
            end if

            Response.Write sMsg 
        else
            Response.Write "希望しない"
        end if
    else
        Response.Write "希望しない"
	end if
	Call RSClose(oRS)

	Response.Write "</td>"
	Response.Write "</tr>"
End Function

%>
