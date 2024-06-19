<%
'*******************************************************************************
'概　要：求職者の会員登録画面のアンケートHTMLを取得
'引　数：
'出　力：
'戻り値：String
'備　考：
'履　歴：2011/07/06 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlStaffRegQuestionnaire()
	Dim sHTML

	Dim frmQ1
	Dim aQ1(6),idx,tmpAry

	frmQ1 = GetForm("frmq1",2)

	'<チェックボックス or ラヂオボタンのデフォルト設定>
	tmpAry = Split(Replace(frmQ1," ",""),",")
	For idx = 0 To UBound(tmpAry)
		aQ1(tmpAry(idx)) = " checked"
	Next
	'</チェックボックス or ラヂオボタンのデフォルト設定>


	sHTML = sHTML & "<table class=""pattern1_1"" >"
	sHTML = sHTML & "<tbody>"

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<th id=""thenq"" class=""first_th"">アンケートにご協力ください<br>（ご登録のきっかけは？）</th>"
	sHTML = sHTML & "<td class=""first_td"">"
	sHTML = sHTML & "<ul class=""left"" style=""margin-right:80px;"">"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""0""" & aQ1(0) & " onclick=""blur();if(needcount)needcount();""> 希望する求人への応募</label></li>"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""1""" & aQ1(1) & " onclick=""blur();if(needcount)needcount();""> 転職サポート(人材紹介・派遣)を希望</label></li>"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""2""" & aQ1(2) & " onclick=""blur();if(needcount)needcount();""> 履歴書・職務経歴書作成の利用</label></li>"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""3""" & aQ1(3) & " onclick=""blur();if(needcount)needcount();""> 会員限定のコンテンツ利用</label></li>"
	sHTML = sHTML & "</ul>"
	sHTML = sHTML & "<ul>"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""4""" & aQ1(4) & " onclick=""blur();if(needcount)needcount();""> 携帯電話・スマートフォンで転職活動</label></li>"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""5""" & aQ1(5) & " onclick=""blur();if(needcount)needcount();""> コンビニ印刷機能を希望</label></li>"
	sHTML = sHTML & "<li><label><input name=""frmq1"" type=""checkbox"" value=""6""" & aQ1(6) & " onclick=""blur();if(needcount)needcount();""> その他</label> （ <input id=""txtQ1"" name=""frmq1other"" type=""text"" maxlength=""200"" value="""" style=""width:180px;"" onkeyup=""if(needcount)needcount();""> ）</li>"
	sHTML = sHTML & "</ul>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	sHTML = sHTML & "</tbody>"
	sHTML = sHTML & "</table>"

	htmlStaffRegQuestionnaire = sHTML
End Function
%>
