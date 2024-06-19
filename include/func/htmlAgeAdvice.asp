<%
'*******************************************************************************
'概　要：年齢制限のアドバイスＨＴＭＬを取得
'引　数：vAgeMin	：年齢下限
'　　　：vAgeMax	：年齢上限
'戻り値：String
'備　考：
'履　歴：2010/11/12 LIS K.Kokubo 作成
'*******************************************************************************
Function htmlAgeAdvice(ByVal vAgeMin,ByVal vAgeMax)
	Dim sHTML

	sHTML = sHTML & "<p style=""margin-bottom:5px;"">"
	sHTML = sHTML & "【年齢制限がＯＫな場合のアドバイス】<br>"
	sHTML = sHTML & "・<span style=""color:#ff0000;"">正社員募集</span>かつ、<span style=""color:#ff0000;"">定年までの年齢を募集</span>している。（下限の記載は×）<br>"
	sHTML = sHTML & "　＜理由例＞６０歳未満の方を募集（定年が６０歳）<br>"
	sHTML = sHTML & "・<span style=""color:#ff0000;"">正社員募集</span>かつ、<span style=""color:#ff0000;"">職業経験不問</span>かつ、<span style=""color:#ff0000;"">新規学卒者と同等の処遇</span>であること。（下限の記載は×）<br>"
	sHTML = sHTML & "　＜理由例＞３５歳未満の方を募集（職務経験不問）<br>"
	sHTML = sHTML & "※上記のケース以外の場合は、年齢制限が認められないと疑った方が良いです。<br>"
	sHTML = sHTML & "</p>"
	sHTML = sHTML & "<p style=""margin-bottom:5px;"">"
	sHTML = sHTML & "【良くあるダメなケース】<br>"
	sHTML = sHTML & "・<span style=""color:#ff0000;"">有期労働契約（契約社員など）</span>の場合はほとんどの場合年齢制限ができません。<br>"
	If vAgeMin <> "" And vAgeMax <> "" Then
		sHTML = sHTML & "・<span style=""color:#ff0000;"">年齢の上限・下限両方</span>の記載があり、<span style=""color:#ff0000;"">30歳～49歳の間に納まっていない</span>場合、ほとんどのケースでできません。<br>"
		sHTML = sHTML & "　※30歳～49歳に納まっていても「３号のロ」に適合していない場合はアウトです。<br>"
	End If
	If vAgeMin <> "" Then
		sHTML = sHTML & "・<span style=""color:#ff0000;"">年齢の下限</span>がある場合は、ほとんどのケースでできません。<br>"
		sHTML = sHTML & "　※例外...労働基準法などにより年齢制限がある仕事の場合（１８歳以上など）<br>"
	End If
	If vAgeMax <> "" Then
		sHTML = sHTML & "・<span style=""color:#ff0000;"">年齢の上限</span>がある場合で<span style=""color:#ff0000;"">経験者を優遇</span>している場合は、ほとんどのケースでできません。<br>"
		sHTML = sHTML & "　※<span style=""color:#ff0000;"">実務経験が必要な資格</span>の所有者を優遇している場合もダメです。<br>"
		sHTML = sHTML & "　※例外...正社員募集かつ、職業経験不問かつ、新規学卒者と同等の処遇であること。<br>"
	End If
	sHTML = sHTML & "</p>"

	htmlAgeAdvice = sHTML
End Function
%>
