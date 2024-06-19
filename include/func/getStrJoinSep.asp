<%
'*******************************************************************************
'概　要：セパレート文字を追加して文字列を結合
'引　数：vStr1	：結合される文字列
'　　　：vStr2	：結合する文字列
'　　　：vSep	：結合される文字列が空文字で無い場合に、結合する区切り文字(セパレート)
'戻り値：String
'備　考：
'履　歴：2011/02/28 LIS K.Kokubo 作成
'*******************************************************************************
Function getStrJoinSep(ByVal vStr1,ByVal vStr2,ByVal vSep)
	getStrJoinSep = vStr1

	If getStrJoinSep <> "" Then getStrJoinSep = getStrJoinSep & vSep
	getStrJoinSep = getStrJoinSep & vStr2
End Function
%>
