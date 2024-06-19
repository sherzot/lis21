<%
'****************************************************
'*	関数名：FormatDate(簡易版)						*
'*	処理　：日付の形式を変換する					*
'*	引数　：YMD  形式日付							*
'*	　　　　MODE 0:0づめを行わない					*
'*	　　　　     1:0づめを行う						*
'*	　　　　SEP  セパレータ							*
'*	戻り値：形式が変換された日付					*
'****************************************************
function FormatDate( YMD, MODE, SEP )
	Dim wYYYY
	Dim wMM
	Dim wDD
	
	FormatDate = ""
	
	if Not isDate(YMD) then
		exit function
	end if
	
	if SEP = "" then
		SEP = "/"
	end if
	
	wYYYY = Year(YMD)
	if MODE = 1 then
		wMM = Right("00" & Month(YMD),2)
		wDD = Right("00" & Day(YMD),2)
	else
		wMM = Month(YMD)
		wDD = Day(YMD)
	end if
	
	FormatDate = wYYYY & SEP & wMM & SEP & wDD
end function
%>