<%
'****************************************************
'*	関数名：FormatCanma								*
'*	処理　：金額のカンマ編集						*
'*	引数　：Kingaku  金額							*
'*	戻り値：カンマ編集された金額					*
'****************************************************
function FormatCanma(Kingaku)
	dim wKin
	dim wLen
	dim wMainasu

	wKin = Kingaku

	if isnull(wKin) = true then
		wKin = 0
	end if

	if wKin < "0" and wKin <> "" then
		wMainasu = true
		wKin = wKin * -1
	else
		wMainasu = false
	end if

	select case len(wKin)
	case 0,1,2,3
		wKin = wKin
	case 4,5,6
		wKin = left(wKin,len(wKin) - 3) & "," & right(wKin,3)
	case 7,8,9
		wKin = left(wKin,len(wKin) - 6) & "," & mid(wKin,len(wKin) - 6 + 1,3) & "," & right(wKin,3)
	case else
		wKin = left(wKin,len(wKin) - 9) & "," &  mid(wKin,len(wKin) - 9 + 1,3) & "," & mid(wKin,len(wKin) - 6 + 1,3) & "," & right(wKin,3)
	end select

	if wMainasu = true then
		wKin = "-" & wKin
	end if

	FormatCanma = wKin

end function
%>
