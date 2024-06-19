<%

On Error Resume Next

Function decode(ByVal url)
	Dim sjis
	Dim error
	Dim re
	Dim fileProtocol
	Set re=New RegExp
	If IsEmpty(url) Then Exit Function
	fileProtocol = InStr(UCase(url),"file:")
	re.IgnoreCase=True
	re.Pattern="UTF-8|EUC|SHIFT_JIS|GOOGLE|MSN|BIGLOBE|VOD4ALL"
	If re.Test(url) Then
		Select Case re.Execute(UCase(url))(0).Value
		Case "EUC" sjis=decodeEUC(url)
		Case "UTF-8" sjis=decodeUTF8(url)
		Case "SHIFT_JIS" sjis=decodeSJIS(url)
		Case "GOOGLE" sjis=decodeUTF8(url)
		Case "MSN" sjis=decodeUTF8(url)
		Case "BIGLOBE" sjis=decodeEUC(url)		
		Case "VOD4ALL" sjis=decodeUTF8(url)		'http://vod4all.biz/search.php?q=
		End Select
	End If
	If sjis="" Then sjis=decodeEUC(url)
	'If sjis="" Then sjis=decodeSJIS(url)
	'If sjis="" Then sjis=decodeUTF8(url)
	'If sjis="" Then sjis=unescape(url)
	'If sjis<>"" Then sjis=Replace(sjis,"&amp;","&")
	'If sjis="" Then sjis=error
	decode = sjis
	
End Function

Function decodeSJIS(ByVal src)
	Dim dst,k,char,asc1,asc2
	src=UnEscape(src)
	For k=1 To Len(src)
  	char=Mid(src,k,1)
  	Asc1=AscW(char)
  	If fileProtocol Then If Asc1=43 Then Asc1=32
  	If (&H00 <= Asc1 And Asc1 <= &H80) Or _
     	(&HA0 <= Asc1 And Asc1 <= &HDF) Then
    	char=Chr(Asc1)
  	ElseIf (&H81 <= Asc1 And Asc1 <= &H9F) Or _
         	(&HE0 <= Asc1 And Asc1 <= &HFF) Then
    	k=k + 1
    	char=Mid(src,k,1)
    	Asc2=AscW(char)
    	If Asc2<64 Or Asc2>252 Then
      	error="Invalid SJIS second byte : " & Right(Hex(Asc2+256),2)
      	Exit Function
    	End If
    	char=Asc1*256+Asc2
    	If (Asc(chr(char)) And &HFFFF&)=char Then
       	char=Chr(char)
    	Else
      	char="%s" & Right(Hex(char+256*256),4)
      	error="Invalid SJIS code : " & char
      	Exit Function
    	End If
  	Else
	    error="Invalid SJIS first byte : " & Right(Hex(Asc1+256),2)
	    Exit Function
  	End If
  	dst=dst & char
	Next
	decodeSJIS=dst
End Function

Function decodeEUC(ByVal buf)
	Dim code
	Dim code1
	Dim code2
	Dim k
	Dim euc
	
	buf=UnEscape(buf)
	For k=1 To Len(buf)
  	code=AscW(Mid(buf,k,1))
  	
  	If fileProtocol Then If code=43 Then code=32
  	If code<&H80 Then
	    euc=euc & Chr(code)
  	ElseIf code=&H8E Then
	    k=k+1
	    code=AscW(Mid(buf,k,1))
	    euc=euc & Chr(code)
  	Else
	    code1=code
	    k=k+1
	    code2=AscW(Mid(buf,k,1))
	    code=jis2sjis((code1 And &H7F) * 256 + (code2 And &H7F))
	    If code Then
      	euc=euc & Chr(code)
    	Else
      	euc=euc & "%" & Right(Hex(code1+256),2) & "%" & Right(Hex(code2+256),2)
      	error="Invalid EUC code : " & Right(Hex(code1+256),2) & Right(Hex(code2+256),2)
      	Exit Function
    	End If
  	End If
	Next
	
	decodeEUC=euc
	
	
End Function

Function jis2sjis(j)
	Dim s
	Dim a
	Dim b
	
	a=j \ 256
	b=j mod 256
	If 33<=a And a<= 94 Then
  	If a mod 2 Then
	    If 33<=b And b<= 95 Then
      	s=a*128+b+28831
    	ElseIf 96<=b And b<=126 Then
      	s=a*128+b+28832
    	Else
      	s=0
    	End if
  	Else
	    If 33<=b And b<=126 Then
      	s=a*128+b+28798
    	Else
      	s=0
    	End If
  	End If
	ElseIf 95<=a And a<=126 Then
  	If a mod 2 Then
	    If 33<=b And b<= 95 Then
      	s=a*128+b+45215
    	ElseIf 96<=b And b<=126 Then
      	s=a*128+b+45216
    	Else
      	s=0
    	End If
  	ElseIf 33<=b And b<=126 Then
	    s=a*128+b+45182
  	Else
	    s=0
  	End If
	Else
  	s=0
	End If
	If s=0 Then
  	error="Invalid JIS Code : " & Right(Hex(a+256),2) & Right(Hex(b+256),2)
	End If
	jis2sjis=s
End Function

Function decodeUTF8(ByVal buf)
	Dim code
	Dim code1
	Dim code2
	Dim code3
	Dim k
	Dim sjis
	
	buf=UnEscape(buf)
	For k=1 To Len(buf)
  	code=AscW(Mid(buf,k,1))
  	If fileProtocol Then If code=43 Then code=32
  	If 0<=code And code<128 Then
	    sjis=sjis & ChrW(code)
  	ElseIf 128+64<=code And code<128+64+32 Then
	    code1=code
	    k=k+1
	    code2=AscW(Mid(buf,k,1))
	    code=(code1 And &H1F) * 64 + (code2 And &H3F)
	    If Chr(Asc(ChrW(code)))=ChrW(code) Then
      	sjis=sjis & ChrW(code)
    	Else
      	sjis=sjis & Escape(ChrW(code))
      	error="Invalid UTF-8 Code : " & Escape(ChrW(code)) & " " & Right(Hex(code1+256),2) & Right(Hex(code2+256),2)
    	End If
  	ElseIf 128+64+32<=code And code<128+64+32+16 Then
	    code1=code
	    k=k+1
	    code2=AscW(Mid(buf,k,1))
	    k=k+1
	    code3=AscW(Mid(buf,k,1))
	    code=(code1 And &H0F) * 16 * 256 + (code2 And &H3F) * 64 + (code3 And &H3F)
	    If Chr(Asc(ChrW(code)))=ChrW(code) Then
      	sjis=sjis & ChrW(code)
    	Else
      	sjis=sjis & Escape(ChrW(code))
      	error="Invalid UTF-8 Code : " & Escape(ChrW(code)) & " " & Right(Hex(code1+256),2) & Right(Hex(code2+256),2) & Right(Hex(code3+256),2)
    	End If
  	Else
	    sjis=sjis & Escape(ChrW(code))
	    error="Invalid UTF-8 Code : " & Escape(ChrW(code)) & " " & Right(Hex(code+256),2)
  	End If
	Next
	decodeUTF8=sjis
End Function

Function Fold(src)
	Dim k
	Dim dst
	Dim char
	Dim col,cols
	For k=1 To Len(src)
		char=Mid(src,k,1)
		Select Case Asc(char) And &HFF00&
		Case 0 col=1
		Case Else col=2
		End Select
		cols=cols+col
		If cols>39 Then
		    dst=dst & vbCRLF & char
		    cols=col
		Else
	    dst=dst & char
  		End If
	Next
	Fold=dst
End Function



'***************************************************************************
'***
'***  ここからプログラム開始
'***
'***************************************************************************

On Error Resume Next
	
	Dim VisitorCode			:	VisitorCode = Request.Cookies("VisitorCode")
	Dim FirstVisitDay
	Dim LastVisitDay
	Dim FirstVisitPage
	Dim LastVisitPage
	Dim SearchArea
	Dim SearchPrefecture
	Dim SearchJobType
	Dim SearchWorkingType
	Dim SearchWord
	Dim OriginAccessPage
	Dim AccessCount
	Dim RegistCode			:	RegistCode = Session("userid") 'Request.Cookies("id_memory")

	
	Dim vSQL
	Dim vRc,vRc2
	Dim bobj
	Dim URLString
	Dim Prm,URLPrm
	Dim PrmCount,PrmCheck
	Dim SQLSearchWord
	Session("AccessCount") = Session("AccessCount") + 1
	
	Set bobj = Server.CreateObject("BASP21")
	'************************************************************
	'*** ■再来訪者であるかを確認する■
	'*** クッキー変数：VisitorCode
	'*** 値：有（再来訪者）、無（初来訪者）
	'************************************************************
	
	vSQL = "Select * from LOG_Visitor where VisitorCode = '" & VisitorCode & "'"
	Set vRc = dbconn.Execute(vSQL)
	
	if Not vRc.EOF then
		vSQL = "Update LOG_Visitor Set "
		vSQL = vSQL & " LastVisitDay = GetDate()"
		vSQL = vSQL & ",LastVisitPage = '" &  Request.ServerVariables("url") & "'"
		
		'*******************************************************************************
		'■過去のアクセス回数以上のトラッフィックを検出したら更新する。
		'*******************************************************************************
		if vRc.Collect("AccessCount") < Session("AccessCount") then
			vSQL = vSQL & ",AccessCount = '" & Session("AccessCount") & "'"
		End If
		
		'*******************************************************************************
		'■お仕事検索（こだわり検索）の利用があった場合はその内容を記録する。
		'*******************************************************************************
		if InStr(Request.ServerVariables("url"),Request.ServerVariables("url")) > 0 then
			if Request("CONF_JT2") <> "" then	vSQL = vSQL & ",SearchJobType = '" &  Request("CONF_JT2") & "'"
			if Request("CONF_AC") <> "" then	vSQL = vSQL & ",SearchArea = '" &  Request("CONF_AC") & "'"
			if Request("CONF_AC2") <> "" then	vSQL = vSQL & ",SearchPrefecture = '" &  Request("CONF_AC2") & "'"
			if Request("CONF_WT") <> "" then	vSQL = vSQL & ",SearchWorkingType = '" &  Request("CONF_WT") & "'"
		End If
		'*******************************************************************************
		'■リファラに検索文字が含まれていれば検索文字列を探す。
		'*******************************************************************************

		if Request.ServerVariables("HTTP_REFERER") <> "" and InStr(Request.ServerVariables("HTTP_REFERER"),"?") > 0 then
			'*** パラメータ付きのＵＲＬを分解
			
			URLString = Split(decode(Request.ServerVariables("HTTP_REFERER")), "?")
			

response.write Err.Number
			'*** 既に検索検索文字列を持っていれば、LastSearchWordの方に記録する。
			if vRC.Collect("SearchWord") <> "" then
				SQLSearchWord = ",LastSearchWord"
			Else
				SQLSearchWord = ",SearchWord"
			End If
			
			
			
			'*** 複数のパラメータが存在するか確認
			if InStr(URLString(1),"&") > 0 then
				URLPrm = Split(URLString(1), "&")
				for PrmCount=0 to UBound(URLPrm)
					Prm = Split(URLPrm(PrmCount),"=")
					
					if Prm(0) = "q" or Prm(0) = "p" or Prm(0) = "MT" then
						vSQL = vSQL & SQLSearchWord & " = '" & Prm(1) & "'"
						vSQL = vSQL & ",OriginAccessPage = '" & Server.HTMLEncode(Request.ServerVariables("HTTP_REFERER")) & "'"
						Exit for
					Elseif InStr(Request.ServerVariables("HTTP_REFERER"),"search.www.infoseek.co.jp") > 0 and Prm(0) = "qt" then
						vSQL = vSQL & SQLSearchWord & " = '" & Prm(1) & "'"
						vSQL = vSQL & ",OriginAccessPage = '" & Server.HTMLEncode(Request.ServerVariables("HTTP_REFERER")) & "'"
						Exit for
					Elseif InStr(decode(Request.ServerVariables("HTTP_REFERER")),"search.jp.aol.com") > 0 and Prm(0) = "query" then
						vSQL = vSQL & SQLSearchWord & " = '" & Prm(1) & "'"
						vSQL = vSQL & ",OriginAccessPage = '" & Server.HTMLEncode(Request.ServerVariables("HTTP_REFERER")) & "'"
						Exit for
					End If
					
				next
			else
				Prm = Split(URLString(1),"=")
				if Prm(0) = "q" or Prm(0) = "p" or Prm(0) = "MT" then
					vSQL = vSQL & SQLSearchWord & " = '" & Prm(1) & "'"
					vSQL = vSQL & ",OriginAccessPage = '" & Server.HTMLEncode(Request.ServerVariables("HTTP_REFERER")) & "'"

				Elseif InStr(Request.ServerVariables("HTTP_REFERER"),"search.www.infoseek.co.jp") > 0 and Prm(0) = "qt" then
					vSQL = vSQL & SQLSearchWord & " = '" & Prm(1) & "'"
					vSQL = vSQL & ",OriginAccessPage = '" & Server.HTMLEncode(Request.ServerVariables("HTTP_REFERER")) & "'"

				Elseif InStr(decode(Request.ServerVariables("HTTP_REFERER")),"search.jp.aol.com") > 0 and Prm(0) = "query" then
					vSQL = vSQL & SQLSearchWord & " = '" & Prm(1) & "'"
					vSQL = vSQL & ",OriginAccessPage = '" & Server.HTMLEncode(Request.ServerVariables("HTTP_REFERER")) & "'"

				End If
			End If
		Elseif Len(vRC.Collect("OriginAccessPage")) = 0 then
			vSQL = vSQL & ",OriginAccessPage = '" & Server.HTMLEncode(Request.ServerVariables("HTTP_REFERER")) & "'"
			
		End If
		
		'*******************************************************************************
		'■しごとナビに登録された場合はスタッフコードをアップデートする
		'*******************************************************************************
		if RegistCode <> "" then
			vSQL = vSQL & ",StaffCode = '" & RegistCode & "'"
		End If
		vSQL = vSQL & " Where VisitorCode = '" & VisitorCode & "'"
		Set vRc2 = dbconn.Execute(vSQL)
		Set vRc2 = Nothing
		
		vRc.Close
	Elseif Session("AccessCount") = 1 then
		
		'初来訪者に対して、来訪者コードを発番する。
		VisitorCode = Year(Date)
		if Len(Month(Date)) = 2 then
			VisitorCode = VisitorCode & Month(Date)
		Else
			VisitorCode = VisitorCode & "0" & Month(Date)
		End If
		if Len(Day(Date)) = 2 then
			VisitorCode = VisitorCode & Day(Date)
		Else
			VisitorCode = VisitorCode & "0" & Day(Date)
		End If
		VisitorCode = VisitorCode & Session.SessionID
		Response.Cookies("VisitorCode") = VisitorCode
		Response.Cookies("VisitorCode").Expires = Date + 365
		
		vSQL = "INSERT INTO LOG_visitor"
        vSQL = vSQL & "(VisitorCode"
        vSQL = vSQL & ",FirstVisitDay"
        vSQL = vSQL & ",LastVisitDay"
        vSQL = vSQL & ",FirstVisitPage"
        vSQL = vSQL & ",LastVisitPage"
        vSQL = vSQL & ",SearchWord"
        vSQL = vSQL & ",OriginAccessPage"
        vSQL = vSQL & ",AccessCount"
        vSQL = vSQL & ") VALUES"
        '訪問者コード
        vSQL = vSQL & "('" & VisitorCode & "'"								'VisitorCode
        vSQL = vSQL & ",GetDate()"											'FirstVisitDay
        vSQL = vSQL & ",GetDate()"											'LastVisitDay
        vSQL = vSQL & ",'" & Request.ServerVariables("url") & "'"			'FistVisitPage
        vSQL = vSQL & ",'" & Request.ServerVariables("url") & "'"			'LastVisitPage
        '**********************************************
        '** リファラより検索ワードが存在すれば記述する
        '**********************************************
		if InStr(Request.ServerVariables("HTTP_REFERER"),"?") > 0 then
			URLString = Split(decode(Request.ServerVariables("HTTP_REFERER")), "?")
			PrmCheck = 0
			if InStr(URLString(1),"&") > 0 then
				URLPrm = Split(URLString(1), "&")
				for PrmCount=0 to UBound(URLPrm)
					Prm = Split(URLPrm(PrmCount),"=")
					if Prm(0) = "q" _
						and InStr(decode(Request.ServerVariables("HTTP_REFERER")),"oshiete1.goo.ne.jp") = 0 _
						and InStr(decode(Request.ServerVariables("HTTP_REFERER")),"question.woman.excite.co.jp") = 0 then
						
						vSQL = vSQL & ",'" & Prm(1) & "'"
						PrmCheck = 1
						Exit for
					
					Elseif Prm(0) = "p" or Prm(0) = "MT" then
						vSQL = vSQL & ",'" & Prm(1) & "'"
						PrmCheck = 1
						Exit for
					Elseif InStr(decode(Request.ServerVariables("HTTP_REFERER")),"search.www.infoseek.co.jp") > 0 and Prm(0) = "qt" then
						vSQL = vSQL & ",'" & Prm(1) & "'"
						PrmCheck = 1
						Exit for
					Elseif InStr(decode(Request.ServerVariables("HTTP_REFERER")),"search.jp.aol.com") > 0 and Prm(0) = "query" then
						vSQL = vSQL & ",'" & Prm(1) & "'"
						PrmCheck = 1
						Exit for
					End If
				next
				if PrmCheck = 0 then vSQL = vSQL & ",null"
				
			Else
				Prm = Split(URLString(1),"=")
				if Prm(0) = "q" _
					and InStr(decode(Request.ServerVariables("HTTP_REFERER")),"oshiete1.goo.ne.jp") = 0 _
					and InStr(decode(Request.ServerVariables("HTTP_REFERER")),"question.woman.excite.co.jp") = 0 then
					
					vSQL = vSQL & ",'" & Prm(1) & "'"
					
				Elseif Prm(0) = "p" or Prm(0) = "MT" then
					vSQL = vSQL & ",'" & Prm(1) & "'"
				Elseif InStr(decode(Request.ServerVariables("HTTP_REFERER")),"search.www.infoseek.co.jp") > 0 and Prm(0) = "qt" then
					vSQL = vSQL & ",'" & Prm(1) & "'"
				Elseif InStr(decode(Request.ServerVariables("HTTP_REFERER")),"search.jp.aol.com") > 0 and Prm(0) = "query" then
					vSQL = vSQL & ",'" & Prm(1) & "'"
				Else
					vSQL = vSQL & ",null"
				End If
			End If
		Else
			vSQL = vSQL & ",null"
		End If
		if InStr(decode(Request.ServerVariables("HTTP_REFERER")),"shigotonavi") > 0 then
	        vSQL = vSQL & ",null"	'OriginAccessPage
		Elseif Trim(Request.ServerVariables("HTTP_REFERER")) <> "" then
	        vSQL = vSQL & ",'" & Server.HTMLEncode(Request.ServerVariables("HTTP_REFERER")) & "'"	'OriginAccessPage
		Else
	        vSQL = vSQL & ",null"	'OriginAccessPage
		End If
        vSQL = vSQL & ",'" & Session("AccessCount") & "'"					'AccessCount
        vSQL = vSQL & ")"
        
        If LEN(Request.Cookies("VisitorCode")) > 8 then Set vRc = dbconn.Execute(vSQL)
        Set vRc = Nothing
	End If

'***************************************************************************
'***
'***  ここでプログラム終了
'***
'***************************************************************************

%>
